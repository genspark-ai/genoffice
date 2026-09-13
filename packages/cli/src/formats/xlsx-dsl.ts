import {
  columnIndex,
  parseAddress,
  parseRange,
  rangeAddresses,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import { InMemoryWorkbookAdapter } from '@genoffice/xlsx-gateway/domain/in-memory-workbook'
import type { ChangePlan, WorkbookSnapshot } from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { CellFormatPatch } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import {
  readBasicWorkbook,
  type CellEdit,
  type SheetStructuralOps,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import {
  parseRelationships,
  parseSheetElements,
  type SheetEditPlan,
} from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import type { StructuralOp } from '@genoffice/xlsx-gateway/gateway/xlsx-structure'
import type { DefinedNameEntry } from '@genoffice/xlsx-gateway/gateway/xlsx-defined-names'
import type { SheetNote } from '@genoffice/xlsx-gateway/gateway/xlsx-notes'
import type { WorkbookStyleEdit } from '@genoffice/xlsx-gateway/shared/edit-schemas'
import {
  isLayoutOp,
  layoutOpLabel,
  structuralOpLabel,
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import JSZip from 'jszip'
import type { PathContext } from '../fs'
import { CliError, EXIT } from '../result'
import {
  buildGatewayPayloads,
  GATEWAY_DSL_OPS,
  REFUSED_DSL_OPS,
  type GatewayPayloads,
  type SheetFileState,
  type SheetTabOps,
  type WorkbookFileState,
} from './xlsx-gateway-ops'

/**
 * The sheets AI DSL, headless. The app validates and expands a batch with the
 * workbook DSL, runs it on Univer and saves through the gateway from Univer's
 * journal. Here the in-memory workbook adapter plays Univer's part: it
 * validates the batch, applies it to a snapshot read straight from the file,
 * and the difference between the snapshots becomes the gateway's edits. Ops
 * the adapter validates but cannot hold (links, filters, rules, visuals, tab
 * state) become the gateway's declarative payloads (xlsx-gateway-ops.ts).
 */
export const ADAPTER_DSL_OPS = [
  'set_cell',
  'set_formula',
  'set_range',
  'clear_cell',
  'clear_range',
  'fill_range',
  'copy_range',
  'find_replace',
  'sort_range',
  'format_range',
  'merge_cells',
  'unmerge_cells',
  'set_row_height',
  'set_col_width',
  'insert_rows',
  'delete_rows',
  'insert_cols',
  'delete_cols',
  'add_sheet',
  'delete_sheet',
  'rename_sheet',
] as const

export const SUPPORTED_DSL_OPS = [...ADAPTER_DSL_OPS, ...GATEWAY_DSL_OPS] as const

/** Gateway ops the adapter refuses to apply to its snapshot (validation still runs through plan()). */
const ADAPTER_SKIPS = new Set<string>(GATEWAY_DSL_OPS.filter((op) => op !== 'add_chart'))
/** Ops that touch no cell address: they may ride in a structural batch. */
const TAB_ONLY = new Set([
  'rename_sheet',
  'protect_sheet',
  'set_sheet_hidden',
  'move_sheet',
  'add_defined_name',
  'delete_defined_name',
  'set_page_setup',
  'set_freeze',
  'clear_filter',
  'clear_conditional_formats',
  'edit_chart',
])
const SHEET_TAB = new Set(['rename_sheet', 'set_sheet_hidden', 'move_sheet'])
/** Layout ops the gateway saves as row/column changes of their sheet. */
const ROW_COL = new Set([
  'merge_cells',
  'unmerge_cells',
  'set_row_height',
  'set_col_width',
  'set_rows_hidden',
  'set_cols_hidden',
])
/** Ops without a worksheet of their own. */
const NO_SHEET = new Set(['add_sheet', 'add_defined_name', 'delete_defined_name', 'edit_chart'])

const STRUCTURAL = new Set([
  'insert_rows',
  'delete_rows',
  'insert_cols',
  'delete_cols',
  'add_sheet',
  'delete_sheet',
  // the copy is cloned from the file's part, so same-batch cell edits would be missing from it
  'duplicate_sheet',
])
/** OOXML column widths are in characters of the default font; 7 px per character is Calibri 11. */
const PX_PER_CHAR = 7

export interface DslOutcome {
  edits: CellEdit[]
  /** original sheet name → new name, for callers that touch the written file afterwards */
  renames: Record<string, string>
  structuralOps: SheetStructuralOps[]
  sheetPlan?: SheetEditPlan
  gateway: GatewayPayloads
  plan: string[]
  warnings: string[]
}

export async function runWorkbookDsl(
  source: Buffer,
  rawOps: unknown[],
  defaultSheet?: string,
  ctx?: PathContext,
): Promise<DslOutcome> {
  const imported = await readBasicWorkbook(source)
  const meta = await readWorkbookMeta(source)
  // readBasicWorkbook stops at cell values; the adapter's merge-overlap check needs the file's merges
  const snapshot: WorkbookSnapshot = {
    ...imported.snapshot,
    sheets: imported.snapshot.sheets.map((sheet) => {
      const merges = meta.merges.get(sheet.name)
      return merges?.length ? { ...sheet, merges } : sheet
    }),
  }
  const idByName = new Map(Object.entries(imported.sheetNamesById).map(([id, name]) => [name, id]))
  const names = [...idByName.keys()]
  if (defaultSheet !== undefined && !idByName.has(defaultSheet)) {
    throw new CliError(EXIT.usage, `sheet not found: ${defaultSheet}`, { sheets: names })
  }
  const active =
    meta.activeSheet !== null && idByName.has(meta.activeSheet) ? meta.activeSheet : null
  const fallback = defaultSheet ?? active ?? snapshot.sheets[0]?.name
  const fallbackId = fallback ? idByName.get(fallback) : undefined
  const ops = rawOps.map((op, i) => normalizeOp(op, i, idByName, fallbackId))
  // the planner used to see the whole batch and refuse this mix; planning per op below would not
  const structuralOps0 = ops.filter((op) => STRUCTURAL.has(op.op))
  const contentOps = ops.filter((op) => !STRUCTURAL.has(op.op) && !TAB_ONLY.has(op.op))
  if (structuralOps0.length && contentOps.length) {
    throw new CliError(
      EXIT.usage,
      `ops rejected: structural ops (${[...new Set(structuralOps0.map((o) => o.op))].join(', ')}) shift addresses and cannot share a batch with content ops; run them as two batches`,
    )
  }
  // the writer refuses name edits next to sheet or row/column changes (sheet indexes move)
  const nameOps = ops.filter(
    (op) => op.op === 'add_defined_name' || op.op === 'delete_defined_name',
  )
  const sheetOps = ops.filter(
    (op) => STRUCTURAL.has(op.op) || SHEET_TAB.has(op.op) || ROW_COL.has(op.op),
  )
  if (nameOps.length && sheetOps.length) {
    throw new CliError(
      EXIT.usage,
      `ops rejected: defined-name ops cannot share a batch with sheet or row/column ops (${[...new Set(sheetOps.map((o) => o.op))].join(', ')}); run them as two batches`,
    )
  }
  // and a new table next to row/column changes on its own sheet
  for (const table of ops.filter((op) => op.op === 'add_table')) {
    const clash = ops.find(
      (op) =>
        op !== table &&
        (STRUCTURAL.has(op.op) || ROW_COL.has(op.op)) &&
        (op.sheetId === undefined || op.sheetId === table.sheetId),
    )
    if (clash) {
      throw new CliError(
        EXIT.usage,
        `ops rejected: add_table cannot share a batch with row/column, merge or size ops on the same sheet (${clash.op}); run them as two batches`,
      )
    }
  }
  const adapter = new InMemoryWorkbookAdapter(snapshot)
  const before = adapter.getSnapshot()
  // one plan per op, applied before the next is planned: sort_range / find_replace
  // expand against the workbook as the previous ops left it, not the file's snapshot
  const plans: ChangePlan[] = []
  const labels: string[] = []
  for (const [index, op] of ops.entries()) {
    // the adapter's plan() rehearses every op on its snapshot, which has nowhere
    // to keep links, rules or visuals: gateway ops are validated by schema only
    if (ADAPTER_SKIPS.has(op.op)) {
      const parsed = workbookOperationSchema.safeParse(op)
      if (!parsed.success) {
        throw new CliError(
          EXIT.usage,
          `ops[${index}] (${op.op}) rejected: ${describeError(parsed.error)}`,
        )
      }
      Object.assign(op, parsed.data)
      const typed = parsed.data
      const label = isLayoutOp(typed) ? layoutOpLabel(typed) : structuralOpLabel(typed as never)
      const sheetId = 'sheetId' in typed ? typed.sheetId : undefined
      labels.push(sheetId ? label.replaceAll(sheetId, idByName_(idByName, sheetId)) : label)
      continue
    }
    let plan: ChangePlan
    try {
      plan = adapter.plan({
        dslVersion: 1,
        transactionId: `genoffice-${Date.now()}-${index}`,
        baseRevision: adapter.getSnapshot().revision,
        summary: 'genoffice sheet apply',
        operations: [op],
      })
    } catch (err) {
      throw new CliError(EXIT.usage, `ops[${index}] (${op.op}) rejected: ${describeError(err)}`, {
        supported: [...SUPPORTED_DSL_OPS],
      })
    }
    try {
      adapter.apply(plan)
    } catch (err) {
      throw new CliError(EXIT.usage, `ops[${index}] (${op.op}) rejected: ${describeError(err)}`)
    }
    plans.push(plan)
    labels.push(
      ...plan.structuralChanges.map((c) => c.label),
      ...plan.formatChanges.map((c) => c.label),
      ...plan.sheetRenames.map((r) => `rename sheet ${r.before} → ${r.after}`),
      ...(plan.cellChanges.length ? [`${plan.cellChanges.length} cell change(s)`] : []),
    )
  }
  const plan = mergePlans(plans)
  const after = adapter.getSnapshot()
  // the gateway addresses sheets by the names in the file; renames ride the sheet plan
  const beforeNames = new Map(before.sheets.map((s) => [s.id, s.name]))
  const namesById = new Map(after.sheets.map((s) => [s.id, beforeNames.get(s.id) ?? s.name]))
  const structural = ops.some((op) => STRUCTURAL.has(op.op))

  const edits = new Map<string, CellEdit>()
  const key = (e: CellEdit) => `${e.sheetName}|${e.row}|${e.column}`
  // structural ops shift addresses inside the gateway itself, so cell diffs would
  // double-apply the shift; the DSL already forbids mixing them with content ops
  if (!structural) {
    for (const e of diffCells(before, after, namesById)) edits.set(key(e), e)
  }
  for (const change of plan.formatChanges) {
    const sheetName = namesById.get(change.sheetId) ?? change.sheetId
    const style = patchToStyleEdit(change.format)
    for (const address of rangeAddresses(parseRange(change.range))) {
      const { row, column } = parseAddress(address)
      const k = `${sheetName}|${row}|${column}`
      const existing = edits.get(k)
      edits.set(
        k,
        existing
          ? { ...existing, style: { ...existing.style, ...style } }
          : { sheetName, row, column, writeValue: false, cell: { value: null }, style },
      )
    }
  }

  const structuralOps = new Map<string, StructuralOp[]>()
  const push = (sheetName: string, op: StructuralOp) =>
    structuralOps.set(sheetName, [...(structuralOps.get(sheetName) ?? []), op])
  for (const { op } of plan.structuralChanges) {
    const sheetName =
      'sheetId' in op && typeof op.sheetId === 'string'
        ? (namesById.get(op.sheetId) ?? imported.sheetNamesById[op.sheetId] ?? op.sheetId)
        : ''
    switch (op.op) {
      case 'insert_rows':
        push(sheetName, { kind: 'insert-rows', index: op.row - 1, count: op.count })
        break
      case 'delete_rows':
        push(sheetName, { kind: 'remove-rows', index: op.row - 1, count: op.count })
        break
      case 'insert_cols':
        push(sheetName, { kind: 'insert-cols', index: columnIndex(op.column), count: op.count })
        break
      case 'delete_cols':
        push(sheetName, { kind: 'remove-cols', index: columnIndex(op.column), count: op.count })
        break
      case 'merge_cells':
        push(sheetName, { kind: 'merge-cells', range: parseRange(op.range) })
        break
      case 'unmerge_cells':
        push(sheetName, { kind: 'unmerge-cells', range: parseRange(op.range) })
        break
      case 'set_row_height':
        push(sheetName, {
          kind: 'set-row-size',
          start: op.row - 1,
          end: op.row - 1 + op.count - 1,
          size: op.heightPoints,
        })
        break
      case 'set_col_width': {
        const start = columnIndex(op.column)
        push(sheetName, {
          kind: 'set-col-size',
          start,
          end: start + op.count - 1,
          size: Math.max(Math.round((op.widthPx / PX_PER_CHAR) * 256) / 256, 1 / 256),
        })
        break
      }
      default:
        // sheet ops land in the sheet plan below; range-level content ops are
        // already covered by the cell diff
        break
    }
  }

  const gateway = await buildGatewayPayloads({
    ops: ops.map((op, i) => ({ ...(op as unknown as WorkbookOperation), __index: i })),
    namesById,
    after,
    file: meta,
    ctx,
  })
  for (const { sheetName, ops: hiddenOps } of gateway.hiddenOps) {
    for (const op of hiddenOps) push(sheetName, op)
  }

  const sheetPlan = buildSheetPlan(before, after, plan, gateway.tabs, meta.order)
  return {
    edits: [...edits.values()],
    renames: Object.fromEntries(plan.sheetRenames.map((r) => [r.before, r.after])),
    structuralOps: [...structuralOps].map(([sheetName, ops]) => ({ sheetName, ops })),
    ...(sheetPlan ? { sheetPlan } : {}),
    gateway,
    plan: labels,
    warnings: [...plan.warnings, ...gateway.warnings],
  }
}

function idByName_(idByName: Map<string, string>, id: string): string {
  return [...idByName].find(([, v]) => v === id)?.[0] ?? id
}

function mergePlans(plans: ChangePlan[]): ChangePlan {
  return {
    transactionId: plans[0]?.transactionId ?? '',
    baseRevision: plans[0]?.baseRevision ?? 0,
    cellChanges: plans.flatMap((p) => p.cellChanges),
    formatChanges: plans.flatMap((p) => p.formatChanges),
    structuralChanges: plans.flatMap((p) => p.structuralChanges),
    sheetRenames: plans.flatMap((p) => p.sheetRenames),
    warnings: plans.flatMap((p) => p.warnings),
  }
}

/** Accepts `sheet` (a name) as well as the DSL's `sheetId`; a missing sheet means the default one. */
function normalizeOp(
  raw: unknown,
  index: number,
  idByName: Map<string, string>,
  fallbackId: string | undefined,
): Record<string, unknown> & { op: string } {
  if (!raw || typeof raw !== 'object' || typeof (raw as { op?: unknown }).op !== 'string') {
    throw new CliError(EXIT.usage, `ops[${index}]: each op must be an object with an "op" name`)
  }
  const op = { ...(raw as Record<string, unknown>) } as Record<string, unknown> & { op: string }
  if (!(SUPPORTED_DSL_OPS as readonly string[]).includes(op.op)) {
    const reason = Object.hasOwn(REFUSED_DSL_OPS, op.op) ? REFUSED_DSL_OPS[op.op] : undefined
    throw new CliError(
      EXIT.usage,
      reason
        ? `ops[${index}]: "${op.op}" is not available headless (${reason}); use the GenOffice app`
        : `ops[${index}]: unknown op "${op.op}"`,
      { supported: [...SUPPORTED_DSL_OPS], not_available: Object.keys(REFUSED_DSL_OPS) },
    )
  }
  // same convention as set_range and --cells: a value starting with "=" is a formula
  if (op.op === 'set_cell' && typeof op.value === 'string' && op.value.startsWith('=')) {
    op.op = 'set_formula'
    op.formula = op.value
    delete op.value
  }
  if (NO_SHEET.has(op.op)) return op
  const wanted = typeof op.sheet === 'string' ? op.sheet : op.sheetId
  delete op.sheet
  if (typeof wanted === 'string') {
    const id =
      idByName.get(wanted) ?? ([...idByName.values()].includes(wanted) ? wanted : undefined)
    if (!id) {
      throw new CliError(EXIT.usage, `ops[${index}]: sheet not found: ${wanted}`, {
        sheets: [...idByName.keys()],
      })
    }
    op.sheetId = id
  } else {
    op.sheetId = fallbackId
  }
  return op
}

interface WorkbookMeta extends WorkbookFileState {
  activeSheet: string | null
  merges: Map<string, string[]>
}

/**
 * What readBasicWorkbook does not surface but the payloads need: the active
 * tab, merged ranges, and the per-sheet state the gateway rewrites as a whole
 * (rules, filter, notes) so existing content is carried or the op refused.
 */
async function readWorkbookMeta(source: Buffer): Promise<WorkbookMeta> {
  const zip = await JSZip.loadAsync(source)
  const workbook = (await zip.file('xl/workbook.xml')?.async('string')) ?? ''
  const rels = (await zip.file('xl/_rels/workbook.xml.rels')?.async('string')) ?? ''
  const targets = new Map(
    parseRelationships(rels)
      .filter((r) => r.id !== undefined && !r.external)
      .map((r) => [r.id!, partPath(r.target)]),
  )
  const sheets = parseSheetElements(workbook).map((sheet) => ({
    name: sheet.name,
    path: sheet.relationshipId ? targets.get(sheet.relationshipId) : undefined,
  }))
  const active = Number(/<workbookView\b[^>]*\bactiveTab="(\d+)"/.exec(workbook)?.[1] ?? 0)
  const merges = new Map<string, string[]>()
  const states = new Map<string, SheetFileState>()
  for (const sheet of sheets) {
    if (!sheet.path) continue
    const xml = await zip.file(sheet.path)?.async('string')
    if (!xml) continue
    const refs = [...xml.matchAll(/<mergeCell\b[^>]*\bref="([^"]+)"/g)].map((m) => m[1]!)
    if (refs.length) merges.set(sheet.name, refs)
    const filter = /<autoFilter\b[^>]*\bref="([^"]+)"[^>]*(\/>|>[\s\S]*?<\/autoFilter>)/.exec(xml)
    states.set(sheet.name, {
      conditionalFormats: (xml.match(/<conditionalFormatting\b/g) ?? []).length,
      dataValidations: (xml.match(/<dataValidation\b/g) ?? []).length,
      autoFilter: filter
        ? { range: parseRange(filter[1]!), hasCriteria: /<filterColumn\b/.test(filter[2]!) }
        : null,
      hiddenRows: [...xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*\bhidden="(?:1|true)"/g)].map(
        (m) => Number(m[1]) - 1,
      ),
      notes: await readNotes(zip, sheet.path),
    })
  }
  const tableNames: string[] = []
  for (const path of Object.keys(zip.files).filter((f) => /^xl\/tables\/[^/]+\.xml$/.test(f))) {
    const xml = await zip.file(path)!.async('string')
    const name = /<table\b[^>]*\b(?:displayName|name)="([^"]+)"/.exec(xml)?.[1]
    if (name) tableNames.push(unescapeXml(name))
  }
  return {
    activeSheet: sheets[active]?.name ?? sheets[0]?.name ?? null,
    merges,
    order: sheets.map((s) => s.name),
    definedNames: readDefinedNames(workbook),
    tableNames,
    sheets: states,
  }
}

/** The comments part of a worksheet as gateway notes (the whole set is rewritten on save). */
async function readNotes(zip: JSZip, sheetPath: string): Promise<SheetNote[]> {
  const relsPath = sheetPath.replace(/([^/]+)$/, '_rels/$1.rels')
  const rels = await zip.file(relsPath)?.async('string')
  if (!rels) return []
  const rel = parseRelationships(rels).find((r) => /\/comments$/.test(r.type))
  if (!rel) return []
  const dir = sheetPath.slice(0, sheetPath.lastIndexOf('/') + 1)
  const target = rel.target.startsWith('/')
    ? rel.target.slice(1)
    : normalizePath(`${dir}${rel.target}`)
  const xml = await zip.file(target)?.async('string')
  if (!xml) return []
  const authors = [...xml.matchAll(/<author>([\s\S]*?)<\/author>/g)].map((m) => unescapeXml(m[1]!))
  const notes: SheetNote[] = []
  for (const m of xml.matchAll(/<comment\b([^>]*)>([\s\S]*?)<\/comment>/g)) {
    const ref = /\bref="([^"]+)"/.exec(m[1]!)?.[1]
    if (!ref) continue
    const authorId = /\bauthorId="(\d+)"/.exec(m[1]!)?.[1]
    const { row, column } = parseAddress(ref)
    const text = [...m[2]!.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)]
      .map((t) => unescapeXml(t[1]!))
      .join('')
    notes.push({ row, column, author: authors[Number(authorId ?? 0)] ?? '', text })
  }
  return notes
}

/** Modeled names only: Excel's `_xlnm.*` built-ins and hidden names stay in the file untouched. */
function readDefinedNames(workbookXml: string): DefinedNameEntry[] {
  const out: DefinedNameEntry[] = []
  for (const m of workbookXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const attrs = m[1]!
    const name = unescapeXml(/\bname="([^"]*)"/.exec(attrs)?.[1] ?? '')
    if (!name || name.startsWith('_xlnm') || /\bhidden="(?:1|true)"/.test(attrs)) continue
    const local = /\blocalSheetId="(\d+)"/.exec(attrs)?.[1]
    out.push({
      name,
      formula: unescapeXml(m[2]!),
      ...(local === undefined ? {} : { sheetIndex: Number(local) }),
    })
  }
  return out
}

function normalizePath(path: string): string {
  const parts: string[] = []
  for (const seg of path.split('/')) {
    if (seg === '..') parts.pop()
    else if (seg !== '.' && seg !== '') parts.push(seg)
  }
  return parts.join('/')
}

function unescapeXml(text: string): string {
  return text
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&amp;', '&')
}

/** Relationship targets come absolute ("/xl/…"), package-relative ("xl/…") or relative to xl/. */
function partPath(target: string): string {
  const t = target.replace(/^\.\//, '')
  if (t.startsWith('/')) return t.slice(1)
  return t.startsWith('xl/') ? t : `xl/${t}`
}

function diffCells(
  before: WorkbookSnapshot,
  after: WorkbookSnapshot,
  namesById: Map<string, string>,
): CellEdit[] {
  const out: CellEdit[] = []
  const previous = new Map(before.sheets.map((s) => [s.id, s]))
  for (const sheet of after.sheets) {
    const sheetName = namesById.get(sheet.id) ?? sheet.name
    const prior = previous.get(sheet.id)
    const addresses = new Set([...Object.keys(sheet.cells), ...Object.keys(prior?.cells ?? {})])
    for (const address of addresses) {
      const a = sheet.cells[address]
      const b = prior?.cells[address]
      if (sameCell(a, b)) continue
      const { row, column } = parseAddress(address)
      out.push({
        sheetName,
        row,
        column,
        writeValue: true,
        cell: a
          ? { value: a.value, ...(a.formula ? { formula: a.formula } : {}) }
          : { value: null },
      })
    }
  }
  return out
}

function sameCell(
  a: { value: unknown; formula?: string | undefined } | undefined,
  b: { value: unknown; formula?: string | undefined } | undefined,
): boolean {
  const av = a?.value ?? null
  const bv = b?.value ?? null
  return av === bv && (a?.formula ?? null) === (b?.formula ?? null)
}

function buildSheetPlan(
  before: WorkbookSnapshot,
  after: WorkbookSnapshot,
  plan: ChangePlan,
  tabs: SheetTabOps,
  fileOrder: string[],
): SheetEditPlan | undefined {
  const beforeIds = new Set(before.sheets.map((s) => s.id))
  const afterIds = new Set(after.sheets.map((s) => s.id))
  const additions = [
    ...after.sheets.filter((s) => !beforeIds.has(s.id)).map((s) => ({ name: s.name })),
    ...tabs.duplicates.map((d) => ({ name: d.name, sourceSheetName: d.source })),
  ]
  const removals = before.sheets.filter((s) => !afterIds.has(s.id)).map((s) => s.name)
  const renames = plan.sheetRenames.map((r) => ({ sheetName: r.before, newName: r.after }))
  if (
    additions.length === 0 &&
    removals.length === 0 &&
    renames.length === 0 &&
    tabs.moves.length === 0 &&
    tabs.hidden.length === 0
  ) {
    return undefined
  }
  const renamed = new Map(renames.map((r) => [r.sheetName, r.newName]))
  const finalName = (original: string) => renamed.get(original) ?? original
  const order = [
    ...fileOrder.filter((n) => !removals.includes(n)).map(finalName),
    ...additions.map((a) => a.name),
  ]
  for (const move of tabs.moves) {
    const current = finalName(move.name)
    const at = order.indexOf(current)
    if (at < 0) continue
    order.splice(at, 1)
    order.splice(Math.min(move.position - 1, order.length), 0, current)
  }
  return {
    renames,
    additions,
    removals,
    order,
    ...(tabs.hidden.length
      ? { hiddenChanges: tabs.hidden.map((h) => ({ sheetName: h.name, hidden: h.hidden })) }
      : {}),
    orderChanged: tabs.moves.length > 0,
  }
}

/** DSL format patch → the gateway's style delta (the app does this via Univer's style model). */
export function patchToStyleEdit(p: CellFormatPatch): WorkbookStyleEdit {
  const out: Record<string, unknown> = {}
  for (const k of ['bold', 'italic', 'underline', 'strikethrough', 'wrapText'] as const) {
    if (p[k] !== undefined) out[k] = p[k] ?? false
  }
  if (p.fontFamily) out.fontFamily = p.fontFamily
  if (typeof p.fontSize === 'number') out.fontSize = p.fontSize
  if (p.fontColor !== undefined) out.fontColor = p.fontColor
  if (p.fillColor !== undefined) out.fillColor = p.fillColor
  if (p.numberFormat !== undefined) out.numberFormat = p.numberFormat ?? 'General'
  if (p.horizontalAlign) out.horizontalAlignment = p.horizontalAlign
  if (p.verticalAlign) out.verticalAlignment = p.verticalAlign
  if (p.indent !== undefined) out.indent = p.indent ?? 0
  if (p.textRotation !== undefined) {
    const r = p.textRotation
    // DSL: -90 (clockwise) … 90 (counterclockwise) or 'vertical'; OOXML: 0-90 ccw, 91-180 = clockwise + 90, 255 stacked
    out.textRotation = r === null ? 0 : r === 'vertical' ? 255 : r >= 0 ? r : 90 - r
  }
  if (p.border) {
    const edge =
      p.border.type === 'none'
        ? null
        : { style: 'thin' as const, ...(p.border.color ? { color: p.border.color } : {}) }
    const sides =
      p.border.type === 'all' || p.border.type === 'none'
        ? ['Top', 'Bottom', 'Left', 'Right']
        : [p.border.type[0]!.toUpperCase() + p.border.type.slice(1)]
    for (const side of sides) out[`border${side}`] = edge
  }
  return out as WorkbookStyleEdit
}

function describeError(err: unknown): string {
  if (err && typeof err === 'object' && 'issues' in err) {
    const issues = (err as { issues: { path: (string | number)[]; message: string }[] }).issues
    return issues.map((i) => `${i.path.join('.') || 'batch'}: ${i.message}`).join('; ')
  }
  return err instanceof Error ? err.message : String(err)
}
