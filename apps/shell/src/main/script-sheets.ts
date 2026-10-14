import { statSync } from 'node:fs'
import { dirname } from 'node:path'
import type { ScriptRunOutcome } from '@genoffice/agent-core'
import type { ScriptApplyResult } from '../shared/feature-channels'
import { readInput } from '@genoffice/cli/headless-docs'
import { runWorkbookDsl, writeWorkbook, type DslOutcome } from '@genoffice/cli/headless-formats'
import { readBasicWorkbook } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import {
  ScriptRunError,
  SHEETS_FORMAT_CAP,
  SHEETS_SNAPSHOT_CELL_CAP,
  type SandboxOutcome,
  type ScriptCellFormat,
  type SheetSnapshotSheet,
  type SheetsSnapshot,
} from './script-shared'

/**
 * The sheets half of the script kind: a bounded, sidecar-free snapshot of the
 * workbook (values, formulas, formats — pure JS through the gateway), and the
 * workbook DSL (the in-app propose_operations vocabulary) as the write path.
 *
 * Reads are capped: a workbook over the snapshot cell cap is refused with a
 * reason, and the sandbox refuses a range read over the per-read cap. The write
 * is the gateway's own atomic batch — one refused op rolls everything back — and
 * lands through writeWorkbook, the same writer the sheet command uses, so
 * formulas are evaluated and cached values written like any other save.
 *
 * The sandbox is synchronous in the shell main, so a huge workbook still costs
 * the process it runs on; the cell caps are the budget until the host moves to a
 * utility process (mods-implementation-report.md §8.2.2).
 */

type Scalar = string | number | boolean | null

export const sheetsHost = {
  async snapshot(documentPath: string): Promise<SheetsSnapshot> {
    return sheetsSnapshot(inputBuffer(documentPath))
  },
  async plan({
    outcome,
    documentPath,
  }: {
    outcome: SandboxOutcome
    documentPath: string
  }): Promise<ScriptRunOutcome> {
    if (!outcome.ops.length) {
      return { applied: 0, plan: [], failures: [], log: outcome.log, ms: outcome.ms }
    }
    const planned = await planSheets(inputBuffer(documentPath), outcome.ops, documentPath)
    return { ...planned, log: outcome.log, ms: outcome.ms }
  },
  async apply({
    outcome,
    documentPath,
  }: {
    outcome: SandboxOutcome
    documentPath: string
  }): Promise<ScriptApplyResult> {
    if (!outcome.ops.length) {
      throw new ScriptRunError(
        'the script proposed nothing, so there is nothing to apply',
        'check what the script reads: $.sheets.range("Sheet1!A1:F10") and $.sheets.propose([...])',
      )
    }
    return applySheets(inputBuffer(documentPath), outcome.ops, documentPath, outcome)
  },
}

/** readInput is the allowed-roots-checked reader the docs path uses; the DSL wants a Buffer. */
function inputBuffer(documentPath: string): Buffer {
  return Buffer.from(readInput(documentPath))
}

/** The snapshot a sheets script sees: values, formulas and formats, all bounded. */
export async function sheetsSnapshot(buffer: Buffer): Promise<SheetsSnapshot> {
  const imported = await readBasicWorkbook(buffer)
  let nonEmpty = 0
  for (const sheet of imported.snapshot.sheets) {
    for (const state of Object.values(sheet.cells)) {
      if (state.value !== null || state.formula !== undefined) nonEmpty++
    }
  }
  if (nonEmpty > SHEETS_SNAPSHOT_CELL_CAP) {
    throw new ScriptRunError(
      'this workbook holds ' +
        nonEmpty +
        ' non-empty cells; the script API reads at most ' +
        SHEETS_SNAPSHOT_CELL_CAP,
      'run it against a smaller workbook, or narrow the work to a few sheets',
    )
  }
  const { styles, activeIndex } = await readPackageExtras(buffer)
  const sheets: SheetSnapshotSheet[] = imported.snapshot.sheets.map((sheet) => {
    const cells: Record<string, Scalar> = {}
    const formulas: Record<string, string> = {}
    for (const [addr, state] of Object.entries(sheet.cells)) {
      if (state.formula !== undefined) formulas[addr] = state.formula
      const raw = (state.rawValue ?? state.value) as Scalar
      if (raw !== null && raw !== undefined) cells[addr] = raw
    }
    const bounds = addressBounds(Object.keys(sheet.cells))
    return {
      name: sheet.name,
      first: bounds.first,
      last: bounds.last,
      rows: bounds.rows,
      cols: bounds.cols,
      cells,
      formulas,
      formats: styles[sheet.name] ?? {},
    }
  })
  const active =
    imported.snapshot.sheets[activeIndex]?.name ?? imported.snapshot.sheets[0]?.name ?? null
  return { format: 'sheets', active, sheets }
}

// ── planning and applying: the workbook DSL as the engine ──────────────────────

/** Ops the dry run drops one at a time before giving up on the rest. */
const MAX_DRY_DROPS = 20

interface PlannedSheets {
  applied: number
  plan: string[]
  failures: string[]
}

/**
 * The dry run is the engine's own: the workbook DSL plans the ops against the
 * real file, and an op it refuses is dropped and reported while the rest still
 * validate (the sheet command's best-effort loop, bounded). The plan is a diff
 * summary — cells affected, then the first op lines — not a raw op list.
 */
async function planSheets(
  buffer: Buffer,
  ops: Record<string, unknown>[],
  documentPath: string,
): Promise<PlannedSheets> {
  const pending = ops.map((op, index) => ({ op, index }))
  const failures: string[] = []
  let outcome: DslOutcome | undefined
  while (pending.length) {
    try {
      outcome = await runWorkbookDsl(
        buffer,
        pending.map((entry) => entry.op),
        undefined,
        { cwd: dirname(documentPath), env: process.env },
        { sourcePath: documentPath },
      )
      break
    } catch (err) {
      const index = pinnedIndex(err)
      const message = errorMessage(err)
      if (index === undefined || index >= pending.length) {
        failures.push(message)
        break
      }
      const dropped = pending.splice(index, 1)[0]!
      failures.push(dropped.index + ' ' + opName(dropped.op) + ': ' + message)
      if (failures.length > MAX_DRY_DROPS) break
    }
  }
  if (!outcome) return { applied: 0, plan: [], failures }
  return { applied: pending.length, plan: diffSummary(outcome), failures }
}

async function applySheets(
  buffer: Buffer,
  ops: Record<string, unknown>[],
  documentPath: string,
  sandbox: SandboxOutcome,
): Promise<ScriptApplyResult> {
  let outcome: DslOutcome
  try {
    outcome = await runWorkbookDsl(
      buffer,
      ops,
      undefined,
      { cwd: dirname(documentPath), env: process.env },
      { sourcePath: documentPath },
    )
  } catch (err) {
    // atomic: one refused op aborts the batch, so the file is never half-written
    throw new ScriptRunError(
      'the workbook engine refused the batch, so nothing was written',
      errorMessage(err),
    )
  }
  await writeWorkbook(buffer, outcome.edits, documentPath, {
    plan: outcome.sheetPlan,
    structuralOps: outcome.structuralOps,
    renames: outcome.renames,
    gateway: outcome.gateway,
  })
  return {
    applied: ops.length,
    plan: diffSummary(outcome),
    failures: [],
    log: sandbox.log,
    ms: sandbox.ms,
    written: documentPath,
    bytes: statSync(documentPath).size,
  }
}

/** The diff summary a sheets run reports: cells affected, then the first op lines. */
function diffSummary(outcome: DslOutcome): string[] {
  const structural = outcome.structuralOps.reduce((n, s) => n + s.ops.length, 0)
  const renamed = Object.keys(outcome.renames).length
  const head =
    outcome.edits.length +
    ' cell write(s)' +
    (structural ? ', ' + structural + ' structural op(s)' : '') +
    (renamed ? ', ' + renamed + ' sheet(s) renamed' : '')
  const plan = [head, ...outcome.plan.slice(0, 10)]
  for (const warning of outcome.warnings.slice(0, 3)) plan.push('warning: ' + warning)
  return plan
}

function opName(op: unknown): string {
  return op && typeof op === 'object' && typeof (op as { op?: unknown }).op === 'string'
    ? (op as { op: string }).op
    : 'op'
}

function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err)
}

/** The op index a DSL rejection pins, when the error carries one. */
function pinnedIndex(err: unknown): number | undefined {
  const detail = (err as { detail?: { failures?: { index?: unknown }[] } } | null)?.detail
  const index = detail?.failures?.[0]?.index
  return typeof index === 'number' ? index : undefined
}

// ── reading: bounds, the active tab, and the compact style read ────────────────

function addressBounds(addresses: string[]): {
  first: string
  last: string
  rows: number
  cols: number
} {
  if (!addresses.length) return { first: 'A1', last: 'A1', rows: 0, cols: 0 }
  let minCol = Infinity
  let maxCol = 0
  let minRow = Infinity
  let maxRow = 0
  for (const address of addresses) {
    const parsed = parseAddress(address)
    if (!parsed) continue
    minCol = Math.min(minCol, parsed.col)
    maxCol = Math.max(maxCol, parsed.col)
    minRow = Math.min(minRow, parsed.row)
    maxRow = Math.max(maxRow, parsed.row)
  }
  return {
    first: colLabel(minCol) + (minRow + 1),
    last: colLabel(maxCol) + (maxRow + 1),
    rows: maxRow - minRow + 1,
    cols: maxCol - minCol + 1,
  }
}

function parseAddress(address: string): { col: number; row: number } | undefined {
  const m = /^([A-Za-z]+)(\d+)$/.exec(address)
  if (!m) return undefined
  let col = 0
  for (const ch of m[1]!.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64)
  return { col: col - 1, row: Number(m[2]) - 1 }
}

function colLabel(col: number): string {
  let label = ''
  let n = col + 1
  while (n > 0) {
    const mod = (n - 1) % 26
    label = String.fromCharCode(65 + mod) + label
    n = (n - 1 - mod) / 26
  }
  return label
}

/** The common builtin number formats, keyed by numFmtId; customs come from the file. */
const BUILTIN_NUMFMTS: Record<number, string> = {
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
}

function decodeXml(text: string): string {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

function rgbColor(raw: string | undefined): string | undefined {
  if (!raw) return undefined
  const hex = raw.replace(/^FF/i, '').replace(/[^0-9A-Fa-f]/g, '')
  return hex.length === 6 ? '#' + hex.toUpperCase() : undefined
}

/**
 * What the snapshot needs beyond cell values, read straight from the package so
 * the script host needs no native sidecar: the active tab, and per sheet the
 * styled cells (number format, bold, italic, fill, font color), capped like
 * every other read.
 */
async function readPackageExtras(buffer: Buffer): Promise<{
  styles: Record<string, Record<string, ScriptCellFormat>>
  activeIndex: number
}> {
  const JSZip = (await import('jszip')).default
  const zip = await JSZip.loadAsync(buffer)
  const workbookXml = (await zip.file('xl/workbook.xml')?.async('string')) ?? ''
  const activeIndex = Number(/activeTab="(\d+)"/.exec(workbookXml)?.[1] ?? 0) || 0
  const stylesXml = await zip.file('xl/styles.xml')?.async('string')
  if (!stylesXml) return { styles: {}, activeIndex }
  const customNumFmts: Record<number, string> = {}
  for (const [, id, code] of stylesXml.matchAll(
    /<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g,
  )) {
    customNumFmts[Number(id)] = decodeXml(code!)
  }
  const fonts: ScriptCellFormat[] = []
  for (const [, font] of stylesXml.matchAll(/<font\b[^>]*>([\s\S]*?)<\/font>/g)) {
    const bold = /<b\b/.test(font!)
    const italic = /<i\b/.test(font!)
    const color = rgbColor(/<color\b[^>]*\brgb="([^"]*)"/.exec(font!)?.[1])
    fonts.push({
      ...(bold ? { bold: true } : {}),
      ...(italic ? { italic: true } : {}),
      ...(color ? { fontColor: color } : {}),
    })
  }
  const fills: (string | undefined)[] = []
  for (const [, fill] of stylesXml.matchAll(/<fill\b[^>]*>([\s\S]*?)<\/fill>/g)) {
    fills.push(rgbColor(/<fgColor\b[^>]*\brgb="([^"]*)"/.exec(fill!)?.[1]))
  }
  const xfs: { numFmtId: number; fontId: number; fillId: number }[] = []
  for (const [, xf] of stylesXml.matchAll(/<xf\b([^>]*?)\/?>/g)) {
    const attrs = xf!
    if (!/\bnumFmtId=/.test(attrs)) continue
    xfs.push({
      numFmtId: Number(/numFmtId="(\d+)"/.exec(attrs)?.[1] ?? 0),
      fontId: Number(/fontId="(\d+)"/.exec(attrs)?.[1] ?? 0),
      fillId: Number(/fillId="(\d+)"/.exec(attrs)?.[1] ?? 0),
    })
  }
  const resolve = (styleIndex: number): ScriptCellFormat | undefined => {
    const xf = xfs[styleIndex]
    if (!xf) return undefined
    const numberFormat = customNumFmts[xf.numFmtId] ?? BUILTIN_NUMFMTS[xf.numFmtId]
    const font = fonts[xf.fontId]
    const fill = fills[xf.fillId]
    const fmt: ScriptCellFormat = {
      ...(numberFormat && numberFormat !== '@' ? { numberFormat } : {}),
      ...(font?.bold ? { bold: true } : {}),
      ...(font?.italic ? { italic: true } : {}),
      ...(fill ? { fillColor: fill } : {}),
      ...(font?.fontColor ? { fontColor: font.fontColor } : {}),
    }
    return Object.keys(fmt).length ? fmt : undefined
  }

  // sheet name → worksheet part, through the workbook's relationships
  const relsXml = (await zip.file('xl/_rels/workbook.xml.rels')?.async('string')) ?? ''
  const relTargets: Record<string, string> = {}
  for (const [, id, target] of relsXml.matchAll(
    /<Relationship\b[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"/g,
  )) {
    const value = target!
    relTargets[id!] = value.startsWith('/')
      ? 'xl' + value
      : value.startsWith('xl/')
        ? value
        : 'xl/' + value
  }
  const styles: Record<string, Record<string, ScriptCellFormat>> = {}
  let budget = SHEETS_FORMAT_CAP
  for (const [, name, rid] of workbookXml.matchAll(
    /<sheet\b[^>]*name="([^"]*)"[^>]*r:id="([^"]*)"/g,
  )) {
    const target = relTargets[rid!]
    if (!target || budget <= 0) continue
    const sheetXml = await zip.file(target)?.async('string')
    if (!sheetXml) continue
    const styled: Record<string, ScriptCellFormat> = {}
    for (const [, attrs] of sheetXml.matchAll(/<c\b([^>]*?)(?:\/>|>[\s\S]*?<\/c>)/g)) {
      if (budget <= 0) break
      const address = /(?:^|\s)r="([^"]*)"/.exec(attrs!)?.[1]
      const styleIndex = Number(/(?:^|\s)s="(\d+)"/.exec(attrs!)?.[1] ?? NaN)
      if (!address || !Number.isInteger(styleIndex)) continue
      const fmt = resolve(styleIndex)
      if (!fmt) continue
      styled[address] = fmt
      budget--
    }
    if (Object.keys(styled).length) styles[decodeXml(name!)] = styled
  }
  return { styles, activeIndex }
}
