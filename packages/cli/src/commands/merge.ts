import { randomBytes } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { basename } from 'node:path'
import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'
import { flagBool, flagString } from '../args'
import {
  applyDocOps,
  blockRangeHtml,
  closeDocument,
  describeDocument,
  openDocument,
  saveDocument,
  type OpenDocument,
} from '../formats/docx'
import {
  applyOps,
  describeDeck,
  openDeck,
  saveDeck,
  type ElementSummary,
  type SlideSummary,
} from '../formats/pptx'
import { readSheet, workbookSummary, writeWorkbook } from '../formats/xlsx'
import { runWorkbookDsl } from '../formats/xlsx-dsl'
import { extension, readInput, resolveInput, resolveOutput, writeOutput } from '../fs'
import { readOpsStream } from '../ops-input'
import type { OpenedPptx } from '@genoffice/pptx-engine'
import type { CommandContext, CommandDef } from '../registry'
import { CliError, EXIT, type CommandResult } from '../result'

const PLACEHOLDER = /\{\{\s*([\w.-]+)\s*\}\}/g
const FORMATS = ['docx', 'pptx', 'xlsx'] as const
type Format = (typeof FORMATS)[number]

type Scalar = string | number | boolean | null
type Values = Map<string, Scalar>

interface Hit {
  placeholder: string
  key: string
  location: Record<string, string | number | boolean>
  /** set while scanning when no fill path reaches the spot */
  reason?: 'unreachable_nested'
}

interface Unresolved extends Omit<Hit, 'reason'> {
  reason: 'no_key' | 'split_placeholder' | 'unreachable_nested'
}

interface Merged {
  found: Hit[]
  /** text of every location holding a placeholder before the fill, and after it where a rescan is possible (docx, pptx) */
  before: Map<string, string>
  after?: Map<string, string>
  write(): Promise<void>
}

export const mergeCommand: CommandDef = {
  name: 'merge',
  summary:
    'Fill {{key}} placeholders in a .docx, .pptx or .xlsx template with values from a JSON object.',
  usage:
    'merge <template.docx|pptx|xlsx> --data <json-file|inline-json|-> --out <path> [--force] [--strict]',
  options: [
    {
      name: 'data',
      value: 'json',
      description:
        'a JSON object of values: a file path, inline JSON, or "-" for stdin; nested objects flatten to dotted keys (a.b), arrays are not expanded',
    },
    { name: 'out', value: 'path', description: 'output file (required)' },
    {
      name: 'force',
      description: 'overwrite an existing output file, or write while GenOffice has the file open',
    },
    {
      name: 'strict',
      description:
        'fail (exit 1, error unresolved_placeholder) instead of writing when a placeholder has no value',
    },
  ],
  async run(args, ctx) {
    const path = resolveInput(args.positionals[0], ctx)
    const format = extension(path)
    if (!(FORMATS as readonly string[]).includes(format)) {
      throw new CliError(
        EXIT.usage,
        `cannot merge .${format} templates`,
        { supported: [...FORMATS] },
        { reason: 'unsupported', suggestion: 'pass a .docx, .pptx or .xlsx template' },
      )
    }
    const { values, ignored } = readData(flagString(args, 'data'), ctx)
    const output = resolveOutput(flagString(args, 'out'), ctx, {
      force: flagBool(args, 'force'),
      fresh: true,
    })
    const merged = await MERGE[format as Format](path, values, output, ctx)
    const report = describe(merged, values)
    if (flagBool(args, 'strict') && report.unresolved.length) {
      throw new CliError(
        EXIT.usage,
        `${report.unresolved.length} unresolved placeholders in ${basename(path)}, nothing written`,
        { unresolved_placeholders: report.unresolved, unused_keys: report.unused },
        {
          reason: 'unresolved_placeholder',
          suggestion:
            'add the keys listed in detail.unresolved_placeholders to --data, or drop --strict to write anyway',
        },
      )
    }
    await merged.write()
    const result: CommandResult = {
      summary: `${report.filled} placeholders filled, ${report.unresolved.length} unresolved`,
      outputPath: output,
      detail: {
        format,
        filled: report.filled,
        used_keys: report.used,
        unresolved_placeholders: report.unresolved,
        unused_keys: report.unused,
        ...(ignored.length ? { ignored_keys: ignored } : {}),
      },
    }
    if (report.unresolved.length) {
      const shown = report.unresolved.slice(0, 5).map((u) => `${u.placeholder} (${where(u)})`)
      const nested = report.unresolved.some((u) => u.reason === 'unreachable_nested')
      result.warnings = [
        {
          code: 'unresolved_placeholder',
          message: `${report.unresolved.length} placeholders were left in place: ${shown.join(', ')}${report.unresolved.length > shown.length ? ', …' : ''}`,
          suggestion: `add the missing keys to --data${nested ? ', ungroup the nested groups' : ''}, or pass --strict to fail instead`,
        },
      ]
    }
    return result
  },
}

function where(u: Unresolved): string {
  return Object.entries(u.location)
    .map(([k, v]) => (v === true ? k : `${k} ${v}`))
    .join(' ')
}

function readData(
  spec: string | undefined,
  ctx: CommandContext,
): { values: Values; ignored: string[] } {
  if (!spec)
    throw new CliError(EXIT.usage, 'missing --data <json-file|inline-json|->', undefined, {
      reason: 'missing_argument',
    })
  const trimmed = spec.trim()
  const [text, source] =
    trimmed === '-'
      ? [readOpsStream(0, 'stdin'), 'stdin']
      : /^[[{]/.test(trimmed)
        ? [trimmed, '--data']
        : [readFileSync(resolveInput(spec, ctx), 'utf-8'), spec]
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch (err) {
    throw new CliError(
      EXIT.usage,
      `${source}: not valid JSON (${(err as Error).message})`,
      undefined,
      { reason: 'invalid_json' },
    )
  }
  if (!isRecord(parsed)) {
    throw new CliError(EXIT.usage, `${source}: expected a JSON object of values`, undefined, {
      reason: 'invalid_argument',
      suggestion: 'pass an object such as {"name": "Ada", "amount": 12}',
    })
  }
  const values: Values = new Map()
  const ignored: string[] = []
  flatten(parsed, '', values, ignored)
  return { values, ignored }
}

function isRecord(v: unknown): v is Record<string, unknown> {
  return v !== null && typeof v === 'object' && !Array.isArray(v)
}

function flatten(obj: Record<string, unknown>, prefix: string, out: Values, ignored: string[]) {
  for (const [k, v] of Object.entries(obj)) {
    const key = prefix ? `${prefix}.${k}` : k
    if (isRecord(v)) flatten(v, key, out, ignored)
    else if (v === null || ['string', 'number', 'boolean'].includes(typeof v))
      out.set(key, v as Scalar)
    else ignored.push(key)
  }
}

function textOf(v: Scalar): string {
  return v === null ? '' : String(v).replace(/\0/g, '')
}

function hits(text: string): Array<Pick<Hit, 'placeholder' | 'key'>> {
  return [...text.matchAll(PLACEHOLDER)].map((m) => ({ placeholder: m[0], key: m[1]! }))
}

/** every placeholder with a value swapped for it in one pass, the rest left as written */
function substitute(text: string, values: Values): string {
  return text.replace(PLACEHOLDER, (m, key: string) =>
    values.has(key) ? textOf(values.get(key)!) : m,
  )
}

const locKey = (location: Hit['location']): string => JSON.stringify(location)

const occurrences = (text: string, part: string): number => text.split(part).length - 1

/**
 * Every hit is classified from the pre-fill scan, so placeholder-shaped text
 * that arrived inside a value is output, not a leftover. The rescan only tells
 * which provided placeholders are still at their spot: where the text is not
 * the expected substitution, the occurrences beyond the expected text's own
 * are placeholders the run-level replace could not match (split over runs).
 */
function describe(merged: Merged, values: Values) {
  const leftover = new Map<string, number>()
  for (const [loc, before] of merged.before) {
    const actual = merged.after?.get(loc)
    const expected = substitute(before, values)
    if (actual === undefined || actual === expected) continue
    for (const h of hits(before)) {
      const k = `${loc}|${h.placeholder}`
      if (!leftover.has(k)) {
        leftover.set(
          k,
          Math.max(0, occurrences(actual, h.placeholder) - occurrences(expected, h.placeholder)),
        )
      }
    }
  }
  const unresolved: Unresolved[] = []
  const used: string[] = []
  let filled = 0
  for (const h of merged.found) {
    const reason = h.reason ?? (values.has(h.key) ? undefined : 'no_key')
    if (reason) {
      unresolved.push({ ...h, reason })
      continue
    }
    const k = `${locKey(h.location)}|${h.placeholder}`
    const left = leftover.get(k) ?? 0
    if (left > 0) {
      leftover.set(k, left - 1)
      unresolved.push({ ...h, reason: 'split_placeholder' })
      continue
    }
    filled++
    if (!used.includes(h.key)) used.push(h.key)
  }
  return {
    filled,
    used,
    unused: [...values.keys()].filter((k) => !used.includes(k)),
    unresolved,
  }
}

type Merge = (path: string, values: Values, output: string, ctx: CommandContext) => Promise<Merged>

const MERGE: Record<Format, Merge> = {
  docx: mergeDocx,
  pptx: mergePptx,
  xlsx: mergeXlsx,
}

/**
 * findReplace ops for the placeholders `provided` names, as one pass: each
 * placeholder first becomes a private-use sentinel, then every sentinel
 * becomes its value, so a value that reads like another placeholder is never
 * matched by a later op.
 */
type ReplaceOp = Record<string, unknown> & {
  op: 'findReplace'
  find: string
  replace: string
  matchCase: true
}

function sentinelOps(
  provided: Hit[],
  values: Values,
  target?: { blockIndexes: number[] },
): ReplaceOp[] {
  const keys = new Map(provided.map((h) => [h.placeholder, h.key]))
  const distinct = [...keys.keys()]
  // plain ASCII: the pptx engine drops private-use characters
  const tag = randomBytes(6).toString('hex')
  const sentinel = (i: number) => `[[goff-${tag}-${i}]]`
  const replace = (find: string, replace: string): ReplaceOp => ({
    op: 'findReplace',
    find,
    replace,
    matchCase: true,
    ...(target ? { target } : {}),
  })
  return [
    ...distinct.map((find, i) => replace(find, sentinel(i))),
    ...distinct.map((find, i) => replace(sentinel(i), textOf(values.get(keys.get(find)!)!))),
  ]
}

function scanDocx(doc: OpenDocument): {
  found: Hit[]
  tables: Set<number>
  texts: Map<string, string>
} {
  const found: Hit[] = []
  const tables = new Set<number>()
  const texts = new Map<string, string>()
  for (const b of describeDocument(doc, undefined, Infinity)) {
    const inBlock = hits(b.text)
    if (!inBlock.length) continue
    texts.set(locKey({ block: b.index }), b.text)
    if (b.type === 'table') tables.add(b.index)
    for (const h of inBlock) found.push({ ...h, location: { block: b.index } })
  }
  return { found, tables, texts }
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

/**
 * Word splits a placeholder over several runs whenever the formatting changes
 * inside it; findReplace matches within one run, so what it misses is rescanned
 * and reported as split_placeholder. findReplace reaches neither table cells
 * nor a line break (a hardBreak only comes in through HTML), so a table with a
 * provided placeholder and a paragraph whose value spans lines are round-tripped
 * through replace_blocks on their restricted HTML; the rest is one findReplace
 * pass.
 */
async function mergeDocx(path: string, values: Values, output: string): Promise<Merged> {
  const source = readInput(path)
  const doc = await openDocument(source)
  try {
    const { found, tables, texts } = scanDocx(doc)
    const provided = found.filter((h) => values.has(h.key))
    const viaHtml = new Set(
      provided
        .filter(
          (h) =>
            tables.has(h.location.block as number) || textOf(values.get(h.key)!).includes('\n'),
        )
        .map((h) => h.location.block as number),
    )
    // the two routes never touch the same block: findReplace is confined to the plain
    // blocks and runs first, the HTML for the others is built from the untouched document
    const plain = provided.filter((h) => !viaHtml.has(h.location.block as number))
    const ops: Record<string, unknown>[] = plain.length
      ? sentinelOps(plain, values, {
          blockIndexes: [...new Set(plain.map((h) => h.location.block as number))],
        })
      : []
    for (const block of [...viaHtml].sort((a, b) => b - a)) {
      ops.push({
        op: 'replace_blocks',
        startBlockIndex: block,
        endBlockIndex: block,
        html: blockRangeHtml(doc, block, block).replace(PLACEHOLDER, (m, key: string) =>
          values.has(key) ? escapeHtml(textOf(values.get(key)!)).replace(/\n/g, '<br>') : m,
        ),
      })
    }
    if (ops.length) await applyDocOps(doc, ops, { mode: 'best_effort' })
    const after = ops.length ? scanDocx(doc).texts : texts
    const bytes = ops.length ? await saveDocument(doc) : source
    return { found, before: texts, after, write: async () => writeOutput(output, bytes) }
  } finally {
    closeDocument(doc)
  }
}

function scanDeck(opened: OpenedPptx): {
  found: Hit[]
  notes: SlideSummary[]
  texts: Map<string, string>
} {
  const found: Hit[] = []
  const notes: SlideSummary[] = []
  const texts = new Map<string, string>()
  const record = (location: Hit['location'], text: string, reason?: Hit['reason']) => {
    const inText = hits(text)
    if (!inText.length) return
    texts.set(locKey(location), text)
    for (const h of inText) found.push({ ...h, location, ...(reason ? { reason } : {}) })
  }
  // deck findReplace rewrites top-level elements and a group's direct text children only
  const walk = (slide: string, el: ElementSummary, depth: number) => {
    const reachable = depth === 0 || (depth === 1 && (el.type === 'text' || el.type === 'shape'))
    record(
      { slide, element: el.id ?? el.type },
      el.text ?? '',
      reachable ? undefined : 'unreachable_nested',
    )
    for (const child of el.children ?? []) walk(slide, child, depth + 1)
  }
  for (const page of describeDeck(opened, undefined, true).pages) {
    for (const el of page.elements) walk(page.id, el, 0)
    if (hits(page.notes ?? '').length) notes.push(page)
    record({ slide: page.id, notes: true }, page.notes ?? '')
  }
  return { found, notes, texts }
}

/** deck-wide findReplace covers text, shapes, tables and groups; notes are rewritten with setNotes */
async function mergePptx(path: string, values: Values, output: string): Promise<Merged> {
  const source = readInput(path)
  const opened = await openDeck(source)
  const { found, notes, texts } = scanDeck(opened)
  const provided = found.filter((h) => values.has(h.key) && !h.location.notes && !h.reason)
  let changed = false
  for (const op of sentinelOps(provided, values)) {
    if (applyOps(opened, [op], { isolation: 'per_op' }).applied) changed = true
  }
  for (const page of notes) {
    const text = substitute(page.notes!, values)
    if (text === page.notes) continue
    const r = applyOps(opened, [{ op: 'setNotes', target: { slide: page.index }, text }], {
      isolation: 'per_op',
    })
    changed ||= r.applied
  }
  const after = changed ? scanDeck(opened).texts : texts
  const bytes = changed ? await saveDeck(opened) : source
  return { found, before: texts, after, write: async () => writeOutput(output, bytes) }
}

/**
 * One typed set_cell per placeholder cell: a cell that is exactly one
 * placeholder takes the value's own type (a number stays a number), every
 * other cell gets its substituted text, marked literal so a value starting
 * with "=" is not turned into a formula. Plain set_cell ops are not subject
 * to the DSL's range-expansion cap, so any sheet size works.
 */
async function mergeXlsx(
  path: string,
  values: Values,
  output: string,
  ctx: CommandContext,
): Promise<Merged> {
  const found: Hit[] = []
  const before = new Map<string, string>()
  const ops: Record<string, unknown>[] = []
  for (const sheet of (await workbookSummary(path)).sheets) {
    if (!sheet.rows || !sheet.columns) continue
    const range = `A1:${columnLabel(sheet.columns - 1)}${sheet.rows}`
    const read = await readSheet(path, { sheet: sheet.name, range })
    read.rows.forEach((row, r) =>
      row.forEach((value, c) => {
        const address = `${columnLabel(c)}${r + 1}`
        if (typeof value !== 'string' || read.formulas[address]) return
        const inCell = hits(value)
        if (!inCell.length) return
        const location = { sheet: sheet.name, cell: address }
        before.set(locKey(location), value)
        for (const h of inCell) found.push({ ...h, location })
        if (!inCell.some((h) => values.has(h.key))) return
        const whole = inCell.length === 1 && value.trim() === inCell[0]!.placeholder
        const typed = whole ? values.get(inCell[0]!.key)! : undefined
        ops.push(
          whole && typeof typed !== 'string'
            ? { op: 'set_cell', sheet: sheet.name, address, value: typed }
            : {
                op: 'set_cell',
                sheet: sheet.name,
                address,
                value: substitute(value, values),
                type: 'text',
              },
        )
      }),
    )
  }
  return {
    found,
    before,
    write: async () => {
      if (!ops.length) return writeOutput(output, readInput(path))
      const source = readFileSync(path)
      const r = await runWorkbookDsl(source, ops, undefined, ctx, { sourcePath: path })
      await writeWorkbook(source, r.edits, output, {
        plan: r.sheetPlan,
        structuralOps: r.structuralOps,
        renames: r.renames,
        gateway: r.gateway,
      })
    },
  }
}
