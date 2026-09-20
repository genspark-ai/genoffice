import { existsSync, readFileSync, statSync } from 'node:fs'
import { basename, dirname, extname, isAbsolute, resolve } from 'node:path'
import { flagBool, flagString } from '../args'
import { assertAllowed, resolveInput, resolveOutput, writeOutput } from '../fs'
import { applyOps, blankDeck, inlineLocalFiles, parseOps, saveDeck } from '../formats/pptx'
import { buildDeckFromDir, buildDeckFromSpec, stageContext } from '../formats/slide-spec'
import {
  additionPlan,
  blankWorkbook,
  cellEditsFromTable,
  csvTable,
  writeWorkbook,
  type CsvTableOptions,
  type TableInput,
} from '../formats/xlsx'
import { exportViaApp } from '../formats/app-export'
import { blankDocument, closeDocument, fillFromHtml, saveDocument } from '../formats/docx'
import { markdownToDocx } from '../formats/markdown'
import { runWorkbookDsl } from '../formats/xlsx-dsl'
import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'
import { readOpsInput } from '../ops-input'
import type { CommandContext, CommandDef } from '../registry'
import type { TxnResult } from '@genoffice/pptx-ops'
import { CliError, EXIT, type CommandResult } from '../result'
import { txnFailure, txnDetail } from './txn'

/** Document types that can be built here. */
const TYPES = ['pptx', 'xlsx', 'docx', 'pdf'] as const

export const createCommand: CommandDef = {
  name: 'create',
  summary: 'Create a new document from structured content.',
  usage:
    'create --type <pptx|xlsx|docx|pdf> (--ops <file|-> | --spec <file|dir|-> [--outline <file>] | --from <file>) --out <path> [--force]',
  options: [
    { name: 'type', value: 'type', description: `document type: ${TYPES.join(', ')}` },
    {
      name: 'ops',
      value: 'file',
      description:
        'pptx: JSON array of ops (or {"ops":[...]}) applied one by one to a blank one-slide deck (later ops can target slides added earlier); "-" reads stdin. See `genoffice guide slides`.',
    },
    {
      name: 'spec',
      value: 'file',
      description:
        'pptx: a deck spec (pages of px-positioned text, shapes and images on a 1280x720 canvas) built into one slide per page, or a directory of one-page spec files taken in name order; "-" reads stdin. See `genoffice guide slides spec` and `genoffice guide slides design`.',
    },
    {
      name: 'outline',
      value: 'file',
      description:
        'pptx --spec <dir>: the deck outline the pages were written from (default: outline.json in or beside the directory); every page file is checked against its entry and the build refuses to run until every outline page has its file',
    },
    {
      name: 'from',
      value: 'file',
      description:
        'xlsx: a .csv, or a .json holding a 2-D array of cell values or { "sheets": [{ "name", "rows" }] }; strings starting with "=" are formulas. docx: a .md file, or a .html file holding a restricted-HTML fragment (see `genoffice guide docs`). pdf: any .md/.html/.docx/.xlsx/.pptx file, printed by the GenOffice renderer',
    },
    {
      name: 'header',
      description:
        'xlsx: treat row 1 of every sheet as a header — freeze it and put an AutoFilter on the used range',
    },
    {
      name: 'decimal',
      value: ', or .',
      description:
        'xlsx --from .csv: the decimal separator the file uses (default "."); with "," a "." is read as a thousands separator',
    },
    { name: 'out', value: 'path', description: 'output file (required)' },
    { name: 'force', description: 'overwrite an existing output file' },
  ],
  async run(args, ctx) {
    const type = flagString(args, 'type')?.toLowerCase()
    if (!type)
      throw new CliError(EXIT.usage, 'missing --type <type>', undefined, {
        reason: 'missing_argument',
      })
    if (!(TYPES as readonly string[]).includes(type)) {
      throw new CliError(
        EXIT.usage,
        `cannot create .${type} yet`,
        { supported: [...TYPES] },
        { reason: 'unsupported', suggestion: 'pick a type from detail.supported' },
      )
    }
    const output = resolveOutput(flagString(args, 'out'), ctx, {
      force: flagBool(args, 'force'),
      fresh: true,
    })
    if (type === 'pdf') {
      const from = flagString(args, 'from')
      if (!from)
        throw new CliError(EXIT.usage, 'pdf needs --from <document>', undefined, {
          reason: 'missing_argument',
        })
      const source = resolveInput(from, ctx)
      const r = await exportViaApp(source, 'pdf', output, { env: ctx.env, log: ctx.log })
      return {
        summary: `created ${basename(output)}`,
        outputPath: output,
        detail: { via: 'genoffice --headless-export', summary: r.summary },
      }
    }
    if (type === 'docx') {
      const r = await createDocx(args, ctx)
      writeOutput(output, r.bytes)
      return { summary: `created ${basename(output)}`, outputPath: output, detail: r.detail }
    }
    if (type === 'xlsx') {
      const r = await createXlsx(args, ctx, output)
      return {
        summary: `created ${basename(output)} (${r.sheets} sheets)`,
        outputPath: output,
        detail: r,
      }
    }
    const result =
      flagString(args, 'spec') !== undefined
        ? await createPptxFromSpec(args, ctx)
        : await createPptx(args, ctx)
    writeOutput(output, result.bytes)
    return {
      summary: `created ${basename(output)} (${result.slides} slides)`,
      outputPath: output,
      detail: result.detail,
    }
  },
}

async function createPptx(
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<{ bytes: Uint8Array; slides: number; detail: CommandResult['detail'] }> {
  const { text, source } = readOpsInput(args, ctx)
  const ops = inlineLocalFiles(parseOps(text, source), ctx)
  const opened = await blankDeck()
  // One transaction per op: the executor validates a batch against the
  // pre-transaction deck, so slides added earlier in the list would not yet
  // exist for later ops. Nothing is written unless every op applies.
  const records: NonNullable<TxnResult['records']> = []
  for (const [index, op] of ops.entries()) {
    const r = applyOps(opened, [op], { isolation: 'atomic' })
    if (!r.applied) throw txnFailure(r, index)
    records.push(...(r.records ?? []))
  }
  return {
    bytes: await saveDeck(opened),
    slides: opened.deck.slides.length,
    detail: txnDetail({ applied: true, records }),
  }
}

async function createPptxFromSpec(
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<{ bytes: Uint8Array; slides: number; detail: CommandResult['detail'] }> {
  const spec = flagString(args, 'spec')!
  const next =
    'genoffice slides audit <file> for the geometry audit; genoffice slides render <file> --out <dir> for PNGs; genoffice slides replace <file> --slide n --spec <page.json> to rebuild one page'
  const outlinePath = flagString(args, 'outline')
  const dir = spec === '-' ? null : specDirectory(spec, ctx)
  if (dir) {
    const stage = stageContext(dir, ctx, outlinePath ? resolveInput(outlinePath, ctx) : undefined)
    const built = await buildDeckFromDir(dir, ctx, stage)
    return {
      bytes: built.bytes,
      slides: built.pages,
      detail: {
        via: 'deck spec directory',
        files: built.files.map((f) => basename(f)),
        issues: built.issues,
        imageFailures: built.imageFailures,
        outline: stage.outline
          ? { file: stage.outline.file, findings: built.outlineFindings }
          : null,
        style: stage.style ? { file: stage.style.file, offPalette: built.offPalette } : null,
        notes: stage.notes,
        next,
      },
    }
  }
  if (outlinePath)
    throw new CliError(EXIT.usage, '--outline goes with --spec <directory>', undefined, {
      reason: 'invalid_argument',
    })
  const source = spec === '-' ? 'stdin' : resolveInput(spec, ctx)
  const text = spec === '-' ? readFileSync(0, 'utf-8') : readFileSync(source, 'utf-8')
  const built = await buildDeckFromSpec(
    text,
    source,
    ctx,
    spec === '-' ? undefined : dirname(source),
  )
  return {
    bytes: built.bytes,
    slides: built.pages,
    detail: { via: 'deck spec', issues: built.issues, imageFailures: built.imageFailures, next },
  }
}

function specDirectory(spec: string, ctx: CommandContext): string | null {
  const abs = isAbsolute(spec) ? spec : resolve(ctx.cwd, spec)
  if (!existsSync(abs) || !statSync(abs).isDirectory()) return null
  return assertAllowed(abs, ctx.env, 'read')
}

async function createXlsx(
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
  output: string,
): Promise<Record<string, unknown>> {
  const from = flagString(args, 'from')
  if (!from)
    throw new CliError(EXIT.usage, 'xlsx needs --from <data.csv|table.json>', undefined, {
      reason: 'missing_argument',
    })
  const source = resolveInput(from, ctx)
  const decimal = flagString(args, 'decimal')
  if (decimal !== undefined && decimal !== ',' && decimal !== '.') {
    throw new CliError(EXIT.usage, `--decimal must be "," or ".", got "${decimal}"`, undefined, {
      reason: 'invalid_argument',
    })
  }
  const tables = readTables(source, basename(source, extname(source)), { decimal })
  const [first, ...rest] = tables
  if (!first)
    throw new CliError(EXIT.usage, `${from}: no sheets`, undefined, { reason: 'invalid_argument' })
  const edits = tables.flatMap((t) => cellEditsFromTable(t.name, t.rows, t.formats))
  let r = await writeWorkbook(await blankWorkbook(first.name), edits, output, {
    plan: additionPlan(
      first.name,
      rest.map((t) => t.name),
    ),
  })
  if (flagBool(args, 'header')) {
    const ops = tables.flatMap((t) => headerOps(t))
    const dsl = await runWorkbookDsl(readFileSync(output), ops, first.name, ctx, {
      sourcePath: output,
    })
    const w = await writeWorkbook(readFileSync(output), dsl.edits, output, {
      plan: dsl.sheetPlan,
      structuralOps: dsl.structuralOps,
      renames: dsl.renames,
      gateway: dsl.gateway,
    })
    r = { ...r, warning: r.warning ?? w.warning, formulaError: r.formulaError ?? w.formulaError }
    for (const note of w.notes ?? []) ctx.warn(note)
  }

  if (r.warning) ctx.warn({ code: 'formulas_not_cached', message: r.warning })
  if (r.formulaError) ctx.warn({ code: 'formula_errors', message: r.formulaError })
  return {
    sheets: tables.length,
    cells: r.cells,
    formulas: r.formulas,
    cached_values: r.cachedValues,
  }
}

/** Freeze row 1 and filter the used range: what Excel's own import wizard does for a header row. */
function headerOps(table: TableInput): Record<string, unknown>[] {
  const width = table.rows.reduce((max, row) => Math.max(max, row.length), 0)
  if (table.rows.length < 1 || width < 1) return []
  const range = `A1:${columnLabel(width - 1)}${Math.max(table.rows.length, 2)}`
  return [
    { op: 'set_freeze', sheet: table.name, rows: 1, columns: 0 },
    { op: 'set_filter', sheet: table.name, range },
  ]
}

function readTables(
  source: string,
  fallbackName: string,
  opts: CsvTableOptions = {},
): TableInput[] {
  const sheetName = fallbackName.slice(0, 31) || 'Sheet1'
  if (extname(source).toLowerCase() === '.csv') {
    return [csvTable(readFileSync(source), sheetName, opts)]
  }
  let parsed: unknown
  try {
    parsed = JSON.parse(readFileSync(source, 'utf-8'))
  } catch (err) {
    throw new CliError(
      EXIT.usage,
      `${source}: not valid JSON (${(err as Error).message})`,
      undefined,
      { reason: 'invalid_json' },
    )
  }
  if (Array.isArray(parsed)) return [{ name: sheetName, rows: parsed as TableInput['rows'] }]
  const sheets = (parsed as { sheets?: unknown })?.sheets
  if (!Array.isArray(sheets) || sheets.length === 0) {
    throw new CliError(
      EXIT.usage,
      `${source}: expected a 2-D array or { "sheets": [{ "name", "rows" }] }`,
      undefined,
      { reason: 'invalid_argument' },
    )
  }
  return sheets.map((s, i) => {
    const t = s as Partial<TableInput>
    if (typeof t.name !== 'string' || !Array.isArray(t.rows)) {
      throw new CliError(EXIT.usage, `${source}: sheets[${i}] needs "name" and "rows"`)
    }
    return { name: t.name, rows: t.rows }
  })
}

async function createDocx(
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<{ bytes: Uint8Array; detail: Record<string, unknown> }> {
  const from = flagString(args, 'from')
  if (!from)
    throw new CliError(EXIT.usage, 'docx needs --from <content.md|fragment.html>', undefined, {
      reason: 'missing_argument',
    })
  const source = resolveInput(from, ctx)
  const text = readFileSync(source, 'utf-8')
  const ext = extname(source).toLowerCase()
  if (ext === '.md' || ext === '.markdown') {
    return {
      bytes: await markdownToDocx(text, { file: source, ctx }),
      detail: { source: 'markdown' },
    }
  }
  if (ext === '.html' || ext === '.htm') {
    const doc = await blankDocument()
    try {
      const blocks = fillFromHtml(doc, text)
      return { bytes: await saveDocument(doc), detail: { source: 'html', blocks } }
    } finally {
      closeDocument(doc)
    }
  }
  throw new CliError(
    EXIT.usage,
    `docx --from needs a .md or .html file, got ${ext || 'no extension'}`,
    undefined,
    { reason: 'unsupported' },
  )
}
