import { readFileSync } from 'node:fs'
import { basename } from 'node:path'
import { flagBool, flagString } from '../args'
import {
  cellEditsFromInputs,
  readSheet,
  workbookSummary,
  writeWorkbook,
  type CellInput,
} from '../formats/xlsx'
import { runWorkbookDsl, SUPPORTED_DSL_OPS } from '../formats/xlsx-dsl'
import { hasGatewayPayloads, type GatewayPayloads } from '../formats/xlsx-gateway-ops'
import { resolveInput, resolveOutput } from '../fs'
import type { CommandContext, CommandDef } from '../registry'
import { CliError, EXIT, type CommandResult } from '../result'

export const sheetCommand: CommandDef = {
  name: 'sheet',
  summary: 'Read or edit an .xlsx with the same write path the app uses.',
  usage: 'sheet <read|apply> <file.xlsx> (--cells <file|-> | --ops <file|->) [options]',
  options: [
    { name: 'sheet', value: 'name', description: 'worksheet (default: the active one)' },
    { name: 'range', value: 'A1:D20', description: 'read: cell range (default: up to 500×100)' },
    {
      name: 'formats',
      description:
        'read: also return cell formats (bold, fill, number format, alignment) by address, column widths and row heights',
    },
    {
      name: 'cells',
      value: 'file',
      description:
        'apply: JSON array of { "cell": "B2", "sheet"?, "value"? | "formula"?, "style"? }; "-" reads stdin',
    },
    {
      name: 'ops',
      value: 'file',
      description:
        'apply: JSON array of workbook DSL ops (the in-app propose_operations vocabulary; `genoffice guide sheets`); "-" reads stdin',
    },
    { name: 'dry-run', description: 'apply --ops: validate and print the plan without writing' },
    { name: 'out', value: 'path', description: 'apply: write here instead of in place' },
    {
      name: 'force',
      description:
        'apply: overwrite an existing --out file, or write while GenOffice has the file open',
    },
  ],
  async run(args, ctx) {
    const [verb, file] = args.positionals
    switch (verb) {
      case 'read':
        return read(file, args, ctx)
      case 'apply':
        return apply(file, args, ctx)
      default:
        throw new CliError(EXIT.usage, 'expected "sheet read <file>" or "sheet apply <file>"')
    }
  },
}

type Args = Parameters<CommandDef['run']>[0]

async function read(
  file: string | undefined,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const data = await readSheet(path, {
    sheet: flagString(args, 'sheet'),
    range: flagString(args, 'range'),
    formats: flagBool(args, 'formats'),
  })
  const filled = data.rows.filter((r) => r.some((v) => v !== null)).length
  return {
    summary: `${basename(path)} · ${data.sheet}!${data.range}: ${filled} non-empty rows`,
    detail: {
      ...data,
      units:
        'rows are 0-based row order within the range; formulas and formats keyed by A1 address; features describe the sheet (merges and hyperlinks: the range)',
    },
  }
}

async function apply(
  file: string | undefined,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const cellsSpec = flagString(args, 'cells')
  const opsSpec = flagString(args, 'ops')
  if (!cellsSpec && !opsSpec)
    throw new CliError(EXIT.usage, 'missing --cells <file|-> or --ops <file|->')
  if (opsSpec) return applyOps(path, opsSpec, args, ctx)
  return applyCells(path, cellsSpec!, args, ctx)
}

function readJsonList(spec: string, flag: string, ctx: CommandContext): unknown[] {
  const text =
    spec === '-' ? readFileSync(0, 'utf-8') : readFileSync(resolveInput(spec, ctx), 'utf-8')
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch (err) {
    throw new CliError(EXIT.usage, `${flag}: not valid JSON (${(err as Error).message})`)
  }
  const list = Array.isArray(parsed)
    ? parsed
    : parsed && typeof parsed === 'object' && Array.isArray((parsed as { ops?: unknown }).ops)
      ? (parsed as { ops: unknown[] }).ops
      : null
  if (!list?.length) throw new CliError(EXIT.usage, `${flag}: expected a non-empty JSON array`)
  return list
}

async function applyOps(
  path: string,
  spec: string,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const ops = readJsonList(spec, '--ops', ctx)
  const source = readFileSync(path)
  const r = await runWorkbookDsl(source, ops, flagString(args, 'sheet'), ctx)
  const detail: Record<string, unknown> = {
    plan: r.plan,
    cells: r.edits.length,
    structural: r.structuralOps.reduce((n, s) => n + s.ops.length, 0),
    ...(hasGatewayPayloads(r.gateway) ? { features: featureCounts(r.gateway) } : {}),
    sheets_changed: r.sheetPlan
      ? r.sheetPlan.additions.length + r.sheetPlan.removals.length + r.sheetPlan.renames.length
      : 0,
    ...(r.warnings.length ? { warnings: r.warnings } : {}),
  }
  if (flagBool(args, 'dry-run')) {
    return { summary: `dry run: ${ops.length} ops validated, nothing written`, detail }
  }
  const output = resolveOutput(flagString(args, 'out'), ctx, {
    fallback: path,
    force: flagBool(args, 'force'),
    // in-place edits overwrite by design; --out onto another existing file needs --force
    fresh: flagString(args, 'out') !== undefined,
  })
  const w = await writeWorkbook(source, r.edits, output, {
    plan: r.sheetPlan,
    structuralOps: r.structuralOps,
    renames: r.renames,
    gateway: r.gateway,
  })
  return {
    summary: `applied ${ops.length} ops to ${basename(output)}`,
    outputPath: output,
    detail: {
      ...detail,
      formulas: w.formulas,
      cached_values: w.cachedValues,
      ...(w.warning ? { warning: w.warning } : {}),
    },
  }
}

/** What the batch writes beyond cells, for the plan report. */
function featureCounts(g: GatewayPayloads): Record<string, number> {
  const counts: Record<string, number> = {
    charts: g.visualAdditions.filter((v) => v.chart).length + g.chartEdits.length,
    images: g.visualAdditions.filter((v) => v.image).length,
    shapes: g.visualAdditions.filter((v) => v.shape).length,
    tables: g.tableAdditions.length,
    hyperlinks: g.hyperlinkEdits.reduce((n, s) => n + s.edits.length, 0),
    notes: g.noteStates.length,
    filters: g.filterStates.length,
    conditional_formats: g.cfStates.reduce((n, s) => n + s.rules.length, 0),
    data_validations: g.dvStates.reduce((n, s) => n + s.rules.length, 0),
    page_setup: g.pageSetupStates.length,
    hidden_ranges: g.hiddenOps.reduce((n, s) => n + s.ops.length, 0),
    sheet_tabs: g.tabs.moves.length + g.tabs.hidden.length + g.tabs.duplicates.length,
    protections: g.sheetProtections.length,
    defined_names: g.definedNamesState?.names.length ?? 0,
  }
  return Object.fromEntries(Object.entries(counts).filter(([, n]) => n > 0))
}

async function applyCells(
  path: string,
  spec: string,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const inputs = readJsonList(spec, '--cells', ctx)
  // addresses are checked before the workbook engine is even started
  cellEditsFromInputs(inputs as CellInput[], '')
  const output = resolveOutput(flagString(args, 'out'), ctx, {
    fallback: path,
    force: flagBool(args, 'force'),
    // in-place edits overwrite by design; --out onto another existing file needs --force
    fresh: flagString(args, 'out') !== undefined,
  })
  const summary = await workbookSummary(path)
  const defaultSheet = flagString(args, 'sheet') ?? summary.activeSheet ?? summary.sheets[0]?.name
  if (!defaultSheet) throw new CliError(EXIT.conversion, 'workbook has no sheets')
  const known = new Set(summary.sheets.map((s) => s.name))
  const edits = cellEditsFromInputs(inputs as CellInput[], defaultSheet)
  const unknown = edits.find((e) => !known.has(e.sheetName))
  if (unknown) {
    throw new CliError(EXIT.usage, `sheet not found: ${unknown.sheetName}`, { sheets: [...known] })
  }
  const r = await writeWorkbook(readFileSync(path), edits, output)
  return {
    summary: `wrote ${r.cells} cells to ${basename(output)}`,
    outputPath: output,
    detail: {
      cells: r.cells,
      formulas: r.formulas,
      cached_values: r.cachedValues,
      ...(r.warning ? { warning: r.warning } : {}),
    },
  }
}

export { SUPPORTED_DSL_OPS }
