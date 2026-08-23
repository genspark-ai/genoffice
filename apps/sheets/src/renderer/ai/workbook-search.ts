/**
 * Workbook-wide search and view navigation for the AI skill layer
 * (find_cells / select_range). Demo workbooks scan the in-memory snapshot;
 * lazy workbooks page through the file model via readSheetRangeMapped —
 * structural ops translated, session cell edits overlaid from the journal —
 * so the search never has to stream the whole workbook into Univer.
 */
import { formatAddress, type RangeBounds } from '../../domain/cell-address'
import type { CellScalar } from '../../domain/workbook.types'
import { ensureLazyRangeLoaded, readSheetRangeMapped } from '../univer-sync'
import type { LazyWorkbookState } from '../univer-state'
import { netAxisDelta } from '../view-transform'
import type {
  FindCellsMatch,
  FindCellsOptions,
  FindCellsOutcome,
  SelectRangeOutcome,
} from './tools'
import type { WorkbookReadContext } from './workbook-readers'

export const ERROR_VALUE_RE = /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!)$/
/** Total cells (by scanned extent) one find_cells call may cover. */
const MAX_SCAN_CELLS = 400_000
/** Row batches sized to stay under the sidecar's per-read cell budget. */
const FILE_READ_BATCH_CELLS = 18_000

type Matcher = (value: CellScalar | undefined, formula: string | undefined) => boolean

function buildMatcher(
  options: FindCellsOptions,
): { ok: true; test: Matcher } | { ok: false; error: string } {
  if (options.errorsOnly) {
    return {
      ok: true,
      test: (value) => typeof value === 'string' && ERROR_VALUE_RE.test(value),
    }
  }
  let matchText: (text: string) => boolean
  if (options.regex) {
    let pattern: RegExp
    try {
      pattern = new RegExp(options.query, 'i')
    } catch (error) {
      return {
        ok: false,
        error: `Invalid regex: ${error instanceof Error ? error.message : options.query}`,
      }
    }
    matchText = (text) => pattern.test(text)
  } else {
    const needle = options.query.toLowerCase()
    matchText = (text) => text.toLowerCase().includes(needle)
  }
  return {
    ok: true,
    test: (value, formula) => {
      if (
        options.lookIn !== 'formulas' &&
        value !== null &&
        value !== undefined &&
        matchText(String(value))
      ) {
        return true
      }
      return options.lookIn !== 'values' && !!formula && matchText(formula)
    },
  }
}

export async function findWorkbookCells(
  ctx: WorkbookReadContext,
  options: FindCellsOptions,
): Promise<FindCellsOutcome> {
  const matcher = buildMatcher(options)
  if (!matcher.ok) {
    return { matches: [], truncated: false, incompleteSheets: [], error: matcher.error }
  }
  const state = ctx.lazyWorkbookRef.current
  return state
    ? findInLazyWorkbook(ctx, state, options, matcher.test)
    : findInDemoWorkbook(ctx, options, matcher.test)
}

function findInDemoWorkbook(
  ctx: WorkbookReadContext,
  options: FindCellsOptions,
  test: Matcher,
): FindCellsOutcome {
  const snapshot = ctx.adapterRef.current.getSnapshot()
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const targets = options.sheetId
    ? snapshot.sheets.filter((sheet) => sheet.id === options.sheetId)
    : snapshot.sheets
  if (options.sheetId && targets.length === 0) {
    return {
      matches: [],
      truncated: false,
      incompleteSheets: [],
      error: `Unknown sheet: ${options.sheetId}`,
    }
  }
  const matches: FindCellsMatch[] = []
  let truncated = false
  for (const sheet of targets) {
    const worksheet = workbook?.getSheetBySheetId(sheet.id)
    for (const [address, cell] of Object.entries(sheet.cells)) {
      let value = cell.value
      // The in-memory model stores value:null for formula cells; matching
      // needs the computed value, backfilled from Univer's formula engine
      if (cell.formula && value === null && worksheet) {
        try {
          value = (worksheet.getRange(address).getValue() as CellScalar) ?? null
        } catch {
          /* fail-open: match on the formula text alone */
        }
      }
      if (!test(value, cell.formula)) continue
      if (matches.length >= options.maxResults) {
        truncated = true
        break
      }
      matches.push({ sheetName: sheet.name, address, value, formula: cell.formula })
    }
    if (truncated) break
  }
  return { matches, truncated, incompleteSheets: [] }
}

async function findInLazyWorkbook(
  ctx: WorkbookReadContext,
  state: LazyWorkbookState,
  options: FindCellsOptions,
  test: Matcher,
): Promise<FindCellsOutcome> {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  if (!workbook) {
    return {
      matches: [],
      truncated: false,
      incompleteSheets: [],
      error: 'No workbook is currently open.',
    }
  }
  const sheets = workbook.getSheets()
  const targets = options.sheetId
    ? sheets.filter((sheet) => sheet.getSheetId() === options.sheetId)
    : sheets
  if (options.sheetId && targets.length === 0) {
    return {
      matches: [],
      truncated: false,
      incompleteSheets: [],
      error: `Unknown sheet: ${options.sheetId}`,
    }
  }
  const matches: FindCellsMatch[] = []
  const incompleteSheets: string[] = []
  let truncated = false
  let scanBudget = MAX_SCAN_CELLS
  const push = (match: FindCellsMatch): boolean => {
    if (matches.length >= options.maxResults) {
      truncated = true
      return false
    }
    matches.push(match)
    return true
  }
  for (const worksheet of targets) {
    const sheetId = worksheet.getSheetId()
    const sheetName = worksheet.getSheetName()
    // Session edits first: they shadow file cells at the same coordinates
    const journal = state.editJournal.cells.get(sheetId)
    const shadowed = new Set<string>()
    for (const entry of journal?.values() ?? []) {
      if (!entry.hasValue) continue
      shadowed.add(`${entry.row}:${entry.column}`)
      const address = formatAddress(entry.row, entry.column)
      let value = entry.value
      // Journal formula entries store value:null; the computed result lives in
      // Univer (journal edits are always applied there) — backfill it so value
      // matches and errors_only see session-written formulas, like the demo path
      if (entry.formula && value === null) {
        try {
          value = (worksheet.getRange(address).getValue() as CellScalar) ?? null
        } catch {
          /* fail-open: match on the formula text alone */
        }
      }
      if (
        test(value, entry.formula) &&
        !push({ sheetName, address, value, formula: entry.formula })
      ) {
        break
      }
    }
    if (truncated) break
    const meta = state.file.sheets.find((candidate) => candidate.id === sheetId)
    // Sheets added this session live entirely in the journal
    if (!meta || meta.rowCount <= 0 || meta.columnCount <= 0) continue
    const ops = state.editJournal.structuralOps.get(sheetId) ?? []
    const screenRows = meta.rowCount + netAxisDelta(ops, 'row')
    const screenColumns = meta.columnCount + netAxisDelta(ops, 'column')
    if (screenRows <= 0 || screenColumns <= 0) continue
    const batchRows = Math.max(1, Math.floor(FILE_READ_BATCH_CELLS / screenColumns))
    let sheetIncomplete = false
    for (let startRow = 0; startRow < screenRows && !truncated; startRow += batchRows) {
      if (scanBudget <= 0) {
        truncated = true
        break
      }
      const endRow = Math.min(startRow + batchRows - 1, screenRows - 1)
      scanBudget -= (endRow - startRow + 1) * screenColumns
      let mapped
      try {
        mapped = await readSheetRangeMapped(
          state,
          sheetId,
          { startRow, endRow, startColumn: 0, endColumn: screenColumns - 1 },
          meta,
        )
      } catch {
        sheetIncomplete = true
        break
      }
      if (!mapped) continue
      if (
        !mapped.raw.indexingComplete &&
        (mapped.indexedThroughScreen === null || mapped.indexedThroughScreen < endRow)
      ) {
        sheetIncomplete = true
      }
      for (const cell of mapped.screen.cells) {
        if (shadowed.has(`${cell.row}:${cell.column}`)) continue
        if (!test(cell.value, cell.formula)) continue
        if (
          !push({
            sheetName,
            address: formatAddress(cell.row, cell.column),
            value: cell.value,
            formula: cell.formula,
          })
        ) {
          break
        }
      }
    }
    if (sheetIncomplete) incompleteSheets.push(sheetName)
    if (truncated) break
  }
  return { matches, truncated, incompleteSheets }
}

/** Activate the target sheet, select the range, and scroll it into view. */
export async function selectWorkbookRange(
  ctx: WorkbookReadContext,
  sheetId: string | undefined,
  bounds: RangeBounds,
  setMessage: (message: string) => void,
): Promise<SelectRangeOutcome> {
  const runtime = ctx.univerRef.current
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!runtime || !workbook) return { ok: false, error: 'No workbook is currently open.' }
  const worksheet = sheetId ? workbook.getSheetBySheetId(sheetId) : workbook.getActiveSheet()
  if (!worksheet) return { ok: false, error: `Unknown sheet: ${sheetId}` }
  try {
    if (worksheet.getSheetId() !== workbook.getActiveSheet()?.getSheetId()) {
      workbook.setActiveSheet(worksheet)
    }
    if (ctx.lazyWorkbookRef.current) {
      // Best-effort: selection works on not-yet-streamed cells too, but
      // loading the range means the user sees data, not an empty grid
      await ensureLazyRangeLoaded(
        runtime,
        ctx.lazyWorkbookRef,
        worksheet,
        { ...bounds },
        setMessage,
      )
    }
    worksheet
      .getRange(
        bounds.startRow,
        bounds.startColumn,
        bounds.endRow - bounds.startRow + 1,
        bounds.endColumn - bounds.startColumn + 1,
      )
      .activate()
    worksheet.scrollToCell(bounds.startRow, bounds.startColumn)
    return { ok: true, sheetName: worksheet.getSheetName() }
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : 'Selection failed',
    }
  }
}
