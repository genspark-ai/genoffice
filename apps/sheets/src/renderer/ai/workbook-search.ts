/**
 * Workbook-wide search and view navigation for the AI skill layer
 * (find_cells / select_range). Demo workbooks scan the in-memory snapshot;
 * lazy workbooks search the file in the sidecar (plain queries) or page
 * through readSheetRangeMapped (regex / errors-only) — structural ops
 * translated, session cell edits overlaid from the journal — so the search
 * never has to stream the whole workbook into Univer.
 */
import {
  formatAddress,
  parseAddress,
  type RangeBounds,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import type { CellScalar } from '@genoffice/xlsx-gateway/domain/workbook.types'
import { findSheetCellsInFile, sidecarFindAvailable } from '../sidecar-find'
import { ensureLazyRangeLoaded, readSheetRangeMapped } from '../univer-sync'
import type { LazyWorkbookState } from '../univer-state'
import { netAxisDelta } from '../view-transform'
import type {
  FindCellsMatch,
  FindCellsOptions,
  FindCellsOutcome,
  SelectRangeOutcome,
} from './tools'
import { redactionsOf, type WorkbookReadContext } from './workbook-readers'

/** Every error value Excel can display (ECMA-376 ST_CellType plus the modern
    data-type errors): Error Checking, find_cells errorsOnly, and formula
    audit all share this taxonomy so no evaluation error is silently skipped. */
export const ERROR_VALUE_RE =
  /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!|FIELD!|CONNECT!|BLOCKED!|UNKNOWN!|GETTING_DATA)$/
/** Total cells (by scanned extent) one range-paged find_cells call may
    cover; the sidecar path has no extent budget. */
export const MAX_SCAN_CELLS = 400_000
/** Indexing of a never-visited sheet may still be running; wait this long. */
const SIDECAR_INDEXING_WAIT_MS = 60_000
/** Row batches sized to stay under the sidecar's per-read cell budget. */
export const FILE_READ_BATCH_CELLS = 18_000

type Matcher = (value: CellScalar | undefined, formula: string | undefined) => boolean

/** Same text the sidecar and the Find dialog match against (booleans 1/0). */
function cellValueText(value: Exclude<CellScalar, null>): string {
  return typeof value === 'boolean' ? (value ? '1' : '0') : String(value)
}

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
        matchText(cellValueText(value))
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
  const redactions = redactionsOf(ctx)
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
      // A withheld cell is not searched at all. Returning a placeholder as a
      // match would still leak: "this value is in Customers!B2" is an oracle
      // the model can test guesses against, and the address alone narrows the
      // search. Withheld means absent from this tool's world.
      if (
        redactions.labelAt(sheet.id, parseAddress(address).row, parseAddress(address).column) !==
        null
      ) {
        continue
      }
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
  const redactions = redactionsOf(ctx)
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
      if (redactions.labelAt(sheetId, entry.row, entry.column) !== null) continue
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
    let sheetIncomplete = false
    let viaSidecar = false
    if (!options.regex && !options.errorsOnly && sidecarFindAvailable()) {
      try {
        const outcome = await findSheetCellsInFile(
          state,
          sheetId,
          {
            query: options.query,
            matchCase: false,
            matchEntireCell: false,
            lookIn: options.lookIn === 'formulas' ? 'formulas_only' : options.lookIn,
            wildcards: false,
          },
          {
            maxMatches: options.maxResults + 1,
            alive: () => !truncated,
            indexingWaitLimitMs: SIDECAR_INDEXING_WAIT_MS,
            onPage: (hits) => {
              let kept = 0
              for (const hit of hits) {
                if (shadowed.has(`${hit.row}:${hit.column}`)) continue
                // The sidecar reads the file, so `shadowed` — which only holds
                // cells edited this session — says nothing about a cell the
                // reader hid weeks ago. Without this the withheld value comes
                // back through the one path that never consults the index.
                if (redactions.labelAt(sheetId, hit.row, hit.column) !== null) continue
                if (
                  !push({
                    sheetName,
                    address: formatAddress(hit.row, hit.column),
                    value: hit.value,
                    formula: hit.formula,
                  })
                ) {
                  break
                }
                kept += 1
              }
              return kept
            },
          },
        )
        viaSidecar = true
        if (!outcome.complete && !outcome.capped && !truncated) sheetIncomplete = true
      } catch {
        /* fall through to the range-paged scan */
      }
    }
    if (!viaSidecar) {
      const ops = state.editJournal.structuralOps.get(sheetId) ?? []
      const screenRows = meta.rowCount + netAxisDelta(ops, 'row')
      const screenColumns = meta.columnCount + netAxisDelta(ops, 'column')
      if (screenRows <= 0 || screenColumns <= 0) continue
      const batchRows = Math.max(1, Math.floor(FILE_READ_BATCH_CELLS / screenColumns))
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
          // a withheld cell never matches at all — the address alone is an oracle
          if (redactions.labelAt(sheetId, cell.row, cell.column) !== null) continue
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
