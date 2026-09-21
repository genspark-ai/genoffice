/**
 * Real-file status bar statistics for streamed workbooks.
 *
 * Univer's StatusBarController computes the footer statistics (count, sum,
 * min/max/average) from the in-memory cell matrix. On a large-file streamed
 * workbook the grid only materializes the regions the user has viewed, so
 * selecting a whole column of a 185k-row file shows a count of the few
 * loaded windows (user report: count 191 on a fully populated column) —
 * the number looks simply wrong.
 *
 * This module listens to the same status bar state the controller publishes
 * and, when the active sheet streams from a file, recomputes the statistics
 * from the file itself through the aggregate_range channel (batched sidecar
 * reads overlaid with this session's edit journal — never the formula
 * engine, never a grid load), then overwrites the published values. When the
 * selection cannot be aggregated (sheet added this session, structural
 * row/column changes pending save, selection beyond the cell budget), the
 * upstream numbers are left untouched.
 */
import { RANGE_TYPE } from '@univerjs/core'
import type { IRange } from '@univerjs/core'
import { FUNCTION_NAMES_MATH, FUNCTION_NAMES_STATISTICAL } from '@univerjs/engine-formula'
import { SheetsSelectionsService } from '@univerjs/sheets'
import { IStatusBarService } from '@univerjs/sheets-ui'
import type { RangeBounds } from '@genoffice/xlsx-gateway/domain/cell-address'

import type { RangeAggregate } from './ai/aggregate'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/** Same per-call budget as the aggregate_range AI tool (post-clamp cells). */
const MAX_STATUSBAR_AGGREGATE_CELLS = 1_000_000
/** Collapses selection-drag bursts into one file read. */
const RECOMPUTE_DEBOUNCE_MS = 150
/** Each visible-row band is one aggregate round trip; past this the upstream numbers stay. */
const MAX_VISIBLE_BANDS = 64

interface StatusBarValue {
  func: string
  value: number
}

interface StatusBarState {
  values: StatusBarValue[]
  pattern: string | null
}

interface StatusBarServiceLike {
  state$: {
    subscribe(next: (state: StatusBarState | null) => void): { unsubscribe(): void }
  }
  setState(state: StatusBarState | null): void
}

export interface StatusBarFileStatsDeps {
  runtime: UniverRuntime
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  aggregate: (
    sheetId: string,
    bounds: RangeBounds,
  ) => Promise<{ ok: true; aggregate: RangeAggregate } | { ok: false; error: string }>
}

/** Whole-row/column/sheet selections carry sentinel edges; expand to grid. */
function selectionBounds(range: IRange, maxRows: number, maxColumns: number): RangeBounds {
  const rows = range.rangeType === RANGE_TYPE.COLUMN || range.rangeType === RANGE_TYPE.ALL
  const columns = range.rangeType === RANGE_TYPE.ROW || range.rangeType === RANGE_TYPE.ALL
  return {
    startRow: rows ? 0 : range.startRow,
    endRow: rows ? maxRows - 1 : range.endRow,
    startColumn: columns ? 0 : range.startColumn,
    endColumn: columns ? maxColumns - 1 : range.endColumn,
  }
}

/**
 * Univer's status bar statistics skip hidden and filtered-out rows by splitting
 * each selection into visible-row bands; mirror that so an active AutoFilter
 * keeps its filtered count. null = too many bands to aggregate one by one.
 */
export function visibleRowBands(
  bounds: RangeBounds,
  isRowVisible: (row: number) => boolean,
  maxBands = MAX_VISIBLE_BANDS,
): RangeBounds[] | null {
  const bands: RangeBounds[] = []
  let start: number | null = null
  for (let row = bounds.startRow; row <= bounds.endRow + 1; row += 1) {
    const visible = row <= bounds.endRow && isRowVisible(row)
    if (visible) {
      if (start === null) start = row
    } else if (start !== null) {
      bands.push({ ...bounds, startRow: start, endRow: row - 1 })
      if (bands.length > maxBands) return null
      start = null
    }
  }
  return bands
}

function combine(parts: readonly RangeAggregate[]): {
  nonEmpty: number
  numericCount: number
  sum: number
  min: number | null
  max: number | null
} {
  let nonEmpty = 0
  let numericCount = 0
  let sum = 0
  let min: number | null = null
  let max: number | null = null
  for (const part of parts) {
    nonEmpty += part.nonEmpty
    numericCount += part.numericCount
    sum += part.sum
    if (part.min !== null) min = min === null ? part.min : Math.min(min, part.min)
    if (part.max !== null) max = max === null ? part.max : Math.max(max, part.max)
  }
  return { nonEmpty, numericCount, sum, min, max }
}

export function installStatusBarFileStats(deps: StatusBarFileStatsDeps): { dispose(): void } {
  const { runtime, lazyWorkbookRef, aggregate } = deps
  const injector = runtime.univer.__getInjector()
  const statusBarService = injector.get(IStatusBarService) as unknown as StatusBarServiceLike
  const selectionsService = injector.get(SheetsSelectionsService)

  let applying = false
  let seq = 0
  let timer: ReturnType<typeof setTimeout> | null = null

  const selectionSignature = (): string =>
    JSON.stringify(
      (selectionsService.getCurrentSelections() ?? []).map((selection) => selection.range),
    )

  const recompute = async (incomingPattern: string | null): Promise<void> => {
    const token = ++seq
    const state = lazyWorkbookRef.current
    // Only cached-value streaming mode shows a partial matrix; formula-mode
    // workbooks are fully loaded (small files) or preloading toward it.
    if (!state || state.formulaMode) return
    const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!worksheet) return
    const sheetId = worksheet.getSheetId()
    const sheetMeta = state.file.sheets.find((candidate) => candidate.id === sheetId)
    // Sheets added or duplicated this session live fully in the grid — the
    // upstream numbers are already right.
    if (!sheetMeta) return
    // aggregate_range refuses after structural changes; skip the round trip.
    if ((state.editJournal.structuralOps.get(sheetId) ?? []).length > 0) return
    const selections = selectionsService.getCurrentSelections()
    if (!selections || selections.length === 0) return

    const maxRows = worksheet.getMaxRows()
    const maxColumns = worksheet.getMaxColumns()
    const sheet = worksheet.getSheet()
    const boundsList: RangeBounds[] = []
    for (const selection of selections) {
      const bands = visibleRowBands(selectionBounds(selection.range, maxRows, maxColumns), (row) =>
        sheet.getRowVisible(row),
      )
      if (!bands) return
      boundsList.push(...bands)
    }
    // Budget on file-extent cells (the aggregate clamps the same way): a
    // whole-column pick on a 200k-row sheet passes; select-all on a very
    // wide sheet keeps the upstream (loaded-window) numbers.
    const clampedCells = boundsList.reduce((total, bounds) => {
      const endRow = Math.min(bounds.endRow, sheetMeta.rowCount - 1)
      const endColumn = Math.min(bounds.endColumn, sheetMeta.columnCount - 1)
      if (endRow < bounds.startRow || endColumn < bounds.startColumn) return total
      return total + (endRow - bounds.startRow + 1) * (endColumn - bounds.startColumn + 1)
    }, 0)
    if (clampedCells === 0 || clampedCells > MAX_STATUSBAR_AGGREGATE_CELLS) return

    const signature = selectionSignature()
    const parts: RangeAggregate[] = []
    for (const bounds of boundsList) {
      const outcome = await aggregate(sheetId, bounds)
      if (!outcome.ok) return
      parts.push(outcome.aggregate)
    }
    // A newer selection/state event supersedes this run.
    if (token !== seq || lazyWorkbookRef.current !== state) return
    if (runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId() !== sheetId) return
    if (selectionSignature() !== signature) return

    const totals = combine(parts)
    const values: StatusBarValue[] = [
      { func: FUNCTION_NAMES_STATISTICAL.COUNTA, value: totals.nonEmpty },
      { func: FUNCTION_NAMES_STATISTICAL.COUNT, value: totals.numericCount },
      { func: FUNCTION_NAMES_MATH.SUM, value: totals.sum },
    ]
    if (totals.min !== null)
      values.push({ func: FUNCTION_NAMES_STATISTICAL.MIN, value: totals.min })
    if (totals.max !== null)
      values.push({ func: FUNCTION_NAMES_STATISTICAL.MAX, value: totals.max })
    if (totals.numericCount > 0) {
      values.push({
        func: FUNCTION_NAMES_STATISTICAL.AVERAGE,
        value: totals.sum / totals.numericCount,
      })
    }
    applying = true
    try {
      statusBarService.setState({ values, pattern: incomingPattern })
    } finally {
      applying = false
    }
  }

  const subscription = statusBarService.state$.subscribe((incoming) => {
    if (applying) return
    if (timer) clearTimeout(timer)
    timer = null
    seq += 1
    // Univer cleared the footer (single cell, no selection): nothing to overwrite
    if (!incoming || incoming.values.length === 0) return
    timer = setTimeout(() => {
      timer = null
      void recompute(incoming.pattern)
    }, RECOMPUTE_DEBOUNCE_MS)
  })

  return {
    dispose() {
      if (timer) clearTimeout(timer)
      subscription.unsubscribe()
    },
  }
}
