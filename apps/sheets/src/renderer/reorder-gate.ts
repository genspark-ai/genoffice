import type { IRange } from '@univerjs/core'

import { rangeCellCount } from '@genoffice/xlsx-gateway/domain/cell-address'

import { REORDER_JOURNAL_MAX_CELLS } from './app-constants'
import { fitsFullLoad, type WorkbookCellCounts } from './load-budget'
import type { LazyWorkbookState } from './univer-state'
import { fileToScreen } from './view-transform'

export type ReorderGateVerdict =
  'allow' | 'loading' | 'offerFullLoad' | 'workbookTooLarge' | 'valueModeFormulas' | 'rangeTooLarge'

export interface ReorderGateInput {
  readonly formulaMode: boolean
  readonly preloadComplete: boolean
  readonly preloadRunning: boolean
  readonly isAddedSheet: boolean
  readonly cellCounts: WorkbookCellCounts
  readonly rangeCells: number
  readonly rangeHasStreamedFormulas: () => boolean
}

export type FullLoadGateVerdict = 'loading' | 'offerFullLoad' | 'workbookTooLarge'

/// What an action that needs the complete data gets while the workbook is
/// still streamed: formula-mode books preload themselves at open, and any
/// running preload (the Load-all prompt, Print / PDF export, headless export)
/// shows the in-progress notice instead of a second offer; otherwise the
/// offer, or the explanation when the book is over the load budget.
export function fullLoadGate(
  input: Pick<ReorderGateInput, 'formulaMode' | 'preloadRunning' | 'cellCounts'>,
): FullLoadGateVerdict {
  if (input.preloadRunning || input.formulaMode) return 'loading'
  return fitsFullLoad(input.cellCounts) ? 'offerFullLoad' : 'workbookTooLarge'
}

/// Sort / move-range / shift / split-text rewrite model content, so every
/// cell they touch must be resident. After Full Load that holds in both
/// modes; value mode additionally keeps formulas sidecar-only, so moving
/// such a cell would save its cached result as a constant.
export function reorderGate(input: ReorderGateInput): ReorderGateVerdict {
  if (input.isAddedSheet) return 'allow'
  if (!input.preloadComplete) return fullLoadGate(input)
  if (input.rangeCells > REORDER_JOURNAL_MAX_CELLS) return 'rangeTooLarge'
  if (!input.formulaMode && input.rangeHasStreamedFormulas()) return 'valueModeFormulas'
  return 'allow'
}

export function largestRangeCells(ranges: readonly IRange[]): number {
  return Math.max(0, ...ranges.map(rangeCellCount))
}

/// Sort / split carry `range`; move-range and move-rows carry from/to (both
/// ends are journaled, so both are checked). Univer's own shift menu items
/// dispatch without params and read the selection, so the selection stands
/// in; a shift then rewrites everything from the range to the sheet edge.
/// With nothing to go on the whole sheet is assumed, failing closed.
export function reorderCommandRanges(
  commandId: string,
  params: unknown,
  selection: IRange | undefined,
  extent: { rows: number; columns: number } | null,
): IRange[] {
  const { range, fromRange, toRange } = (
    typeof params === 'object' && params !== null ? params : {}
  ) as { range?: IRange; fromRange?: IRange; toRange?: IRange }
  let ranges = [range, fromRange, toRange].filter((candidate): candidate is IRange =>
    isRange(candidate),
  )
  if (ranges.length === 0 && isRange(selection)) ranges = [selection]
  if (ranges.length === 0) {
    return extent
      ? [{ startRow: 0, endRow: extent.rows - 1, startColumn: 0, endColumn: extent.columns - 1 }]
      : []
  }
  const shift = /-move-(left|right|up|down)$/.exec(commandId)?.[1]
  if (!shift || !extent) return ranges
  return ranges.map((candidate) =>
    shift === 'left' || shift === 'right'
      ? { ...candidate, endColumn: Math.max(candidate.endColumn, extent.columns - 1) }
      : { ...candidate, endRow: Math.max(candidate.endRow, extent.rows - 1) },
  )
}

function isRange(value: unknown): value is IRange {
  if (typeof value !== 'object' || value === null) return false
  const { startRow, endRow, startColumn, endColumn } = value as Record<string, unknown>
  return [startRow, endRow, startColumn, endColumn].every((n) => typeof n === 'number')
}

/// formulaText is keyed in file coordinates; the ranges arrive in screen
/// coordinates, so each stored formula is mapped through the session's
/// structural ops before the containment test.
export function rangeHasStreamedFormulas(
  state: Pick<
    LazyWorkbookState,
    'formulaText' | 'formulaTextTruncated' | 'sharedFormulaGroups' | 'editJournal'
  >,
  sheetId: string,
  ranges: readonly IRange[],
): boolean {
  if (ranges.length === 0) return false
  // An index that overflowed and was dropped cannot prove the range clean.
  if (state.formulaTextTruncated.has(sheetId)) return true
  const ops = state.editJournal.structuralOps.get(sheetId) ?? []
  const inRanges = (fileRow: number, fileColumn: number): boolean => {
    const row = ops.length === 0 ? fileRow : fileToScreen(ops, 'row', fileRow)
    const column = ops.length === 0 ? fileColumn : fileToScreen(ops, 'column', fileColumn)
    if (row === null || column === null) return false
    return ranges.some(
      (range) =>
        row >= range.startRow &&
        row <= range.endRow &&
        column >= range.startColumn &&
        column <= range.endColumn,
    )
  }
  for (const key of state.formulaText.get(sheetId)?.keys() ?? []) {
    const separator = key.indexOf(':')
    if (inRanges(Number(key.slice(0, separator)), Number(key.slice(separator + 1)))) return true
  }
  let hit = false
  state.sharedFormulaGroups.get(sheetId)?.forEachFollower((row, column) => {
    hit ||= inRanges(row, column)
  })
  return hit
}
