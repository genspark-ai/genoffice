/**
 * Per-sheet state of "large-sheet filtering": filters evaluated over column
 * projections read from the sidecar instead of over the Univer model, for
 * streamed workbooks past the Full Load budget. Kept apart from large-filter.ts
 * so univer-sync can consult it without an import cycle.
 */
import type { LazyWorkbookState } from './univer-state'

export type ProjectionValue = string | number | boolean | null

export interface ColumnProjection {
  /// Screen row of values[0]; index = row - startRow.
  readonly startRow: number
  readonly values: ProjectionValue[]
  readonly styleIndex: (number | undefined)[]
}

export interface LargeFilterSheet {
  readonly projections: Map<number, ColumnProjection>
  /// Structural op count the projections were mapped under; a change
  /// invalidates them (screen rows moved).
  opsStamp: number
  /// Rows each criteria column hides, in screen coordinates.
  readonly hiddenByColumn: Map<number, Set<number>>
  /// Once a filter has been evaluated here, the filter model owns every row
  /// in the filter span: file-declared hidden rows inside it must not be
  /// re-hidden as blocks stream in.
  active: boolean
  /// In-flight evaluation, so a second apply waits instead of racing.
  pending: Promise<void> | null
  /// Bumped when the AutoFilter is removed; an apply still in flight then
  /// discards its result instead of reviving the filter.
  generation: number
}

const sheets = new WeakMap<LazyWorkbookState, Map<string, LargeFilterSheet>>()

export function largeFilterSheet(state: LazyWorkbookState, sheetId: string): LargeFilterSheet {
  let perSheet = sheets.get(state)
  if (!perSheet) {
    perSheet = new Map()
    sheets.set(state, perSheet)
  }
  let entry = perSheet.get(sheetId)
  if (!entry) {
    entry = {
      projections: new Map(),
      opsStamp: state.editJournal.structuralOps.get(sheetId)?.length ?? 0,
      hiddenByColumn: new Map(),
      active: false,
      pending: null,
      generation: 0,
    }
    perSheet.set(sheetId, entry)
  }
  return entry
}

/// The AutoFilter was removed: the model no longer owns the span and hides
/// no rows. Projections stay; they describe the file, not the filter.
export function resetLargeFilter(state: LazyWorkbookState, sheetId: string): void {
  const entry = sheets.get(state)?.get(sheetId)
  if (!entry) return
  entry.active = false
  entry.hiddenByColumn.clear()
  entry.generation += 1
}

export function largeFilterOwnsSpan(state: LazyWorkbookState, sheetId: string): boolean {
  return sheets.get(state)?.get(sheetId)?.active === true
}

/// Union of the rows the evaluated filter hides, when this sheet is in
/// large-filter mode; the viewport loader budgets visible rows by it.
export function largeFilterHiddenRows(
  state: LazyWorkbookState,
  sheetId: string,
): ReadonlySet<number> | undefined {
  const entry = sheets.get(state)?.get(sheetId)
  if (!entry?.active) return undefined
  const union = new Set<number>()
  for (const rows of entry.hiddenByColumn.values()) for (const row of rows) union.add(row)
  return union
}
