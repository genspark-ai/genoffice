import type { LazyWorkbookState } from './univer-state'

interface LargeFilterSheet {
  /** The evaluated union in physical sheet coordinates. The grid and the
   * viewport loader share it rather than making a copy on every scroll. */
  visibility: { readonly filteredOutRows: Set<number> } | undefined
  pending: Promise<void> | null
  generation: number
}

const sheets = new WeakMap<LazyWorkbookState, Map<string, LargeFilterSheet>>()

export function largeFilterSheet(state: LazyWorkbookState, sheetId: string): LargeFilterSheet {
  let perSheet = sheets.get(state)
  if (!perSheet) sheets.set(state, (perSheet = new Map()))
  let entry = perSheet.get(sheetId)
  if (!entry)
    perSheet.set(sheetId, (entry = { visibility: undefined, pending: null, generation: 0 }))
  return entry
}

export function resetLargeFilter(state: LazyWorkbookState, sheetId: string): void {
  const entry = sheets.get(state)?.get(sheetId)
  if (!entry) return
  entry.visibility = undefined
  entry.generation += 1
}

export function largeFilterOwnsSpan(state: LazyWorkbookState, sheetId: string): boolean {
  return largeFilterHiddenRows(state, sheetId) !== undefined
}

export function largeFilterHiddenRows(
  state: LazyWorkbookState,
  sheetId: string,
): ReadonlySet<number> | undefined {
  return sheets.get(state)?.get(sheetId)?.visibility?.filteredOutRows
}
