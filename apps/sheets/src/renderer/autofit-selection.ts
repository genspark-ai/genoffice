import type { IRange } from '@univerjs/core'

/// Excel's Home > Format > AutoFit sizes every row (or column) the selection
/// touches by all of that row's cells, not only the selected ones.
export function fullRowSpans(selections: readonly IRange[], columnCount: number): IRange[] {
  return selections.map((range) => ({
    ...range,
    startColumn: 0,
    endColumn: Math.max(0, columnCount - 1),
  }))
}

export function fullColumnSpans(selections: readonly IRange[], rowCount: number): IRange[] {
  return selections.map((range) => ({
    ...range,
    startRow: 0,
    endRow: Math.max(0, rowCount - 1),
  }))
}
