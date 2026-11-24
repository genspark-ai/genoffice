import {
  columnLabel,
  parseAddress,
  type RangeBounds,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import type { ChartGridValue } from '@genoffice/xlsx-gateway/domain/chart-visual'

/// The grid a named-columns chart reads, fetched one column at a time: the
/// cell limit counts only the named columns, so a read of the whole rectangle
/// would refuse a range the plan accepted. Unnamed columns stay holes.
export async function readNamedChartColumns(
  bounds: RangeBounds,
  columns: readonly string[],
  readRange: (range: string) => Promise<ChartGridValue[][]>,
): Promise<ChartGridValue[][]> {
  const rows = bounds.endRow - bounds.startRow + 1
  const grid: ChartGridValue[][] = Array.from({ length: rows }, () => [])
  const seen = new Set<number>()
  for (const letter of columns) {
    const column = parseAddress(`${letter.toUpperCase()}1`).column
    if (column < bounds.startColumn || column > bounds.endColumn || seen.has(column)) continue
    seen.add(column)
    const label = columnLabel(column)
    const values = await readRange(`${label}${bounds.startRow + 1}:${label}${bounds.endRow + 1}`)
    for (let row = 0; row < rows; row++) {
      grid[row]![column - bounds.startColumn] = values[row]?.[0]
    }
  }
  return grid
}
