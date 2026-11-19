type Cell = string | number | boolean | null | undefined

/**
 * The category vector Excel caches for a chart data range. A range that spans
 * both rows and columns is a level hierarchy, and a plain strRef cache holds
 * only the innermost level: the last row for row-oriented series, the last
 * column for column-oriented ones. Flattening it row-major would double the
 * category count and shift every label off its data point.
 */
export type SeriesOrientation = 'row' | 'column'

export function chartCategoryVector(
  grid: readonly (readonly Cell[])[],
  cachedCount: number,
  orientation?: SeriesOrientation,
): Cell[] {
  const rows = grid.length
  const columns = grid[0]?.length ?? 0
  if (rows <= 1 || columns <= 1 || cachedCount === rows * columns) return grid.flat()
  const lastRow = [...(grid[rows - 1] ?? [])]
  const lastColumn = grid.map((line) => line[line.length - 1])
  if (orientation === 'row') return lastRow
  if (orientation === 'column') return lastColumn
  if (cachedCount === columns && cachedCount !== rows) return lastRow
  if (cachedCount === rows && cachedCount !== columns) return lastColumn
  return columns >= rows ? lastRow : lastColumn
}

/** Which way a value reference runs; undefined for 2-D or unparseable refs. */
export function seriesOrientation(bounds: {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}): SeriesOrientation | undefined {
  if (bounds.startRow === bounds.endRow && bounds.startColumn !== bounds.endColumn) return 'row'
  if (bounds.startColumn === bounds.endColumn && bounds.startRow !== bounds.endRow) return 'column'
  return undefined
}
