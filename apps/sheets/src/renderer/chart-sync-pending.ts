import type { CellBounds } from '../domain/chart-visual'

export interface FormulaCellLike {
  readonly f?: string | null | undefined | void
  readonly si?: string | null | undefined | void
  readonly v?: unknown
}

/// A formula cell the engine has not evaluated yet carries `f`/`si` but no
/// `v` (evaluated empties are coerced to 0 by formula-null-result). Baking
/// such a vector into a chart cache would store 0 for every pending point.
export function hasPendingFormulaCells(
  cells: readonly (FormulaCellLike | null | undefined | void)[],
): boolean {
  return cells.some((cell) => cell != null && (cell.f != null || cell.si != null) && cell.v == null)
}

export type ChartGridValue = string | number | boolean | null | undefined

export interface StreamedCellLike {
  readonly row: number
  readonly column: number
  readonly value: ChartGridValue
  readonly formula?: string | undefined
}

export interface StreamedJournalEntryLike extends StreamedCellLike {
  readonly hasValue: boolean
}

export interface StreamedChartGrid {
  readonly values: ChartGridValue[][]
  /// Some cell in the range is a formula with no value to read: a journaled
  /// formula (the journal never receives engine results) or a sidecar
  /// formula saved without a cached `<v>`.
  readonly pendingFormula: boolean
}

/// Streaming-mode chart range: journal edits win, then sidecar-mapped screen
/// cells, then bulk constant fills.
export function buildStreamedChartGrid(
  bounds: CellBounds,
  screenCells: readonly StreamedCellLike[],
  journalEntries: readonly StreamedJournalEntryLike[],
  fillAt: (
    row: number,
    column: number,
  ) => { found: true; value: ChartGridValue } | { found: false },
): StreamedChartGrid {
  const cells = new Map<string, ChartGridValue>()
  const pending = new Set<string>()
  const put = (cell: StreamedCellLike): void => {
    const key = `${cell.row}:${cell.column}`
    cells.set(key, cell.value)
    if (cell.formula !== undefined && cell.value == null) pending.add(key)
    else pending.delete(key)
  }
  for (const cell of screenCells) put(cell)
  for (const entry of journalEntries) if (entry.hasValue) put(entry)
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      const fill = fillAt(row, column)
      if (fill.found) put({ row, column, value: fill.value })
    }
  }
  const values: ChartGridValue[][] = []
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    const line: ChartGridValue[] = []
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      line.push(cells.get(`${row}:${column}`))
    }
    values.push(line)
  }
  return { values, pendingFormula: pending.size > 0 }
}

export interface ChartSeriesRefsLike {
  readonly valuesRef?: string | undefined
  readonly categoriesRef?: string | undefined
  readonly values: readonly unknown[]
  readonly categories: readonly unknown[]
}

/// Touched refs whose file vector is empty (no strCache/numCache): the
/// renderer hydrates those from the grid, so they are re-read, never baked.
export function cacheLessRefs(
  series: ChartSeriesRefsLike,
  touched: (ref: string | undefined) => boolean,
): string[] {
  const refs: string[] = []
  if (series.valuesRef && series.values.length === 0 && touched(series.valuesRef)) {
    refs.push(series.valuesRef)
  }
  if (series.categoriesRef && series.categories.length === 0 && touched(series.categoriesRef)) {
    refs.push(series.categoriesRef)
  }
  return refs
}

export function unionBounds(a: CellBounds | undefined, b: CellBounds): CellBounds {
  if (!a) return b
  return {
    startRow: Math.min(a.startRow, b.startRow),
    endRow: Math.max(a.endRow, b.endRow),
    startColumn: Math.min(a.startColumn, b.startColumn),
    endColumn: Math.max(a.endColumn, b.endColumn),
  }
}
