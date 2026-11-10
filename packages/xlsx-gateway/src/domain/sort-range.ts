import { columnIndex, formatAddress, parseRange } from './cell-address'
import type { CellScalar, CellState } from './workbook.types'

/**
 * Computes the cell rewrites a sort produces, so sorting rides the existing
 * per-cell preview / CAS / apply machinery in both demo and lazy mode.
 * Regions containing formulas are rejected: moving formula text between rows
 * silently re-targets relative references, which is exactly the class of
 * damage the preview model exists to prevent.
 */

export interface SortComputedChange {
  readonly address: string
  readonly before: CellScalar
  readonly after: CellScalar
  /// The cell's display text before the sort. `after` is the raw model value,
  /// but the CAS precondition is compared against the cell's `value`, so a
  /// formatted cell (currency, date) must be guarded on this, not on `before`.
  readonly expectedValue: CellScalar
}

export interface SortSpec {
  readonly range: string
  readonly byColumn: string
  readonly ascending: boolean
  readonly hasHeader: boolean
}

/// blanks always sort last; numbers before booleans before text (Excel order)
function compareScalars(a: CellScalar, b: CellScalar): number {
  if (a === null && b === null) return 0
  if (a === null) return 1
  if (b === null) return -1
  const rank = (value: Exclude<CellScalar, null>): number =>
    typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2
  const rankA = rank(a)
  const rankB = rank(b)
  if (rankA !== rankB) return rankA - rankB
  // A subtraction comparator returns NaN for NaN/Infinity-Infinity inputs,
  // which makes Array.sort nondeterministic. Order non-finite numbers
  // (error values) after finite ones instead.
  if (typeof a === 'number' && typeof b === 'number') {
    const aFinite = Number.isFinite(a)
    const bFinite = Number.isFinite(b)
    if (aFinite && bFinite) return a < b ? -1 : a > b ? 1 : 0
    if (aFinite) return -1
    if (bFinite) return 1
    return 0
  }
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b)
  return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' })
}

/// Stable data-row order for a sort: returns source-row indexes in their
/// sorted sequence (order[target] = source). Blanks stay last in both
/// directions (Excel behavior). Shared by the per-cell expansion path and the
/// range-level bulk executor so both sort identically.
export function computeSortedRowOrder(
  rows: readonly (readonly CellScalar[])[],
  keyOffset: number,
  ascending: boolean,
): number[] {
  return rows
    .map((row, index) => ({ key: row[keyOffset] ?? null, index }))
    .sort((a, b) => {
      const cmp = compareScalars(a.key, b.key)
      const oriented = a.key === null || b.key === null || cmp === 0 ? cmp : ascending ? cmp : -cmp
      return oriented !== 0 ? oriented : a.index - b.index
    })
    .map((entry) => entry.index)
}

export function computeSortChanges(
  spec: SortSpec,
  readCell: (address: string) => CellState,
): SortComputedChange[] {
  const bounds = parseRange(spec.range)
  const keyColumn = columnIndex(spec.byColumn)
  if (keyColumn < bounds.startColumn || keyColumn > bounds.endColumn) {
    throw new Error(`Sort column ${spec.byColumn} is outside the range ${spec.range}.`)
  }
  const firstDataRow = bounds.startRow + (spec.hasHeader ? 1 : 0)
  if (firstDataRow >= bounds.endRow) {
    throw new Error('The sort range needs at least two data rows.')
  }

  const rows: { key: CellScalar; cells: CellScalar[]; display: CellScalar[] }[] = []
  for (let row = firstDataRow; row <= bounds.endRow; row += 1) {
    const cells: CellScalar[] = []
    const display: CellScalar[] = []
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      const state = readCell(formatAddress(row, column))
      if (state.formula) {
        throw new Error(
          `The sort range contains a formula at ${formatAddress(row, column)} — sorting would silently re-target its references. Sort values only, or convert formulas to values first.`,
        )
      }
      // Raw model values: `value` is display text, so formatted numbers and
      // dates would sort lexicographically AND be rewritten as text by the
      // moves. Raw serials sort numerically — Excel's order.
      cells.push(state.rawValue !== undefined ? state.rawValue : state.value)
      display.push(state.value)
    }
    rows.push({ key: cells[keyColumn - bounds.startColumn] ?? null, cells, display })
  }

  const order = computeSortedRowOrder(
    rows.map((row) => row.cells),
    keyColumn - bounds.startColumn,
    spec.ascending,
  )

  const changes: SortComputedChange[] = []
  order.forEach((sourceIndex, offset) => {
    const targetRow = firstDataRow + offset
    rows[sourceIndex]?.cells.forEach((value, columnOffset) => {
      const address = formatAddress(targetRow, bounds.startColumn + columnOffset)
      const before = rows[offset]?.cells[columnOffset] ?? null
      if (before !== value) {
        changes.push({
          address,
          before,
          after: value,
          expectedValue: rows[offset]?.display[columnOffset] ?? null,
        })
      }
    })
  })
  return changes
}
