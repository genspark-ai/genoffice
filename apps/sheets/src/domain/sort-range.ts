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
  if (typeof a === 'number' && typeof b === 'number') return a - b
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

  const rows: { key: CellScalar; cells: CellScalar[] }[] = []
  for (let row = firstDataRow; row <= bounds.endRow; row += 1) {
    const cells: CellScalar[] = []
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
    }
    rows.push({ key: cells[keyColumn - bounds.startColumn] ?? null, cells })
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
      if (before !== value) changes.push({ address, before, after: value })
    })
  })
  return changes
}
