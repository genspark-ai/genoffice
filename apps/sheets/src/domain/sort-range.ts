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
      cells.push(state.value)
    }
    rows.push({ key: cells[keyColumn - bounds.startColumn] ?? null, cells })
  }

  // Stable sort; blanks stay last in both directions (Excel behavior).
  const sorted = rows
    .map((row, index) => ({ row, index }))
    .sort((a, b) => {
      const cmp = compareScalars(a.row.key, b.row.key)
      const oriented =
        a.row.key === null || b.row.key === null || cmp === 0 ? cmp : spec.ascending ? cmp : -cmp
      return oriented !== 0 ? oriented : a.index - b.index
    })
    .map((entry) => entry.row)

  const changes: SortComputedChange[] = []
  sorted.forEach((row, offset) => {
    const targetRow = firstDataRow + offset
    row.cells.forEach((value, columnOffset) => {
      const address = formatAddress(targetRow, bounds.startColumn + columnOffset)
      const before = rows[offset]?.cells[columnOffset] ?? null
      if (before !== value) changes.push({ address, before, after: value })
    })
  })
  return changes
}
