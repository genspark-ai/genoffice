import type { UniverWorksheet } from './univer-state'

export interface TableRegion {
  readonly startRow: number
  readonly startColumn: number
  readonly endRow: number
  readonly endColumn: number
}

function cellHasValue(worksheet: UniverWorksheet, row: number, column: number): boolean {
  try {
    const cell = worksheet.getRange(row, column, 1, 1)
    // getValue returns the raw value (null for empty); getDisplayValue covers formatted empties
    const raw = (cell as unknown as { getValue?: () => unknown }).getValue?.()
    if (raw !== null && raw !== undefined && raw !== '') {
      if (typeof raw === 'object') {
        const v = (raw as { v?: unknown }).v
        if (v !== undefined) return v !== null && v !== ''
      } else {
        return true
      }
    }
    const display = (
      cell as unknown as { getDisplayValue?: () => string | null }
    ).getDisplayValue?.()
    return display !== null && display !== undefined && display !== ''
  } catch {
    return false
  }
}

/**
 * Infer the continuous data region (Excel CurrentRegion) around the anchor
 * cell. Expands outward while any cell on the adjacent edge has a value,
 * bounded by the sheet's used extent. Returns null when the anchor itself
 * is empty or the region would be a single cell.
 */
export function inferContinuousRegion(
  worksheet: UniverWorksheet,
  anchorRow: number,
  anchorColumn: number,
): TableRegion | null {
  const region = expandContiguousRegion(
    (row, column) => cellHasValue(worksheet, row, column),
    { row: anchorRow, column: anchorColumn },
    { row: worksheet.getLastRow(), column: worksheet.getLastColumn() },
  )
  // Need at least header plus one data row
  if (!region || region.endRow <= region.startRow) return null
  return region
}

/// CurrentRegion over a value predicate; null when the anchor is empty.
export function expandContiguousRegion(
  hasValue: (row: number, column: number) => boolean,
  anchor: { row: number; column: number },
  last: { row: number; column: number },
): TableRegion | null {
  if (!hasValue(anchor.row, anchor.column)) return null
  let startRow = anchor.row
  let endRow = anchor.row
  let startColumn = anchor.column
  let endColumn = anchor.column
  const lastRow = Math.max(anchor.row, last.row)
  const lastColumn = Math.max(anchor.column, last.column)
  const rowHasValue = (row: number): boolean => {
    for (let c = startColumn; c <= endColumn; c += 1) if (hasValue(row, c)) return true
    return false
  }
  const columnHasValue = (column: number): boolean => {
    for (let r = startRow; r <= endRow; r += 1) if (hasValue(r, column)) return true
    return false
  }
  let grew = true
  while (grew) {
    grew = false
    if (startRow > 0 && rowHasValue(startRow - 1)) {
      startRow -= 1
      grew = true
    }
    if (endRow < lastRow && rowHasValue(endRow + 1)) {
      endRow += 1
      grew = true
    }
    if (startColumn > 0 && columnHasValue(startColumn - 1)) {
      startColumn -= 1
      grew = true
    }
    if (endColumn < lastColumn && columnHasValue(endColumn + 1)) {
      endColumn += 1
      grew = true
    }
  }
  return { startRow, startColumn, endRow, endColumn }
}
