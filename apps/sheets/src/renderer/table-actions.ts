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
  if (!cellHasValue(worksheet, anchorRow, anchorColumn)) return null
  let startRow = anchorRow
  let endRow = anchorRow
  let startColumn = anchorColumn
  let endColumn = anchorColumn
  const lastRow = Math.max(anchorRow, worksheet.getLastRow())
  const lastColumn = Math.max(anchorColumn, worksheet.getLastColumn())
  let grew = true
  while (grew) {
    grew = false
    if (startRow > 0) {
      for (let c = startColumn; c <= endColumn; c += 1) {
        if (cellHasValue(worksheet, startRow - 1, c)) {
          startRow -= 1
          grew = true
          break
        }
      }
    }
    if (endRow < lastRow) {
      for (let c = startColumn; c <= endColumn; c += 1) {
        if (cellHasValue(worksheet, endRow + 1, c)) {
          endRow += 1
          grew = true
          break
        }
      }
    }
    if (startColumn > 0) {
      for (let r = startRow; r <= endRow; r += 1) {
        if (cellHasValue(worksheet, r, startColumn - 1)) {
          startColumn -= 1
          grew = true
          break
        }
      }
    }
    if (endColumn < lastColumn) {
      for (let r = startRow; r <= endRow; r += 1) {
        if (cellHasValue(worksheet, r, endColumn + 1)) {
          endColumn += 1
          grew = true
          break
        }
      }
    }
  }
  // Need at least header plus one data row
  if (endRow <= startRow) return null
  return { startRow, startColumn, endRow, endColumn }
}
