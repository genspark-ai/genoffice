/**
 * Excel's row autofit measures wrapped text against a narrower budget than
 * its layout does: a line stays whole in layout up to roughly column - 7 px,
 * but the height measure already counts two lines once the text exceeds
 * column - 16 px. Calibrated on a Calibri 11 workbook (column 176 px): Excel
 * gives a 162.8 px string a two-line row while drawing it on one line, and
 * keeps 157.9 / 158.4 px strings at one line, so the measure budget lies in
 * (158.4, 162.8) px = column - 4 (Univer's own padding) - (9.2 .. 13.6).
 * Narrow the column only while the auto-height measure runs; the rendered
 * line breaks stay put.
 */
import { SpreadsheetSkeleton } from '@univerjs/engine-render'

/// Excel autofit budget (column - 16 px) beyond Univer's own 2+2 padding.
export const AUTOFIT_EXTRA_PX = 12

export interface AutofitSkeletonLike {
  _worksheetData: {
    columnData: Record<number, { w?: number } | undefined>
    defaultColumnWidth: number
  }
  calculateAutoHeightForCell(row: number, col: number): number | undefined
}

const patched = new WeakSet<AutofitSkeletonLike>()

export function installAutofitWrapBudget(
  proto: AutofitSkeletonLike = SpreadsheetSkeleton.prototype as unknown as AutofitSkeletonLike,
): void {
  if (patched.has(proto)) return
  patched.add(proto)
  const original = proto.calculateAutoHeightForCell
  proto.calculateAutoHeightForCell = function (this: AutofitSkeletonLike, row, col) {
    const { columnData, defaultColumnWidth } = this._worksheetData
    const entry = columnData[col]
    const width = entry?.w ?? defaultColumnWidth
    columnData[col] = { ...entry, w: width - AUTOFIT_EXTRA_PX }
    try {
      return original.call(this, row, col)
    } finally {
      if (entry === undefined) delete columnData[col]
      else columnData[col] = entry
    }
  }
}
