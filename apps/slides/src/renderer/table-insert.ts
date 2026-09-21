/**
 * Insert > Table geometry (PowerPoint for Mac, blank 16:9 layout, 5 x 2 default): the table is
 * 2/3 of the slide wide and centered, its top sits at 10.5% of the slide height (56.67 pt on a
 * 540 pt slide - not vertically centered), and every row is 29.2 pt. Extra rows overflow the
 * bottom edge exactly as PowerPoint lets them.
 */
import { PX_PER_INCH } from './app-constants'

const EMU_PER_PT = 12700
export const TABLE_ROW_HEIGHT_PT = 29.2
export const TABLE_ROW_HEIGHT_EMU = Math.round(TABLE_ROW_HEIGHT_PT * EMU_PER_PT)
export const TABLE_ROW_HEIGHT_PX = (TABLE_ROW_HEIGHT_PT * PX_PER_INCH) / 72
export const TABLE_WIDTH_FRACTION = 2 / 3
export const TABLE_TOP_FRACTION = 0.105

export interface TableInsertSpec {
  x: number
  y: number
  w: number
  h: number
  rowHeightEmu: number
}

/** Fractional px are intentional: the main process rounds once, to EMU. */
export function tableInsertSpec(
  slide: { widthPx: number; heightPx: number },
  rows: number,
): TableInsertSpec {
  const w = slide.widthPx * TABLE_WIDTH_FRACTION
  return {
    x: (slide.widthPx - w) / 2,
    y: slide.heightPx * TABLE_TOP_FRACTION,
    w,
    h: rows * TABLE_ROW_HEIGHT_PX,
    rowHeightEmu: TABLE_ROW_HEIGHT_EMU,
  }
}
