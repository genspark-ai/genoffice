import { describe, it, expect } from 'vitest'
import {
  TABLE_ROW_HEIGHT_EMU,
  TABLE_ROW_HEIGHT_PX,
  tableInsertSpec,
} from '../src/renderer/table-insert'

// 16:9 deck (12192000 x 6858000 EMU = 960 x 540 pt) laid out at FIT_WIDTH 1280: 1 px = 9525 EMU
const SLIDE = { widthPx: 1280, heightPx: 720 }
const EMU_PER_PX = 9525
const EMU_PER_PT = 12700
const toEmu = (px: number) => Math.round(px * EMU_PER_PX)
const toPt = (emu: number) => emu / EMU_PER_PT

describe('Insert > Table default geometry (PowerPoint for Mac measurements, 5 x 2 dialog default)', () => {
  const spec = tableInsertSpec(SLIDE, 2)

  it('is 2/3 of the slide wide and horizontally centered: 640 pt at left 160 pt', () => {
    expect(toPt(toEmu(spec.w))).toBeCloseTo(640, 6)
    expect(toPt(toEmu(spec.x))).toBeCloseTo(160, 6)
    expect(spec.x + spec.w / 2).toBeCloseTo(SLIDE.widthPx / 2, 9)
  })

  it('tops out at 10.5% of the slide height (56.67 pt), not vertically centered', () => {
    expect(toPt(toEmu(spec.y))).toBeCloseTo(56.7, 1)
    expect(spec.y).toBeCloseTo(75.6, 9)
    expect(spec.y + spec.h / 2).not.toBeCloseTo(SLIDE.heightPx / 2, 0)
  })

  it('rows are exactly 29.2 pt = 370840 EMU, and the frame is rows x that', () => {
    expect(TABLE_ROW_HEIGHT_EMU).toBe(370840)
    expect(toPt(TABLE_ROW_HEIGHT_EMU)).toBeCloseTo(29.2, 9)
    expect(spec.rowHeightEmu).toBe(TABLE_ROW_HEIGHT_EMU)
    expect(TABLE_ROW_HEIGHT_PX).toBeCloseTo(38.93, 2)
    expect(spec.h).toBeCloseTo(2 * TABLE_ROW_HEIGHT_PX, 9)
    expect(toPt(2 * spec.rowHeightEmu)).toBeCloseTo(58.4, 9)
  })

  it('more rows push the table past the bottom edge instead of shrinking the rows', () => {
    const tall = tableInsertSpec(SLIDE, 20)
    expect(tall.rowHeightEmu).toBe(TABLE_ROW_HEIGHT_EMU)
    expect(tall.y).toBe(spec.y)
    expect(tall.y + tall.h).toBeGreaterThan(SLIDE.heightPx)
  })
})
