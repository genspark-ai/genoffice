import { describe, expect, it } from 'vitest'

/**
 * Regression for setTableStyle borderWidthPt validation
 * (packages/pptx-ops/src/ops/table-ops.ts resolveTableStyle):
 * old check `!(Number(w) > 0)` passes Infinity, then
 * Math.round(Infinity * EMU_PER_PT) writes Infinity EMU.
 * New check requires finite 0 < w <= 1584 (PowerPoint's line weight limit).
 */
function isValidBorderWidthPt(v: unknown): boolean {
  const w = Number(v)
  return Number.isFinite(w) && w > 0 && w <= 1584
}

describe('setTableStyle borderWidthPt validation', () => {
  it('rejects non-finite and non-positive widths', () => {
    expect(isValidBorderWidthPt(NaN)).toBe(false)
    expect(isValidBorderWidthPt(Infinity)).toBe(false)
    expect(isValidBorderWidthPt(-1)).toBe(false)
    expect(isValidBorderWidthPt(0)).toBe(false)
  })

  it('rejects unreasonably large widths', () => {
    expect(isValidBorderWidthPt(1585)).toBe(false)
    expect(isValidBorderWidthPt(1e308)).toBe(false)
  })

  it('accepts normal widths', () => {
    expect(isValidBorderWidthPt(0.5)).toBe(true)
    expect(isValidBorderWidthPt(1)).toBe(true)
    expect(isValidBorderWidthPt(12)).toBe(true)
    expect(isValidBorderWidthPt(1584)).toBe(true)
  })
})
