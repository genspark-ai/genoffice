import { describe, expect, it } from 'vitest'

import {
  clampZoomPercent,
  clampZoomRatio,
  SHEET_ZOOM_MAX,
  SHEET_ZOOM_MIN,
  stepZoomPercent,
} from '../src/renderer/zoom-range'

describe('zoom range', () => {
  it("clamps to Excel's 10-400 and rounds to whole percents", () => {
    expect(clampZoomPercent(5)).toBe(SHEET_ZOOM_MIN)
    expect(clampZoomPercent(10)).toBe(10)
    expect(clampZoomPercent(149.6)).toBe(150)
    expect(clampZoomPercent(400)).toBe(400)
    expect(clampZoomPercent(1000)).toBe(SHEET_ZOOM_MAX)
    expect(clampZoomPercent(Number.NaN)).toBe(100)
  })

  it('clamps ratios the same way', () => {
    expect(clampZoomRatio(0.05)).toBe(0.1)
    expect(clampZoomRatio(4.5)).toBe(4)
    expect(clampZoomRatio(0.333)).toBe(0.33)
  })

  it('steps by 10 and stops at the bounds', () => {
    expect(stepZoomPercent(100, 1)).toBe(110)
    expect(stepZoomPercent(15, -1)).toBe(10)
    expect(stepZoomPercent(395, 1)).toBe(400)
  })
})
