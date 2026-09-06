import { describe, expect, it } from 'vitest'
import { tabSegmentWidth } from '../src/renderer/editor/decoration-extensions'

const line = (left: number, top: number) => ({ left, top, bottom: top + 20 })

describe('tabSegmentWidth', () => {
  it('measures a segment that stays on the tab line', () => {
    expect(tabSegmentWidth(line(100, 0), line(340, 0), 600, 1)).toBe(240)
    expect(tabSegmentWidth(line(100, 0), line(340, 0), 600, 2)).toBe(120)
    expect(tabSegmentWidth(line(100, 0), line(80, 5), 600, 1)).toBe(0)
  })

  // a wrapped segment is at least the column wide; measuring its endpoints
  // pinned a claim's tab to a space width when the last line ended flush right
  it('treats a segment ending on a later line as column-wide', () => {
    expect(tabSegmentWidth(line(100, 0), line(620, 28), 600, 1)).toBe(600)
    expect(tabSegmentWidth(line(100, 0), line(90, 56), 600, 1)).toBe(600)
  })
})
