// Vertical ruler: which page it rides on, and the top/bottom margin geometry.
import { describe, expect, it } from 'vitest'

import {
  clampVerticalMarginDrag,
  pageIndexAt,
  rulerPageIndex,
  pageTopsOnScreen,
  verticalRulerDims,
} from '../src/renderer/components/VerticalRuler'
import { MIN_BODY_TWIPS } from '../src/renderer/components/Ruler'
import type { SectionSettings } from '@genoffice/docx-engine'

const a4 = { pageHeight: 16838, marginTop: 1418, marginBottom: 1134 } as SectionSettings

describe('pageTopsOnScreen', () => {
  it('starts each later page a gray band below the closing page bottom margin', () => {
    // measured on an A4 canvas: gap tops 935 and 2085, bottom margin + unused body 329.6px
    const tops = pageTopsOnScreen(
      143,
      [
        { top: 935, marginBottomPx: 329.6 },
        { top: 2085, marginBottomPx: 329.6 },
      ],
      1,
    )
    expect(tops).toEqual([143, 935 + 329.6 + 28, 2085 + 329.6 + 28])
  })

  it('scales the gap metrics by the canvas zoom', () => {
    const [, second] = pageTopsOnScreen(0, [{ top: 1000, marginBottomPx: 100 }], 1.5)
    expect(second).toBe(1000 + 128 * 1.5)
  })
})

describe('pageIndexAt', () => {
  const tops = [100, 1300, 2500]
  it('picks the page the caret line is on', () => {
    expect(pageIndexAt(tops, 120)).toBe(0)
    expect(pageIndexAt(tops, 1400)).toBe(1)
    expect(pageIndexAt(tops, 9000)).toBe(2)
  })
  it('keeps the first page for a caret above the canvas', () => {
    expect(pageIndexAt(tops, -50)).toBe(0)
  })
})

describe('rulerPageIndex', () => {
  const tops = [100, 1300, 2500]
  it("stays on the caret's page while it shows", () => {
    expect(rulerPageIndex(tops, 200, 0, 900)).toBe(0)
    expect(rulerPageIndex(tops, 200, 1200, 2100)).toBe(0) // its bottom still shows
  })
  it('takes the page in view once the caret page has scrolled away', () => {
    expect(rulerPageIndex(tops, 200, 1400, 2300)).toBe(1)
    expect(rulerPageIndex(tops, 200, 2400, 3300)).toBe(2)
  })
})

describe('verticalRulerDims', () => {
  it('reads the section page height and margins', () => {
    expect(verticalRulerDims(a4)).toEqual({
      pageHeight: 16838,
      marginTop: 1418,
      marginBottom: 1134,
      topSign: 1,
      bottomSign: 1,
    })
  })
  it('keeps the sign of an exact (negative) margin and survives junk', () => {
    const d = verticalRulerDims({ ...a4, marginTop: -1440, pageHeight: Number.NaN })
    expect(d.pageHeight).toBe(15840)
    expect(d.marginTop).toBe(1440)
    expect(d.topSign).toBe(-1)
  })
})

describe('clampVerticalMarginDrag', () => {
  const dims = verticalRulerDims(a4)
  it('moves the top margin but leaves half an inch of body', () => {
    expect(clampVerticalMarginDrag('top', 2000, dims)).toBe(2000)
    expect(clampVerticalMarginDrag('top', 99999, dims)).toBe(16838 - 1134 - MIN_BODY_TWIPS)
    expect(clampVerticalMarginDrag('top', -10, dims)).toBe(0)
  })
  it('measures the bottom margin up from the page bottom', () => {
    expect(clampVerticalMarginDrag('bottom', 16838 - 2000, dims)).toBe(2000)
    expect(clampVerticalMarginDrag('bottom', 0, dims)).toBe(16838 - 1418 - MIN_BODY_TWIPS)
  })
})
