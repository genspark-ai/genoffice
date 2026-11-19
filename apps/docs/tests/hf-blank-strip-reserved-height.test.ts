import { afterEach, describe, expect, it, vi } from 'vitest'
import type { HfParagraph } from '@genoffice/docx-engine'
import { hfHasVisibleContent, hfReservedHeightPx } from '../src/renderer/editor/hf-dom'
import { estimateHfHeight } from '../src/renderer/line-metrics'

// A4 text column with 720-twip side margins
const contentW = ((11906 - 720 - 720) / 1440) * 96

/** jsdom lays nothing out: hand the strip the height a real layout gives it */
function stubStripHeight(px: number) {
  return vi.spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (
    this: Element,
  ) {
    const h = this.classList.contains('page-hf') ? px : 0
    return { x: 0, y: 0, top: 0, left: 0, right: 0, bottom: h, width: 0, height: h } as DOMRect
  })
}

afterEach(() => vi.restoreAllMocks())

describe('blank header strip reserved height', () => {
  // titlePg first-page header = one self-closing <w:p/> (Normal silent): the parser
  // hands over only the docDefaults w:after 160; the rendered line is Calibri 11 x
  // line 259 = 19.3px, the estimate's 10.5pt strut only 17.1px
  const blank: HfParagraph[] = [{ runs: [], spaceAfter: 160 }]

  it('measures the rendered strip instead of the 10.5pt estimate', () => {
    const value = { text: '', paras: blank }
    expect(hfHasVisibleContent(value)).toBe(false)
    const est = estimateHfHeight(value, contentW)
    const dom = 19.3 + 160 / 15
    expect(est).toBeLessThan(dom - 2)
    stubStripHeight(dom)
    expect(hfReservedHeightPx('header', value, contentW)).toBeCloseTo(dom, 5)
  })

  it('an unmeasured blank strip still falls back to the estimate', () => {
    const value = { text: '', paras: blank }
    const w = contentW + 3
    stubStripHeight(0)
    expect(hfReservedHeightPx('header', value, w)).toBe(estimateHfHeight(value, w))
  })

  it('a part without paragraphs never probes', () => {
    const spy = stubStripHeight(40)
    expect(hfReservedHeightPx('header', { text: '' }, contentW + 5)).toBe(0)
    expect(spy).not.toHaveBeenCalled()
  })
})
