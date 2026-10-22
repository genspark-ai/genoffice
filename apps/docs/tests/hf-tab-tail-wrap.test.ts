import { afterEach, describe, expect, it, vi } from 'vitest'
import type { HeaderFooter } from '@genoffice/docx-engine'
import { hfReservedHeightPx, hfTabSegments, makeGapHfEl } from '../src/renderer/editor/hf-dom'

// Letter page, 1417-twip side margins: 9406-twip column (627 px)
const geom = {
  pageW: 816,
  pageH: 1056,
  marginLeft: 94.5,
  marginRight: 94.5,
  marginTop: 83,
  marginBottom: 83,
  headerStripTop: 48,
  footerDist: 48,
}
const stops = [
  { pos: 4680, val: 'center' as const },
  { pos: 9360, val: 'right' as const },
]
const long = 'Guia para la elaboracion del proyecto '.repeat(6)

describe('tab tail overflowing the column wraps like body text', () => {
  it('keeps a long tail after a center tab in flow instead of a positioned segment', () => {
    const para = {
      runs: [{ text: '\t' }, { text: long, sizeHalfPoints: 16 }],
      tabStops: stops,
      align: 'center' as const,
    }
    const layout = hfTabSegments(para, undefined, undefined, 627)!
    expect(layout.segments).toHaveLength(0)
    expect(layout.flow?.runs.map((r) => r.text)).toEqual([long])
    // the centered text would start before the line start: the tab collapses
    expect(layout.flow?.gapPx).toBe(0)
    // w:jc is applied per wrapped line by the browser, not as a line shift
    expect(layout.shift).toBeUndefined()
  })

  it('a tail that fits stays a positioned segment (no column = no change)', () => {
    const para = { runs: [{ text: '\tPage 1', sizeHalfPoints: 16 }], tabStops: stops }
    expect(hfTabSegments(para, undefined, undefined, 627)!.flow).toBeUndefined()
    const wide = { runs: [{ text: '\t' }, { text: long, sizeHalfPoints: 16 }], tabStops: stops }
    expect(hfTabSegments(wide)!.flow).toBeUndefined()
    expect(hfTabSegments(wide)!.segments).toHaveLength(1)
  })

  it('text right-aligned at a stop inside the margin may still reach it', () => {
    // 9360-twip right stop on a 9070-twip column: Word keeps the text at the stop
    const para = { runs: [{ text: '\t\tPage 1 of 9', sizeHalfPoints: 24 }], tabStops: stops }
    const layout = hfTabSegments(para, undefined, undefined, 604.7)!
    expect(layout.flow).toBeUndefined()
    expect(layout.segments[0].endPx).toBeCloseTo(624, 0)
  })

  it('a lead keeps its tab gap before the wrapping tail', () => {
    const para = {
      runs: [
        { text: 'Left', sizeHalfPoints: 16 },
        { text: '\t' + long, sizeHalfPoints: 16 },
      ],
      tabStops: [{ pos: 1440, val: 'left' as const }],
    }
    const layout = hfTabSegments(para, undefined, undefined, 627)!
    expect(layout.lead.map((r) => r.text)).toEqual(['Left'])
    expect(layout.flow?.gapPx).toBeGreaterThan(0)
    expect(layout.flow!.gapPx).toBeCloseTo(96 - 4 * 8 * 0.5 * (4 / 3), 1)
  })

  it('renders the tail as wrapping in-flow text with the paragraph jc', () => {
    const value: HeaderFooter = {
      text: '\t' + long,
      pageNumber: false,
      paras: [
        {
          runs: [{ text: '\t' }, { text: long, sizeHalfPoints: 16 }],
          tabStops: stops,
          align: 'center',
        },
      ],
    }
    const el = makeGapHfEl({ kind: 'footer', value, pageNo: 1, pageTotal: 9, geom })
    const para = el.querySelector<HTMLElement>('.page-hf-para.page-hf-tabbed')!
    expect(para.classList.contains('page-hf-tabbed-flow')).toBe(true)
    expect(para.style.textAlign).toBe('center')
    expect(para.querySelector('.page-hf-tabseg')).toBeNull()
    expect(para.textContent).toContain(long.trim())
  })
})

describe('the reserved-height probe lays the tail out wrapped too', () => {
  afterEach(() => vi.restoreAllMocks())

  it('measures the probe strip with the tail in flow, not as a one-line segment', () => {
    const value: HeaderFooter = {
      text: '\t' + long,
      pageNumber: false,
      paras: [
        {
          runs: [{ text: '\t' }, { text: long, sizeHalfPoints: 16 }],
          tabStops: stops,
          align: 'center',
        },
      ],
    }
    // jsdom lays nothing out: a flow tail of this length takes two 8pt lines
    // (~13.3 px each), a positioned nowrap segment one
    const lineH = 8 * (4 / 3) * 1.25
    let flowProbed = false
    vi.spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (
      this: Element,
    ) {
      if (!this.classList.contains('page-hf')) return new DOMRect()
      const flow = !!this.querySelector('.page-hf-tabbed-flow')
      flowProbed ||= flow
      const h = flow ? 2 * lineH : lineH
      return { x: 0, y: 0, top: 0, left: 0, right: 0, bottom: h, width: 0, height: h } as DOMRect
    })
    expect(hfReservedHeightPx('footer', value, 627)).toBeGreaterThanOrEqual(2 * lineH)
    expect(flowProbed).toBe(true)
  })
})
