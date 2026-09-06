import { afterEach, describe, expect, it, vi } from 'vitest'
import type { HfParagraph } from '@genoffice/docx-engine'
import { hfReservedHeightPx, makeGapHfEl, hfStackedSpacingPx } from '../src/renderer/editor/hf-dom'
import { estimateHfHeight } from '../src/renderer/line-metrics'

// A4 text column of SAS-2 sample 017 (11909 - 1701 - 1418 twips)
const contentW = ((11909 - 1701 - 1418) / 1440) * 96

const calibri = (text: string, extra: Partial<HfParagraph['runs'][number]> = {}) => ({
  text,
  font: 'Calibri',
  fontAscii: 'Calibri',
  sizeHalfPoints: 22,
  ...extra,
})

// 11pt italic notice / empty 10pt paragraph / 11pt version line with a right tab
const footer: HfParagraph[] = [
  {
    lineRule: 'auto',
    lineRawTwips: 240,
    lineSpacing: 1,
    tabStops: [{ pos: 8451, val: 'right' }],
    runs: [
      calibri(
        'Please do not distribute without permission.  Revised Nov 2023 (remove this in the final document)',
        { italic: true },
      ),
    ],
  },
  { lineRule: 'auto', lineRawTwips: 240, lineSpacing: 1, runs: [] },
  {
    lineRule: 'auto',
    lineRawTwips: 240,
    lineSpacing: 1,
    tabStops: [{ pos: 8306, val: 'right' }],
    runs: [
      calibri('Version 1.0 (Your Organisation’s Name) '),
      { text: '\t', sizeHalfPoints: 18 },
      { text: 'Page 7 of 22', sizeHalfPoints: 20 },
    ],
  },
]

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

describe('footer reserved height follows the rendered strip', () => {
  it('a three-line footer the estimate over-wraps reserves the DOM height, not the estimate', () => {
    const value = { text: '', paras: footer }
    const est = estimateHfHeight(value, contentW)
    // the width model breaks the 99-char notice (Word keeps it on one line)
    expect(est).toBeGreaterThan(60)
    const dom = 52.5
    stubStripHeight(dom)
    expect(hfReservedHeightPx('footer', value, contentW)).toBeCloseTo(dom, 5)
  })

  it('an unmeasured strip (no layout) still falls back to the estimate', () => {
    const value = { text: '', paras: footer }
    // the probe cache keys on content and width: a fresh width forces a re-measure
    const w = contentW + 7
    stubStripHeight(0)
    expect(hfReservedHeightPx('footer', value, w)).toBe(estimateHfHeight(value, w))
  })

  it('a wrapped floating header image still floors the DOM measure', () => {
    const value = { text: '', paras: [{ runs: [calibri('Title')] }] }
    const geom = { marginTopPx: 96, headerDistPx: 48 }
    const images = [
      {
        dataUrl: 'data:,x',
        floating: true,
        wrap: 'square' as const,
        posVRel: 'page' as const,
        posYPx: 0,
        heightPx: 200,
      },
    ]
    stubStripHeight(20)
    // image bottom at 200px on the page, 152px below the header strip top
    expect(hfReservedHeightPx('header', value, contentW, images, geom)).toBe(152)
  })
})

describe('strip paragraphs stack Word paragraph spacing', () => {
  const strip = (paras: HfParagraph[]) =>
    makeGapHfEl({ kind: 'footer', value: { text: '', paras }, pageNo: 1, pageTotal: 1 })

  it('carries the previous after into the next top margin and keeps the last after', () => {
    const el = strip([
      { runs: [{ text: 'a' }], spaceBefore: 120, spaceAfter: 200 },
      { runs: [{ text: 'b' }], spaceBefore: 60 },
      { runs: [], spaceAfter: 160 },
    ])
    const paras = el.querySelectorAll<HTMLElement>('.page-hf-para')
    expect(parseFloat(paras[0].style.marginTop)).toBeCloseTo(120 / 15, 1)
    expect(paras[0].style.marginBottom).toBe('')
    expect(parseFloat(paras[1].style.marginTop)).toBeCloseTo((200 + 60) / 15, 1)
    expect(paras[2].style.marginTop).toBe('')
    expect(parseFloat(paras[2].style.marginBottom)).toBeCloseTo(160 / 15, 1)
  })

  it('a layout-table row consumes the carried after ahead of itself', () => {
    const el = strip([
      { runs: [{ text: 'a' }], spaceAfter: 150 },
      { runs: [], cells: [{ paras: [] }], row: {} } as unknown as HfParagraph,
      { runs: [{ text: 'b' }] },
    ])
    const paras = el.querySelectorAll<HTMLElement>('.page-hf-para')
    expect(parseFloat(paras[1].style.marginTop)).toBeCloseTo(150 / 15, 1)
    expect(paras[2].style.marginTop).toBe('')
    // the preview/export strip shares the same margins
    expect(
      hfStackedSpacingPx([
        { runs: [], spaceAfter: 300 },
        { runs: [], spaceBefore: 30 },
      ]),
    ).toEqual([{}, { top: 22 }])
  })

  it('skips anchored-box paragraphs (drawn at the anchor, off the strip flow)', () => {
    const el = strip([
      { runs: [{ text: 'a' }], spaceAfter: 150 },
      { runs: [{ text: 'box' }], boxAnchored: true, spaceBefore: 900 },
      { runs: [{ text: 'b' }] },
    ])
    const paras = el.querySelectorAll<HTMLElement>('.page-hf-para')
    expect(paras[1].style.marginTop).toBe('')
    expect(parseFloat(paras[2].style.marginTop)).toBeCloseTo(150 / 15, 1)
  })
})
