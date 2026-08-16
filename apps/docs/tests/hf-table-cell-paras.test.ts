import { describe, expect, it } from 'vitest'
import type { HeaderFooter, HfImage } from '@genoffice/docx-engine'
import {
  hfFloatPagePos,
  hfHasVisibleContent,
  makeGapHfEl,
  makeHfFloatImgEl,
} from '../src/renderer/editor/hf-dom'
import { estimateHfHeight } from '../src/renderer/line-metrics'

describe('header table cells keep per-paragraph lines', () => {
  const value: HeaderFooter = {
    text: 'Line one Line two',
    paras: [
      {
        runs: [],
        cells: [
          { paras: [[]], fill: 'C00000', widthPct: 10 },
          { paras: [[{ text: 'Line one', bold: true }], [{ text: 'Line two' }]], widthPct: 90 },
        ],
      },
    ],
  }

  it('renders one block line per cell paragraph', () => {
    const el = makeGapHfEl({ kind: 'header', value, pageNo: 1, pageTotal: 1 })
    const cells = el.querySelectorAll('.page-hf-cell')
    expect(cells).toHaveLength(2)
    const titleParas = cells[1].querySelectorAll('.page-hf-cell-para')
    expect(titleParas).toHaveLength(2)
    expect(titleParas[0].textContent).toBe('Line one')
    expect(titleParas[1].textContent).toBe('Line two')
    // the shaded cell keeps a line for its lone empty paragraph
    const shaded = cells[0] as HTMLElement
    expect(shaded.style.backgroundColor).toBeTruthy()
    expect(shaded.querySelectorAll('.page-hf-cell-para')).toHaveLength(1)
  })

  it('estimateHfHeight sizes a table row by its tallest cell paragraph stack', () => {
    const oneLine = estimateHfHeight(
      { text: '', paras: [{ runs: [], cells: [{ paras: [[{ text: 'only' }]] }] }] },
      600,
    )
    const twoLines = estimateHfHeight(
      {
        text: '',
        paras: [
          {
            runs: [],
            cells: [{ paras: [[]] }, { paras: [[{ text: 'one' }], [{ text: 'two' }]] }],
          },
        ],
      },
      600,
    )
    expect(oneLine).toBeGreaterThan(0)
    expect(twoLines).toBeGreaterThan(oneLine * 1.5)
  })
})

describe('floating header image positioning', () => {
  const box = {
    pageW: 816,
    pageH: 1056,
    marginLeft: 96,
    marginRight: 96,
    marginTop: 96,
    marginBottom: 96,
  }

  it('page-relative posOffsets measure from the page corner', () => {
    const img: HfImage = {
      dataUrl: 'data:,',
      posXPx: 10,
      posYPx: 20,
      posHRel: 'page',
      posVRel: 'page',
    }
    expect(hfFloatPagePos(img, box)).toEqual({ x: 10, y: 20, translateX: 0, translateY: 0 })
  })

  it('margin-relative posOffsets measure from the margin box', () => {
    const img: HfImage = {
      dataUrl: 'data:,',
      posXPx: 10,
      posYPx: 20,
      posHRel: 'margin',
      posVRel: 'margin',
    }
    expect(hfFloatPagePos(img, box)).toEqual({ x: 106, y: 116, translateX: 0, translateY: 0 })
  })

  it('alignment fields reproduce the legacy margin-box anchors', () => {
    expect(hfFloatPagePos({ dataUrl: 'data:,' }, box)).toEqual({
      x: 96,
      y: 96,
      translateX: 0,
      translateY: 0,
    })
    expect(hfFloatPagePos({ dataUrl: 'data:,', posH: 'center', posV: 'bottom' }, box)).toEqual({
      x: 408,
      y: 960,
      translateX: -50,
      translateY: -100,
    })
  })

  it('gap-hosted element positions from the next page origin (gap bottom = marginTop above it)', () => {
    const img: HfImage = {
      dataUrl: 'data:,',
      posXPx: 0,
      posYPx: 0,
      posHRel: 'page',
      posVRel: 'page',
      widthPx: 816,
      heightPx: 1056,
      behind: true,
      washout: true,
    }
    const el = makeHfFloatImgEl(img, box, 'gap')
    expect(el.className).toBe('page-hf-float-img')
    expect(el.style.left).toBe('0px')
    expect(el.style.top).toBe('calc(100% - 96px)')
    expect(el.style.width).toBe('816px')
    expect(el.style.filter).toContain('brightness')
  })

  it('lead-hosted element positions from the first page content origin', () => {
    const img: HfImage = {
      dataUrl: 'data:,',
      posXPx: 5,
      posYPx: 6,
      posHRel: 'page',
      posVRel: 'page',
    }
    const el = makeHfFloatImgEl(img, box, 'lead')
    expect(el.style.left).toBe('-91px')
    expect(el.style.top).toBe('-90px')
  })
})

describe('cell run images (header logo inside a layout-table cell)', () => {
  const value: HeaderFooter = {
    text: 'Title',
    paras: [
      {
        runs: [],
        cells: [
          {
            paras: [
              [
                {
                  text: '',
                  image: {
                    dataUrl: 'data:image/png;base64,x',
                    xml: '<w:drawing/>',
                    widthPx: 96,
                    heightPx: 48,
                  },
                },
              ],
            ],
            widthPct: 12,
          },
          { paras: [[{ text: 'Title' }]], widthPct: 88 },
        ],
      },
    ],
  }

  it('renders the image inside its cell paragraph', () => {
    const el = makeGapHfEl({ kind: 'header', value, pageNo: 1, pageTotal: 1 })
    const img = el.querySelector<HTMLImageElement>('.page-hf-cell .page-hf-cell-img')
    expect(img).not.toBeNull()
    expect(img!.style.width).toBe('96px')
    expect(img!.style.height).toBe('48px')
    // no part-level image strip involved
    expect(el.querySelector('.page-hf-images')).toBeNull()
  })

  it('estimateHfHeight grows the row to the cell image height', () => {
    const textOnly = estimateHfHeight(
      { text: '', paras: [{ runs: [], cells: [{ paras: [[{ text: 'Title' }]] }] }] },
      600,
    )
    const withLogo = estimateHfHeight(value, 600)
    expect(withLogo).toBeGreaterThanOrEqual(48)
    expect(withLogo).toBeLessThan(textOnly + 48) // image joins the row line box, no extra stacked band
  })
})

describe('empty header paragraphs', () => {
  it('height follows the paragraph mark run size, not the 10.5pt default', () => {
    const part = (runs: Array<{ text: string; sizeHalfPoints?: number }>): HeaderFooter => ({
      text: '',
      paras: [{ runs }],
    })
    const def = estimateHfHeight(part([{ text: '' }]), 600)
    expect(def).toBeCloseTo(10.5 * (96 / 72) * 1.22, 3)
    const sized = estimateHfHeight(part([{ text: ' ', sizeHalfPoints: 36 }]), 600)
    expect(sized).toBeCloseTo(18 * (96 / 72) * 1.22, 3)
  })
})

describe('empty header with only a floating watermark (sample-17 shape)', () => {
  it('estimateHfHeight reserves nothing', () => {
    const floating = [{ heightPx: 954, floating: true }]
    expect(estimateHfHeight(null, 600, floating)).toBe(0)
    expect(estimateHfHeight({ text: '', paras: [] }, 600, floating)).toBe(0)
  })

  it('hfHasVisibleContent is false for an empty part with an empty inline-image list', () => {
    expect(hfHasVisibleContent({ text: '', paras: [] }, [])).toBe(false)
    expect(hfHasVisibleContent(null, [])).toBe(false)
  })
})
