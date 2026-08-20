import { describe, expect, it } from 'vitest'
import type { HeaderFooter, HfImage, SectionSettings } from '@genoffice/docx-engine'
import {
  hfFloatPagePos,
  hfHasVisibleContent,
  makeGapHfEl,
  makeHfFloatImgEl,
} from '../src/renderer/editor/hf-dom'
import { estimateHfHeight, hfHeaderGeom } from '../src/renderer/line-metrics'
import { effectiveTopPx } from '../src/renderer/pagination'

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
    headerDist: 48,
    sectMarginTop: 80,
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

  it('margin-relative wrapped image reserves and places from the same raw sectPr margin', () => {
    const set = { marginTop: 1200, headerDist: 720 } as SectionSettings // 80px / 48px
    const img: HfImage = {
      dataUrl: 'data:,',
      floating: true,
      wrap: 'square',
      posXPx: 10,
      posYPx: 20,
      posHRel: 'margin',
      posVRel: 'margin',
      heightPx: 100,
    }
    const headerPx = estimateHfHeight({ text: '', paras: [] }, 600, [img], hfHeaderGeom(set))
    const effTop = effectiveTopPx(set, headerPx)
    expect(effTop).toBeCloseTo(80 + 20 + 100, 5)
    const pos = hfFloatPagePos(img, { ...box, marginTop: effTop, sectMarginTop: 80 })
    expect(pos.y).toBeCloseTo(100, 5) // raw margin + offset, not the pushed-down margin
    expect(pos.y + img.heightPx!).toBeCloseTo(effTop, 5) // bottom edge meets the body top
  })

  it('paragraph-relative posOffsets measure from the header strip top, not the pushed margin', () => {
    const img: HfImage = {
      dataUrl: 'data:,',
      posXPx: 10,
      posYPx: -14,
      posHRel: 'margin',
      posVRel: 'paragraph',
    }
    expect(hfFloatPagePos(img, box)).toEqual({ x: 106, y: 34, translateX: 0, translateY: 0 })
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

describe('wrapped anchored header images push the body below their bottom edge', () => {
  const twipsToPx = (t: number) => (t / 1440) * 96

  it('page-relative offsets (prod_090 shape): body top clears the lower image bottom', () => {
    // header dist 720 twips, top margin 907 twips; two wrapTopAndBottom logos + one wrapNone
    const set = { marginTop: 907, headerDist: 720 } as SectionSettings
    const geom = hfHeaderGeom(set)
    const images: HfImage[] = [
      {
        dataUrl: 'x',
        floating: true,
        wrap: 'topBottom',
        posYPx: 35,
        posVRel: 'page',
        heightPx: 40,
      },
      {
        dataUrl: 'x',
        floating: true,
        wrap: 'topBottom',
        posYPx: 80,
        posVRel: 'page',
        heightPx: 34,
      },
      { dataUrl: 'x', floating: true, wrap: 'none', posYPx: 0, posVRel: 'page', heightPx: 300 },
    ]
    const headerPx = estimateHfHeight({ text: '', paras: [] }, 600, images, geom)
    expect(effectiveTopPx(set, headerPx)).toBeCloseTo(80 + 34, 5) // 114px ≈ Word's ~85pt
    expect(effectiveTopPx(set, 0)).toBeCloseTo(twipsToPx(907), 5) // without the push it was the raw margin
  })

  it('paragraph-relative offset (prod_004 shape): bottom measures from the header strip top', () => {
    // header dist 708 twips, top margin 1417 twips; wrapSquare logo, posOffset -14px, 101px tall
    const set = { marginTop: 1417, headerDist: 708 } as SectionSettings
    const geom = hfHeaderGeom(set)
    const images: HfImage[] = [
      {
        dataUrl: 'x',
        floating: true,
        wrap: 'square',
        posYPx: -14,
        posVRel: 'paragraph',
        heightPx: 101,
      },
    ]
    const headerPx = estimateHfHeight({ text: '', paras: [] }, 600, images, geom)
    expect(effectiveTopPx(set, headerPx)).toBeCloseTo(twipsToPx(708) - 14 + 101, 5) // ≈134px, was 94.5px
  })

  it('watermarks keep reserving nothing even with geometry supplied', () => {
    const set = { marginTop: 907, headerDist: 720 } as SectionSettings
    const geom = hfHeaderGeom(set)
    const images: HfImage[] = [
      { dataUrl: 'x', floating: true, posYPx: 0, posVRel: 'page', heightPx: 954 }, // no wrap
      {
        dataUrl: 'x',
        floating: true,
        wrap: 'square',
        behind: true,
        posYPx: 0,
        posVRel: 'page',
        heightPx: 954,
      },
    ]
    expect(estimateHfHeight(null, 600, images, geom)).toBe(0)
  })
})
