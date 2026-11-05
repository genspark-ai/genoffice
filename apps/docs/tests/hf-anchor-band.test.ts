import { describe, expect, it } from 'vitest'
import type { HfImage, SectionSettings } from '@genoffice/docx-engine'
import { estimateHfHeight, hfHeaderGeom } from '../src/renderer/line-metrics'
import { hfBodyObstacles, hfSideWrapObstacles, hfStripGeom } from '../src/renderer/editor/hf-dom'
import { effectiveTopPx } from '../src/renderer/pagination-sections'

const pt = (n: number) => (n * 96) / 72

// A4, top margin = header distance = 966 twips (48.3pt); one Times New Roman
// 12 header paragraph; 75 x 43.9pt logo anchored wrapTopAndBottom at
// paragraph offset -24pt (Word probe 2026-11-03, compatibilityMode 15 vs legacy)
const set = {
  pageWidth: 11906,
  pageHeight: 16838,
  marginLeft: 851,
  marginRight: 851,
  marginTop: 966,
  marginBottom: 851,
  headerDist: 966,
  footerDist: 709,
} as SectionSettings
const logo = (posYPt: number, heightPt = 43.9): HfImage => ({
  dataUrl: 'x',
  floating: true,
  wrap: 'topBottom',
  widthPx: pt(75),
  heightPx: pt(heightPt),
  posH: 'center',
  posHRel: 'margin',
  posVRel: 'paragraph',
  posYPx: pt(posYPt),
})
const part = {
  text: '',
  paras: [{ runs: [], emptyRunSizeHalfPoints: 24, emptyRunFontFamily: 'Times New Roman' }],
}
const width = pt(510.2)

describe('header wrapTopAndBottom band vs the anchor paragraph line (probe 2026-11-03)', () => {
  it('compatibilityMode 15: a band cutting into the line moves the line below it', () => {
    const h = estimateHfHeight(part, width, [logo(-24)], hfHeaderGeom(set))
    const line = estimateHfHeight(part, width, [], hfHeaderGeom(set))
    // band bottom 19.9pt under the strip top + the pushed line
    expect(h).toBeCloseTo(pt(43.9 - 24) + line, 1)
    expect(effectiveTopPx(set, h) / pt(1)).toBeCloseTo(48.3 + 19.9 + line / pt(1), 1)
  })

  it('a band wholly above the paragraph does not push', () => {
    const h = estimateHfHeight(part, width, [logo(-60)], hfHeaderGeom(set))
    expect(h).toBeCloseTo(estimateHfHeight(part, width, [], hfHeaderGeom(set)), 5)
  })

  it('a band starting below the line leaves the line and pushes the later paragraphs', () => {
    const two = { text: '', paras: [part.paras[0], part.paras[0]] }
    const line = estimateHfHeight(part, width, [], hfHeaderGeom(set))
    const h = estimateHfHeight(two, width, [logo(20)], hfHeaderGeom(set))
    expect(h).toBeCloseTo(pt(20 + 43.9) + line, 1)
  })

  it('later paragraphs that fit in the gap above the band stay there', () => {
    const three = { text: '', paras: [part.paras[0], part.paras[0], part.paras[0]] }
    const line = estimateHfHeight(part, width, [], hfHeaderGeom(set))
    // the band starts two lines down: paragraph 2 fits above it, only paragraph 3 is pushed
    const h = estimateHfHeight(three, width, [logo((2 * line) / pt(1) + 1)], hfHeaderGeom(set))
    expect(h).toBeCloseTo(pt((2 * line) / pt(1) + 1 + 43.9) + line, 1)
  })

  it('compatibilityMode 15 does not clamp a negative offset at the page top', () => {
    const h = estimateHfHeight(part, width, [logo(-70, 100)], hfHeaderGeom(set))
    const line = estimateHfHeight(part, width, [], hfHeaderGeom(set))
    expect(h).toBeCloseTo(pt(30) + line, 1)
  })

  it('legacy layout: the box clamps to the page top and the line stays at the strip top', () => {
    const geom = hfHeaderGeom(set, true)
    expect(estimateHfHeight(part, width, [logo(-24)], geom)).toBeCloseTo(pt(43.9 - 24), 1)
    expect(estimateHfHeight(part, width, [logo(-70, 100)], geom)).toBeCloseTo(pt(100 - 48.3), 1)
  })

  it('the strip layout gets the band as a column-wide float only under compatibilityMode 15', () => {
    const modern = hfSideWrapObstacles('header', [logo(-24)], hfStripGeom(set))
    expect(modern).toHaveLength(1)
    expect(modern[0].top).toBe(0)
    expect(modern[0].height).toBeCloseTo(pt(43.9 - 24), 1)
    expect(modern[0].width).toBeCloseTo(width, 1)
    expect(hfSideWrapObstacles('header', [logo(-24)], hfStripGeom(set, true))).toHaveLength(0)
  })
})

describe('side-wrapped header drawings reaching into the body (probe 2026-11-03)', () => {
  // landscape Letter, header distance 36pt, top margin 72pt: a 504 x 85.2pt
  // behindDoc wrapTight banner centred on the margins at paragraph offset
  // -27.75pt occupies page y 8.3..93.5 and the first body lines flow beside it
  const land = {
    pageWidth: 15840,
    pageHeight: 12240,
    marginLeft: 1440,
    marginRight: 1440,
    marginTop: 1440,
    marginBottom: 1440,
    headerDist: 720,
    footerDist: 720,
  } as SectionSettings
  const banner: HfImage = {
    dataUrl: 'x',
    floating: true,
    behind: true,
    wrap: 'tight',
    widthPx: pt(504),
    heightPx: pt(85.2),
    posH: 'center',
    posHRel: 'margin',
    posVRel: 'paragraph',
    posYPx: pt(-27.75),
    wrapDistPx: { l: pt(9), r: pt(9), t: 0, b: 0 },
  }

  it('is an obstacle of the first body lines, in column px from the body top', () => {
    const [o] = hfBodyObstacles([banner], hfStripGeom(land), pt(72))
    expect(o.left / pt(1)).toBeCloseTo(72 - 9, 1)
    expect(o.right / pt(1)).toBeCloseTo(72 + 504 + 9, 1)
    expect(o.bottom / pt(1)).toBeCloseTo(93.5 - 72, 1)
  })

  it('a drawing ending above the body top and a wrapNone one register nothing', () => {
    expect(hfBodyObstacles([{ ...banner, heightPx: pt(50) }], hfStripGeom(land), pt(72))).toEqual(
      [],
    )
    expect(hfBodyObstacles([{ ...banner, wrap: 'none' }], hfStripGeom(land), pt(72))).toEqual([])
  })

  it('does not move the body top itself', () => {
    const h = estimateHfHeight({ text: '', paras: [] }, pt(648), [banner], hfHeaderGeom(land))
    expect(effectiveTopPx(land, h)).toBeCloseTo(pt(72), 3)
  })
})
