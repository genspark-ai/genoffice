/** Floating vs inline image classification unit tests: no wasm. */
import { describe, expect, it } from 'vitest'
import { classifyFloatImages, suppressTextShadowImages } from '../src/analyze/floats'
import { clusterCombiningMarks, groupIntoLines } from '../src/analyze/lines'
import { splitIntoUnits } from '../src/analyze/units'
import type { ImageBlock, PdfChar } from '../src/ir'
import { mkText } from './helpers/chars'

const unitsOf = (chars: PdfChar[]) => splitIntoUnits(groupIntoLines(clusterCombiningMarks(chars)))

function image(x0: number, y0: number, x1: number, y1: number): ImageBlock {
  return {
    kind: 'image',
    box: { x0, y0, x1, y1 },
    data: new Uint8Array(0),
    mime: 'image/png',
    pixelWidth: 10,
    pixelHeight: 10,
  }
}

describe('classifyFloatImages: page-edge decor band (P16 F)', () => {
  const pageBox = { x0: 0, y0: 0, x1: 595, y1: 842 }
  const bodyChars = [
    ...mkText('body paragraph line one running here', 72, { y: 700 }).chars,
    ...mkText('body paragraph line two running here', 72, { y: 686 }).chars,
  ]

  it('a footer bar hugging the page bottom floats behind at its position', () => {
    // full-width green bar + sprout PNG inside the bottom 12% (y-up 3..72)
    const bar = image(11, 3, 595, 72)
    const { floats, inline } = classifyFloatImages([bar], unitsOf(bodyChars), 595 * 842, pageBox)
    expect(inline).toHaveLength(0)
    expect(floats).toHaveLength(1)
    expect(bar.float?.wrap).toBe('behind')
  })

  it('a mid-page trailing image stays inline', () => {
    const fig = image(150, 300, 450, 500)
    const { floats, inline } = classifyFloatImages([fig], unitsOf(bodyChars), 595 * 842, pageBox)
    expect(floats).toHaveLength(0)
    expect(inline).toHaveLength(1)
  })

  it('an edge image under body text keeps flowing when text overlaps it', () => {
    const chars = [
      ...mkText('caption printed over the band area', 72, { y: 30 }).chars,
      ...mkText('caption second line over the band', 72, { y: 16 }).chars,
    ]
    const bar = image(11, 3, 595, 72)
    const { inline, floats } = classifyFloatImages([bar], unitsOf(chars), 595 * 842, pageBox)
    // text ON the image → the existing over-rule floats it behind anyway;
    // this documents that the band rule does not steal precedence
    expect(floats.length + inline.length).toBe(1)
  })
})

describe('classifyFloatImages', () => {
  it('text running beside the image floats it on the empty side', () => {
    const chars = [
      ...mkText('text wrapping around the picture', 72, { y: 700 }).chars,
      ...mkText('continues along its left side too', 72, { y: 686 }).chars,
      ...mkText('and one more wrapped line here', 72, { y: 672 }).chars,
    ]
    const img = image(400, 660, 520, 710) // right of the text
    const { floats, inline } = classifyFloatImages([img], unitsOf(chars))
    expect(inline).toHaveLength(0)
    expect(floats).toHaveLength(1)
    expect(img.float?.wrap).toBe('square-right')
    expect(img.float?.xOffsetPt).toBeCloseTo(400 - 72, 5)
  })

  it('text on top of the image floats it behind', () => {
    const img = image(60, 650, 460, 720)
    const chars = [
      ...mkText('caption text over the backdrop', 100, { y: 700 }).chars,
      ...mkText('second line over the image too', 100, { y: 686 }).chars,
    ]
    classifyFloatImages([img], unitsOf(chars))
    expect(img.float?.wrap).toBe('behind')
  })

  it('an image in its own vertical band stays inline (P1 behavior)', () => {
    const chars = [
      ...mkText('paragraph above the figure', 72, { y: 700 }).chars,
      ...mkText('paragraph below the figure', 72, { y: 500 }).chars,
    ]
    const img = image(236, 560, 376, 660)
    const { floats, inline } = classifyFloatImages([img], unitsOf(chars))
    expect(floats).toHaveLength(0)
    expect(inline).toHaveLength(1)
    expect(img.float).toBeUndefined()
  })

  it('a single overlapping line (heading over rule art) does not float', () => {
    const chars = mkText('one line only beside it', 72, { y: 700 }).chars
    const img = image(400, 690, 500, 712)
    const { floats } = classifyFloatImages([img], unitsOf(chars))
    expect(floats).toHaveLength(0)
  })

  it('tiny marker icons inside the text span float behind (P5)', () => {
    const chars = [
      ...mkText('reference one with a favicon', 84, { y: 700 }).chars,
      ...mkText('reference two with a favicon', 84, { y: 500 }).chars,
    ]
    // 8×8 margin icon between the two lines (timeline-dot position)
    const icon = image(64, 600, 72, 608)
    const { floats, inline } = classifyFloatImages([icon], unitsOf(chars))
    expect(inline).toHaveLength(0)
    expect(floats).toHaveLength(1)
    expect(icon.float?.wrap).toBe('behind')
  })

  it('a tiny image outside the text span stays inline', () => {
    const chars = mkText('all page text down here', 72, { y: 500 }).chars
    const icon = image(64, 700, 72, 708) // above every text unit
    const { floats, inline } = classifyFloatImages([icon], unitsOf(chars))
    expect(floats).toHaveLength(0)
    expect(inline).toHaveLength(1)
  })

  it('a page-covering wallpaper photo floats behind even with one text unit (P11 A)', () => {
    const page = { x0: 0, y0: 0, x1: 960, y1: 540 }
    const chars = mkText('lone slide title', 72, { y: 400 }).chars
    const img = image(0, 0, 960, 540)
    const { floats, inline } = classifyFloatImages([img], unitsOf(chars), 960 * 540, page)
    expect(inline).toHaveLength(0)
    expect(floats).toHaveLength(1)
    expect(img.float?.wrap).toBe('behind')
  })

  it('an over-scan wallpaper (box past the page edges) still floats behind (P11 A)', () => {
    const page = { x0: 0, y0: 0, x1: 960, y1: 540 }
    const chars = mkText('lone slide title', 72, { y: 400 }).chars
    const img = image(-120, -40, 1020, 560) // clipped to the page when drawn
    const { floats } = classifyFloatImages([img], unitsOf(chars), 960 * 540, page)
    expect(floats).toHaveLength(1)
    expect(img.float?.wrap).toBe('behind')
  })

  it('a half-page photo does not trip the wallpaper rule', () => {
    const page = { x0: 0, y0: 0, x1: 960, y1: 540 }
    const chars = [
      ...mkText('paragraph above the figure', 72, { y: 500 }).chars,
      ...mkText('paragraph below the figure', 72, { y: 100 }).chars,
    ]
    const img = image(0, 140, 960, 420)
    const { floats, inline } = classifyFloatImages([img], unitsOf(chars), 960 * 540, page)
    expect(floats).toHaveLength(0)
    expect(inline).toHaveLength(1)
  })
})

describe('suppressTextShadowImages (P11 B)', () => {
  it('drops a glyph-raster shadow image hugging its display line', () => {
    // 4 CJK chars at 36pt starting at x=57 → line box ≈ (57,453)-(201,489)
    const chars = mkText('赴日就業', 57, { y: 461, fontSize: 36 }).chars
    // blur-padded raster of the same text, centered on the line
    const shadow = image(34, 425, 225, 504)
    const kept = suppressTextShadowImages([shadow], unitsOf(chars))
    expect(kept).toHaveLength(0)
  })

  it('keeps a real figure that happens to contain one caption line', () => {
    const chars = mkText('small caption', 120, { y: 300, fontSize: 9 }).chars
    const figure = image(80, 250, 480, 500) // area far beyond 3.5× the line
    const kept = suppressTextShadowImages([figure], unitsOf(chars))
    expect(kept).toHaveLength(1)
  })

  it('keeps a banner photo much wider than its title line', () => {
    const chars = mkText('標題', 400, { y: 500, fontSize: 28 }).chars
    const banner = image(0, 470, 960, 540) // line spans only ~6% of the width
    const kept = suppressTextShadowImages([banner], unitsOf(chars))
    expect(kept).toHaveLength(1)
  })

  it('keeps body-text-size overlaps (inline formula images etc.)', () => {
    const chars = mkText('formula text', 100, { y: 400, fontSize: 10 }).chars
    const img = image(98, 392, 165, 412)
    const kept = suppressTextShadowImages([img], unitsOf(chars))
    expect(kept).toHaveLength(1)
  })
})
