/**
 * Skia (Chromium printToPDF) wraps svg pictures in a form XObject whose matrix
 * flips y and pre-flips the image matrix to compensate. PDFium renders the
 * image in its own frame, so the pixels come back upside down unless the
 * enclosing mirror is undone.
 */
import { describe, expect, it } from 'vitest'
import { extractPage, mirrorBetweenFrames, mirrorRgba, withPdfDocument } from '../src/extract'
import { buildFormMirroredImagePdf } from './helpers/fixtures'
import { decodePngRgba } from './helpers/png-decode'
import { loadPdfium } from './helpers/wasm'

describe('images inside a mirroring form XObject', () => {
  it('come out upright: the page-space top of the picture is its top row', async () => {
    const m = await loadPdfium()
    const pdf = await buildFormMirroredImagePdf()
    const page = withPdfDocument(m, pdf, (doc) => extractPage(m, doc, 0))

    expect(page.images).toHaveLength(1)
    const img = page.images[0]!
    // form matrix 0.25 × 400 = 100pt square at (100, 500)..(200, 600)
    expect(img.box.x0).toBeCloseTo(100, 0)
    expect(img.box.x1).toBeCloseTo(200, 0)
    expect(img.box.y0).toBeCloseTo(500, 0)
    expect(img.box.y1).toBeCloseTo(600, 0)
    expect(img.mime).toBe('image/png')
    const { rgba, width, height } = decodePngRgba(img.data)
    const px = (x: number, y: number) => rgba.subarray((y * width + x) * 4, (y * width + x) * 4 + 3)
    // red half on top, blue half at the bottom (page space, row 0 = top)
    expect(px(width >> 1, 1)[0]).toBeGreaterThan(200)
    expect(px(width >> 1, 1)[2]).toBeLessThan(60)
    expect(px(width >> 1, height - 2)[2]).toBeGreaterThan(200)
    expect(px(width >> 1, height - 2)[0]).toBeLessThan(60)
  })

  it('mirrorBetweenFrames: sign flips only, never rotations or zero scales', () => {
    // Skia layout: y-flipping form, pre-flipped image
    expect(mirrorBetweenFrames([400, 0, 0, -400, 0, 400], [100, 0, 0, 100, 100, 500])).toEqual({
      x: false,
      y: true,
    })
    // form mirrors x
    expect(mirrorBetweenFrames([100, 0, 0, 100, 0, 0], [-100, 0, 0, 100, 300, 0])).toEqual({
      x: true,
      y: false,
    })
    // upright in both frames
    expect(mirrorBetweenFrames([100, 0, 0, 100, 0, 0], [50, 0, 0, 50, 0, 0])).toEqual({
      x: false,
      y: false,
    })
    // a quarter-turn form zeroes the diagonals: not a mirror
    expect(mirrorBetweenFrames([100, 0, 0, 100, 0, 0], [0, 100, -100, 0, 300, 0])).toEqual({
      x: false,
      y: false,
    })
    expect(mirrorBetweenFrames([0, 100, -100, 0, 0, 0], [100, 0, 0, 100, 0, 0])).toEqual({
      x: false,
      y: false,
    })
  })

  it('mirrorRgba flips rows and columns independently', () => {
    const px = {
      width: 2,
      height: 2,
      rgba: Uint8Array.from([1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4]),
    }
    expect(
      Array.from(mirrorRgba(px, { x: false, y: true }).rgba.filter((_, i) => i % 4 === 0)),
    ).toEqual([3, 4, 1, 2])
    expect(
      Array.from(mirrorRgba(px, { x: true, y: false }).rgba.filter((_, i) => i % 4 === 0)),
    ).toEqual([2, 1, 4, 3])
    expect(
      Array.from(mirrorRgba(px, { x: true, y: true }).rgba.filter((_, i) => i % 4 === 0)),
    ).toEqual([4, 3, 2, 1])
    expect(mirrorRgba(px, { x: false, y: false })).toBe(px)
  })
})
