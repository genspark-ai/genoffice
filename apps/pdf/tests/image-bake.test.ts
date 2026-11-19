import { describe, expect, it } from 'vitest'
import { cropRect, flipPixels, multiplyAlpha } from '../src/renderer/image-bake'
import type { PixelImage } from '../src/renderer/cutout'

/** 2×2 test image; each pixel r=index, g=100+index, b=200+index, a=40+index */
const img2x2 = (): PixelImage => {
  const data = new Uint8ClampedArray(16)
  for (let i = 0; i < 4; i++) {
    data[i * 4] = i
    data[i * 4 + 1] = 100 + i
    data[i * 4 + 2] = 200 + i
    data[i * 4 + 3] = 40 + i
  }
  return { data, width: 2, height: 2 }
}

const pixelAt = (data: Uint8ClampedArray, w: number, x: number, y: number) => [
  ...data.slice((y * w + x) * 4, (y * w + x) * 4 + 4),
]

describe('flipPixels', () => {
  it('mirrors horizontally: rows keep their y, columns swap', () => {
    const src = img2x2()
    const out = flipPixels(src, 'h')
    // original layout: [0 1] / [2 3] → [1 0] / [3 2]
    expect(pixelAt(out, 2, 0, 0)).toEqual(pixelAt(src.data, 2, 1, 0))
    expect(pixelAt(out, 2, 1, 0)).toEqual(pixelAt(src.data, 2, 0, 0))
    expect(pixelAt(out, 2, 0, 1)).toEqual(pixelAt(src.data, 2, 1, 1))
    expect(pixelAt(out, 2, 1, 1)).toEqual(pixelAt(src.data, 2, 0, 1))
  })

  it('mirrors vertically: columns keep their x, rows swap', () => {
    const src = img2x2()
    const out = flipPixels(src, 'v')
    expect(pixelAt(out, 2, 0, 0)).toEqual(pixelAt(src.data, 2, 0, 1))
    expect(pixelAt(out, 2, 1, 1)).toEqual(pixelAt(src.data, 2, 1, 0))
  })

  it('double flip restores the original and never mutates the input', () => {
    const src = img2x2()
    const snapshot = [...src.data]
    const twice = flipPixels({ data: flipPixels(src, 'h'), width: 2, height: 2 }, 'h')
    expect([...twice]).toEqual(snapshot)
    expect([...src.data]).toEqual(snapshot)
  })
})

describe('multiplyAlpha', () => {
  it('scales only the alpha channel, rounding to integers', () => {
    const src = img2x2()
    const out = multiplyAlpha(src, 0.5)
    expect(pixelAt(out, 2, 0, 0)).toEqual([0, 100, 200, 20])
    expect(pixelAt(out, 2, 1, 1)).toEqual([3, 103, 203, Math.round(43 * 0.5)])
  })

  it('clamps the factor to 0..1', () => {
    const src = img2x2()
    expect(multiplyAlpha(src, 2)[3]).toBe(40)
    expect(multiplyAlpha(src, -1)[3]).toBe(0)
  })
})

describe('cropRect', () => {
  const rect = [100, 700, 300, 800] as const // 200 wide, 100 tall (PDF y-up)

  it('keeps everything for the identity crop', () => {
    expect(cropRect(rect, { l: 0, t: 0, r: 1, b: 1 })).toEqual([100, 700, 300, 800])
  })

  it('maps display-top fractions to the PDF top edge (y2)', () => {
    // keep the top-left display quadrant: left half, top half
    expect(cropRect(rect, { l: 0, t: 0, r: 0.5, b: 0.5 })).toEqual([100, 750, 200, 800])
    // keep the bottom-right display quadrant
    expect(cropRect(rect, { l: 0.5, t: 0.5, r: 1, b: 1 })).toEqual([200, 700, 300, 750])
  })

  it('keeps an interior window proportional on both axes', () => {
    expect(cropRect(rect, { l: 0.25, t: 0.1, r: 0.75, b: 0.9 })).toEqual([150, 710, 250, 790])
  })
})
