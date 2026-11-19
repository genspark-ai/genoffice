/**
 * Unit tests for the cutout (background removal) core algorithm: background color sampling / tolerance threshold / edge flood fill.
 */
import { describe, expect, it } from 'vitest'
import {
  colorDistance,
  removeBackground,
  sampleBackgroundColors,
  toleranceToThreshold,
  type PixelImage,
  type RGB,
} from '../src/renderer/cutout'

/** Build a solid-color base image (RGBA, alpha 255) */
function makeImage(width: number, height: number, fill: RGB): PixelImage {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let i = 0; i < width * height; i++) {
    data[i * 4] = fill[0]
    data[i * 4 + 1] = fill[1]
    data[i * 4 + 2] = fill[2]
    data[i * 4 + 3] = 255
  }
  return { data, width, height }
}

/** Draw a filled rectangle on the image (inclusive of x0/y0, exclusive of x1/y1) */
function fillRect(
  img: PixelImage,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  color: RGB,
  alpha = 255,
) {
  for (let y = y0; y < y1; y++) {
    for (let x = x0; x < x1; x++) {
      const p = (y * img.width + x) * 4
      img.data[p] = color[0]
      img.data[p + 1] = color[1]
      img.data[p + 2] = color[2]
      img.data[p + 3] = alpha
    }
  }
}

function alphaAt(data: Uint8ClampedArray, img: PixelImage, x: number, y: number): number {
  return data[(y * img.width + x) * 4 + 3]!
}

const WHITE: RGB = [255, 255, 255]
const RED: RGB = [200, 30, 30]

describe('toleranceToThreshold', () => {
  it('0 → 0, increases monotonically with tolerance, clamps out-of-range values', () => {
    expect(toleranceToThreshold(0)).toBe(0)
    expect(toleranceToThreshold(50)).toBeGreaterThan(toleranceToThreshold(10))
    expect(toleranceToThreshold(-5)).toBe(0)
    expect(toleranceToThreshold(999)).toBe(toleranceToThreshold(100))
  })

  it('tolerance 100 covers the black-white distance (whole image classifiable as background)', () => {
    expect(toleranceToThreshold(100)).toBeGreaterThanOrEqual(
      colorDistance([0, 0, 0], [255, 255, 255]),
    )
  })
})

describe('colorDistance', () => {
  it('same color is 0, black vs white is 255, symmetric', () => {
    expect(colorDistance(RED, RED)).toBe(0)
    expect(colorDistance([0, 0, 0], [255, 255, 255])).toBe(255)
    expect(colorDistance(WHITE, RED)).toBeCloseTo(colorDistance(RED, WHITE))
  })
})

describe('sampleBackgroundColors', () => {
  it('solid background + centered foreground → single representative background color', () => {
    const img = makeImage(20, 20, WHITE)
    fillRect(img, 5, 5, 15, 15, RED)
    const colors = sampleBackgroundColors(img)
    expect(colors).toHaveLength(1)
    expect(colorDistance(colors[0]!, WHITE)).toBeLessThan(1)
  })

  it('two-color top/bottom border → two representative colors, dominant one first', () => {
    const img = makeImage(20, 20, WHITE)
    // Bottom quarter is black (entire bottom edge + lower parts of left/right edges are black)
    fillRect(img, 0, 15, 20, 20, [0, 0, 0])
    const colors = sampleBackgroundColors(img)
    expect(colors.length).toBe(2)
    // White has more border samples -> ranks first
    expect(colorDistance(colors[0]!, WHITE)).toBeLessThan(1)
    expect(colorDistance(colors[1]!, [0, 0, 0])).toBeLessThan(1)
  })

  it('fully transparent border yields no representative colors', () => {
    const img = makeImage(10, 10, WHITE)
    for (let i = 0; i < 100; i++) img.data[i * 4 + 3] = 0
    expect(sampleBackgroundColors(img)).toHaveLength(0)
  })

  it('empty image returns an empty array', () => {
    expect(
      sampleBackgroundColors({ data: new Uint8ClampedArray(0), width: 0, height: 0 }),
    ).toHaveLength(0)
  })
})

describe('removeBackground', () => {
  it('background made transparent, foreground kept, removedCount correct', () => {
    const img = makeImage(10, 10, WHITE)
    fillRect(img, 3, 3, 7, 7, RED) // 4×4 foreground
    const { data, removedCount } = removeBackground(img, 30)
    expect(removedCount).toBe(100 - 16)
    expect(alphaAt(data, img, 0, 0)).toBe(0) // background corner
    expect(alphaAt(data, img, 5, 5)).toBe(255) // foreground center
    expect(alphaAt(data, img, 3, 3)).toBe(255) // foreground edge
  })

  it('flood fill: enclosed regions matching the background but surrounded by foreground are not removed', () => {
    const img = makeImage(12, 12, WHITE)
    fillRect(img, 3, 3, 9, 9, RED) // solid red block
    fillRect(img, 5, 5, 7, 7, WHITE) // carve a white hole in the center (same color as background but not connected)
    const { data } = removeBackground(img, 30)
    expect(alphaAt(data, img, 0, 0)).toBe(0) // outer white -> transparent
    expect(alphaAt(data, img, 5, 5)).toBe(255) // inner white -> kept
    expect(alphaAt(data, img, 4, 4)).toBe(255) // red ring kept
  })

  it('tolerance takes effect: near-background pixels kept at low tolerance, removed at high tolerance', () => {
    const img = makeImage(10, 10, WHITE)
    fillRect(img, 3, 3, 7, 7, RED)
    // Place a near-white pixel outside the foreground (connected to the edge background)
    fillRect(img, 1, 1, 2, 2, [240, 240, 240])
    const low = removeBackground(img, 0)
    expect(alphaAt(low.data, img, 1, 1)).toBe(255) // tolerance 0: only pure white removed
    const high = removeBackground(img, 30)
    expect(alphaAt(high.data, img, 1, 1)).toBe(0) // tolerance 30: near-white removed too
  })

  it('tolerance 100: all connected background across the image becomes transparent', () => {
    const img = makeImage(8, 8, WHITE)
    fillRect(img, 2, 2, 6, 6, RED)
    const { removedCount } = removeBackground(img, 100)
    expect(removedCount).toBe(64)
  })

  it('does not mutate input pixels', () => {
    const img = makeImage(6, 6, WHITE)
    const before = Array.from(img.data)
    removeBackground(img, 50)
    expect(Array.from(img.data)).toEqual(before)
  })

  it('already-transparent pixels count as connected background and are not double-counted', () => {
    const img = makeImage(10, 10, WHITE)
    fillRect(img, 3, 3, 7, 7, RED)
    // A top strip already transparent (the flood still reaches the background below through it)
    fillRect(img, 0, 0, 10, 1, [9, 9, 9], 0)
    const { data, removedCount } = removeBackground(img, 30)
    expect(alphaAt(data, img, 0, 5)).toBe(0) // background below still flooded
    expect(removedCount).toBe(100 - 16 - 10) // the 10 already-transparent pixels aren't counted
    expect(alphaAt(data, img, 5, 5)).toBe(255)
  })

  it('explicit background color param: uses the given color instead of auto-sampling', () => {
    const img = makeImage(10, 10, RED) // border is red
    fillRect(img, 4, 4, 6, 6, WHITE)
    // White specified as background -> red edges don't pass, nothing is removed
    const { removedCount } = removeBackground(img, 10, [WHITE])
    expect(removedCount).toBe(0)
  })
})
