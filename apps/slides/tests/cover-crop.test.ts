/** AI image insertion must never stretch: coverCropFractions (object-fit: cover math). */
import { describe, it, expect } from 'vitest'
import { coverCropFractions } from '../src/shared/cover-crop'

describe('coverCropFractions', () => {
  it('returns null when the aspect ratios already match', () => {
    expect(coverCropFractions(1600, 900, 640, 360)).toBeNull()
    expect(coverCropFractions(500, 500, 200, 200)).toBeNull()
  })

  it('tolerates sub-1% aspect mismatch (rounding noise)', () => {
    expect(coverCropFractions(1601, 900, 640, 360)).toBeNull()
  })

  it('crops left/right when the image is wider than the frame', () => {
    // 2:1 image into a 1:1 frame → half the width goes, split evenly
    const crop = coverCropFractions(2000, 1000, 400, 400)
    expect(crop).toEqual({ l: 0.25, t: 0, r: 0.25, b: 0 })
  })

  it('crops top/bottom when the image is taller than the frame', () => {
    // 1:2 image into a 1:1 frame → half the height goes, split evenly
    const crop = coverCropFractions(1000, 2000, 400, 400)
    expect(crop).toEqual({ l: 0, t: 0.25, r: 0, b: 0.25 })
  })

  it('handles the common square-generated-image into 16:9 banner case', () => {
    const crop = coverCropFractions(1024, 1024, 1280, 720)
    expect(crop!.l).toBe(0)
    expect(crop!.r).toBe(0)
    expect(crop!.t).toBeCloseTo((1 - 720 / 1280) / 2, 6)
    expect(crop!.t).toBe(crop!.b)
    // remaining region has exactly the frame's aspect ratio
    const remainingH = 1024 * (1 - crop!.t - crop!.b)
    expect(1024 / remainingH).toBeCloseTo(1280 / 720, 6)
  })

  it('returns null on degenerate sizes (decode failure yields 0×0)', () => {
    expect(coverCropFractions(0, 0, 400, 300)).toBeNull()
    expect(coverCropFractions(800, 600, 0, 300)).toBeNull()
  })
})
