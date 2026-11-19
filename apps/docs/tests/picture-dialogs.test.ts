import { describe, expect, it } from 'vitest'
import { fitCropPreview } from '../src/renderer/components/PictureDialogs'

describe('crop preview sizing', () => {
  it('contains wide and tall images within both available dimensions', () => {
    expect(fitCropPreview(4000, 1000, 500, 300)).toEqual({ w: 500, h: 125 })
    expect(fitCropPreview(1000, 4000, 500, 300)).toEqual({ w: 75, h: 300 })
  })

  it('does not enlarge small images', () => {
    expect(fitCropPreview(120, 80, 500, 300)).toEqual({ w: 120, h: 80 })
  })
})
