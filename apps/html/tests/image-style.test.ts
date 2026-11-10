import { describe, expect, it } from 'vitest'
import {
  flipImage,
  identityTransform,
  imageAlignOf,
  imageAlignStyles,
  parseImageTransform,
  rotateImage,
  serializeImageTransform,
} from '../src/renderer/document/image-style'

describe('image transform', () => {
  it('round-trips the authored form and drops the declaration at identity', () => {
    expect(serializeImageTransform(identityTransform())).toBeNull()
    const t = { rotate: 90, flipH: true, flipV: false }
    const s = serializeImageTransform(t)
    expect(s).toBe('rotate(90deg) scale(-1, 1)')
    expect(parseImageTransform(s!)).toEqual(t)
    expect(parseImageTransform('')).toEqual(identityTransform())
    expect(parseImageTransform('translateX(3px)')).toEqual(identityTransform())
    expect(parseImageTransform('scaleX(-1)')).toEqual({ rotate: 0, flipH: true, flipV: false })
  })

  it('rotates in quarter turns and wraps', () => {
    const t = identityTransform()
    expect(rotateImage(t, 1).rotate).toBe(90)
    expect(rotateImage(t, -1).rotate).toBe(270)
    expect(rotateImage({ ...t, rotate: 270 }, 1).rotate).toBe(0)
  })

  it('flips in screen space: a horizontal flip of a rotated image negates the angle', () => {
    const rotated = rotateImage(identityTransform(), 1)
    const flipped = flipImage(rotated, 'h')
    expect(flipped).toEqual({ rotate: 270, flipH: true, flipV: false })
    // flipping twice is a no-op
    expect(flipImage(flipped, 'h')).toEqual(rotated)
    expect(flipImage(identityTransform(), 'v')).toEqual({ rotate: 0, flipH: false, flipV: true })
  })
})

describe('image alignment', () => {
  it('reads back only what it writes', () => {
    for (const a of ['left', 'center', 'right'] as const) {
      const s = imageAlignStyles(a)
      expect(s.display).toBe('block')
      expect(imageAlignOf(s['margin-left']!, s['margin-right']!)).toBe(a)
    }
    expect(imageAlignOf('', '')).toBeNull()
    expect(imageAlignOf('12px', '0')).toBeNull()
  })
})
