import { describe, expect, it } from 'vitest'

import { SHAPE_GALLERY } from '../src/renderer/components/shape-gallery'
import { shapeBackgroundCss, shapePreviewPathD } from '../src/renderer/editor/shape-svg'

const galleryPrsts = SHAPE_GALLERY.flatMap((group) => group.shapes.map((shape) => shape.prst))

describe('shape gallery', () => {
  it('every gallery prst has preview and background geometry', () => {
    const missing = galleryPrsts.filter(
      (prst) =>
        !shapePreviewPathD(prst, 20, 20) || !shapeBackgroundCss(prst, 189, 113, '4472C4', '2F5496'),
    )
    expect(missing).toEqual([])
  })

  it('has no duplicate prsts', () => {
    expect(new Set(galleryPrsts).size).toBe(galleryPrsts.length)
  })

  it('background css is a data-uri svg with fill and stroke', () => {
    const css = shapeBackgroundCss('heart', 189, 113, '4472C4', '2F5496')!
    const encoded = /data:image\/svg\+xml,([^"]+)"/.exec(css)?.[1]
    expect(encoded).toBeTruthy()
    const svg = decodeURIComponent(encoded!)
    expect(svg).toContain('#4472C4')
    expect(svg).toContain('#2F5496')
  })
})
