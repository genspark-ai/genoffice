import { describe, it, expect } from 'vitest'
import { filterUsableImages, MIN_USABLE_IMAGE_PX } from '../src/index'
import type { ImageSearchResult } from '../src/shared'

const img = (over: Partial<ImageSearchResult> = {}): ImageSearchResult => ({
  title: 't',
  imageUrl: 'https://example.com/a.jpg',
  sourceUrl: 'https://example.com/',
  source: 'example.com',
  ...over,
})

describe('filterUsableImages', () => {
  it('drops images a backend reports smaller than the usable floor', () => {
    const r = filterUsableImages([
      img({ width: 64, height: 64 }), // icon
      img({ width: 180, height: 400 }), // narrow strip
      img({ width: 800, height: 600 }),
      img({ width: 1920, height: MIN_USABLE_IMAGE_PX }),
    ])
    expect(r.map((x) => x.width)).toEqual([800, 1920])
  })

  it('keeps entries without dimension metadata (no evidence, no verdict)', () => {
    const noMeta = img()
    expect(filterUsableImages([noMeta])).toEqual([noMeta])
    // one axis known and usable, the other unknown: keep
    expect(filterUsableImages([img({ width: 1024 })]).length).toBe(1)
    // one axis known and too small: drop
    expect(filterUsableImages([img({ height: 90 })]).length).toBe(0)
  })

  it('keeps images exactly at the floor', () => {
    expect(
      filterUsableImages([img({ width: MIN_USABLE_IMAGE_PX, height: MIN_USABLE_IMAGE_PX })]).length,
    ).toBe(1)
  })
})
