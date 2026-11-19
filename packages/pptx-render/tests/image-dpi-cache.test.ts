import { beforeEach, describe, expect, it } from 'vitest'
import {
  cacheKeyFor,
  clearImageDpiCache,
  IMAGE_DPI_CACHE_MAX,
  imageDpiCacheKeys,
  imageDpiCacheSize,
  imageDpiFromDataUrl,
} from '../src/image-dpi'

function pngWithPhys(ppm: number): string {
  const chunk = (type: string, data: number[]) => {
    const len = data.length
    return [
      (len >>> 24) & 255,
      (len >>> 16) & 255,
      (len >>> 8) & 255,
      len & 255,
      ...type.split('').map((c) => c.charCodeAt(0)),
      ...data,
      0,
      0,
      0,
      0,
    ]
  }
  const be32 = (v: number) => [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255]
  const bytes = [
    0x89,
    0x50,
    0x4e,
    0x47,
    0x0d,
    0x0a,
    0x1a,
    0x0a,
    ...chunk('IHDR', [...be32(2), ...be32(2), 8, 6, 0, 0, 0]),
    ...chunk('pHYs', [...be32(ppm), ...be32(ppm), 1]),
    ...chunk('IDAT', [0, 0]),
    ...chunk('IEND', []),
  ]
  return `data:image/png;base64,${Buffer.from(bytes).toString('base64')}`
}

describe('image dpi cache keys', () => {
  beforeEach(() => {
    clearImageDpiCache()
  })

  it('hits the cache for identical bytes without growing the map', () => {
    const url = pngWithPhys(5906)
    const first = imageDpiFromDataUrl(url)
    expect(first?.x).toBeCloseTo(150, 0)
    const sizeAfterFirst = imageDpiCacheSize()
    expect(sizeAfterFirst).toBe(1)
    const second = imageDpiFromDataUrl(url)
    expect(second).toEqual(first)
    expect(imageDpiCacheSize()).toBe(sizeAfterFirst)
  })

  it('does not collide for distinct test vectors', () => {
    const urls = [
      pngWithPhys(5906),
      pngWithPhys(2953),
      'data:image/png;base64,AAAA',
      'https://example.com/a.png',
    ]
    const keys = urls.map((u) => cacheKeyFor(u))
    expect(new Set(keys).size).toBe(urls.length)
    // One-char difference still separates keys.
    expect(cacheKeyFor('data:image/png;base64,AAAA')).not.toBe(
      cacheKeyFor('data:image/png;base64,AAAB'),
    )
  })

  it('keeps size bounded after many distinct inserts and evicts the oldest', () => {
    const total = IMAGE_DPI_CACHE_MAX + 50
    const firstUrl = 'data:image/png;base64,evict-me-0'
    imageDpiFromDataUrl(firstUrl)
    const firstKey = cacheKeyFor(firstUrl)
    for (let i = 1; i < total; i++) {
      imageDpiFromDataUrl(`data:image/png;base64,evict-me-${i}`)
    }
    expect(imageDpiCacheSize()).toBeLessThanOrEqual(IMAGE_DPI_CACHE_MAX)
    expect(imageDpiCacheSize()).toBe(IMAGE_DPI_CACHE_MAX)
    expect(imageDpiCacheKeys()).not.toContain(firstKey)
    expect(imageDpiCacheKeys()).toContain(
      cacheKeyFor(`data:image/png;base64,evict-me-${total - 1}`),
    )
  })

  it('never retains full data URL strings as keys', () => {
    const large = `data:image/png;base64,${'A'.repeat(200_000)}B`
    imageDpiFromDataUrl(large)
    const keys = imageDpiCacheKeys()
    expect(keys).toHaveLength(1)
    expect(keys[0]).not.toBe(large)
    expect(keys[0]).toBe(cacheKeyFor(large))
    expect(keys[0]).toMatch(/^\d+:[0-9a-f]{8}$/)
    expect((keys[0] as string).length).toBeLessThan(32)
    // Non-data URLs use the same compact key shape.
    clearImageDpiCache()
    const plain = 'https://example.com/image.png'
    expect(imageDpiFromDataUrl(plain)).toBeUndefined()
    expect(imageDpiCacheKeys()).toEqual([cacheKeyFor(plain)])
  })
})
