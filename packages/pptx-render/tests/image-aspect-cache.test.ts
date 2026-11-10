import { beforeEach, describe, expect, it } from 'vitest'
import { cacheKeyFor } from '../src/image-dpi'
import {
  IMAGE_ASPECT_CACHE_MAX,
  clearImageAspectCache,
  imageAspect,
  imageAspectCacheKeys,
  imageAspectCacheSize,
} from '../src/text-layout'

function pngDataUrl(width: number, height: number): string {
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
    0,
    0,
    0,
    13,
    0x49,
    0x48,
    0x44,
    0x52,
    ...be32(width),
    ...be32(height),
  ]
  return `data:image/png;base64,${Buffer.from(bytes).toString('base64')}`
}

function gifDataUrl(width: number, height: number): string {
  const le16 = (v: number) => [v & 255, (v >>> 8) & 255]
  const bytes = [0x47, 0x49, 0x46, 0x38, 0x39, 0x61, ...le16(width), ...le16(height)]
  return `data:image/gif;base64,${Buffer.from(bytes).toString('base64')}`
}

describe('image aspect cache', () => {
  beforeEach(() => {
    clearImageAspectCache()
  })

  it('still reads the aspect ratio from PNG and GIF headers', () => {
    expect(imageAspect(pngDataUrl(200, 100))).toBeCloseTo(2, 6)
    expect(imageAspect(gifDataUrl(64, 32))).toBeCloseTo(2, 6)
    // unrecognised / non-base64 input falls back to 1
    expect(imageAspect('https://example.com/a.png')).toBe(1)
    expect(imageAspect('data:image/png;base64,AAAA')).toBe(1)
  })

  it('hits the cache for an identical data URL without growing the map', () => {
    const url = pngDataUrl(200, 100)
    const first = imageAspect(url)
    expect(first).toBeCloseTo(2, 6)
    expect(imageAspectCacheSize()).toBe(1)
    expect(imageAspect(url)).toBeCloseTo(first, 12)
    expect(imageAspectCacheSize()).toBe(1)
  })

  it('does not collide across distinct test vectors', () => {
    const urls = [
      pngDataUrl(200, 100),
      pngDataUrl(100, 200),
      gifDataUrl(64, 32),
      gifDataUrl(32, 64),
      'data:image/png;base64,AAAA',
      'https://example.com/a.png',
    ]
    expect(new Set(urls.map((u) => cacheKeyFor(u))).size).toBe(urls.length)
    urls.forEach((u) => imageAspect(u))
    expect(imageAspectCacheSize()).toBe(urls.length)
  })

  it('keeps the size bounded after many distinct inserts and evicts the oldest', () => {
    const total = IMAGE_ASPECT_CACHE_MAX + 50
    const first = pngDataUrl(2, 1)
    imageAspect(first)
    const firstKey = cacheKeyFor(first)
    for (let i = 0; i < total; i++) imageAspect(`data:image/png;base64,evict-me-${i}`)
    expect(imageAspectCacheSize()).toBe(IMAGE_ASPECT_CACHE_MAX)
    expect(imageAspectCacheKeys()).not.toContain(firstKey)
    expect(imageAspectCacheKeys()).toContain(
      cacheKeyFor(`data:image/png;base64,evict-me-${total - 1}`),
    )
  })

  it('a cache hit refreshes recency so the least recently used entry is evicted', () => {
    const kept = pngDataUrl(4, 1)
    const dropped = pngDataUrl(5, 1)
    imageAspect(kept)
    imageAspect(dropped)
    // touching `kept` makes `dropped` the oldest entry
    imageAspect(kept)
    expect(imageAspectCacheKeys().indexOf(cacheKeyFor(kept))).toBeGreaterThan(
      imageAspectCacheKeys().indexOf(cacheKeyFor(dropped)),
    )
    // one insert past the cap now evicts `dropped`, not the entry just used
    for (let i = 0; i < IMAGE_ASPECT_CACHE_MAX - 1; i++) {
      imageAspect(`data:image/png;base64,filler-${i}`)
    }
    const keys = imageAspectCacheKeys()
    expect(keys).toContain(cacheKeyFor(kept))
    expect(keys).not.toContain(cacheKeyFor(dropped))
  })

  it('never retains full data URL strings as keys', () => {
    const large = `data:image/png;base64,${'A'.repeat(200_000)}B`
    imageAspect(large)
    const keys = imageAspectCacheKeys()
    expect(keys).toHaveLength(1)
    expect(keys[0]).not.toBe(large)
    expect(keys[0]).toBe(cacheKeyFor(large))
    expect(keys[0]).toMatch(/^\d+:[0-9a-f]{8}$/)
    expect((keys[0] as string).length).toBeLessThan(32)
  })

  it('two URLs sharing a hashed head but differing past it resolve to the same aspect', () => {
    // The aspect comes from the header region the key hashes, so two URLs that
    // differ only past that region share a key and cannot disagree on the answer.
    const shared = `${pngDataUrl(120, 60)}${'A'.repeat(99_000)}`
    const a = `${shared}X`
    const b = `${shared}Y`
    expect(cacheKeyFor(a)).toBe(cacheKeyFor(b))
    expect(imageAspect(a)).toBeCloseTo(2, 6)
    expect(imageAspect(b)).toBeCloseTo(2, 6)
    expect(imageAspectCacheSize()).toBe(1)
  })
})
