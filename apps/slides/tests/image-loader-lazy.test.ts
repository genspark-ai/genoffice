// #763 follow-up: media is decoded on demand, bounded by *how much* is kept, and
// released once it leaves the window.
//
// There is deliberately no downscaling here: decoding smaller than the source is a
// fidelity trade the maintainers should decide on, not something a perf change should
// slip in. As configured it bought ~5% of process memory on the measured deck and
// made fast scrolling visibly worse, so it lives in a parked branch
// (perf/slides-media-cap-parked) until someone asks for it.
//
// These tests cover what is purely about when and how many images are decoded.
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createImageLoader } from '../src/renderer/image-loader'

/** Natural sizes the fake browser reports, keyed by url. */
const sizes = new Map<string, { w: number; h: number }>()
const decoded: string[] = []

class FakeImage {
  onload: (() => void) | null = null
  onerror: (() => void) | null = null
  naturalWidth = 0
  naturalHeight = 0
  width = 0
  height = 0
  set src(value: string) {
    const size = sizes.get(value) || { w: 10, h: 10 }
    decoded.push(value)
    this.naturalWidth = size.w
    this.naturalHeight = size.h
    this.width = size.w
    this.height = size.h
    queueMicrotask(() => this.onload?.())
  }
  get src(): string {
    return ''
  }
}

const settle = async (rounds = 6) => {
  for (let i = 0; i < rounds; i++) await new Promise((r) => setTimeout(r, 0))
}

beforeEach(() => {
  sizes.clear()
  decoded.length = 0
  vi.stubGlobal('Image', FakeImage)
})

afterEach(() => {
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

describe('createImageLoader', () => {
  it('decodes only the urls it is asked for', async () => {
    const applied: string[] = []
    const loader = createImageLoader((entries) => {
      for (const [url] of entries) applied.push(url)
    })
    const a = 'data:image/png;base64,a'
    const b = 'data:image/png;base64,b'
    for (const url of [a, b]) sizes.set(url, { w: 20, h: 20 })

    loader.load([a])
    await settle()
    expect(applied).toEqual([a])
  })

  it('does not decode the same url twice', async () => {
    const loader = createImageLoader(() => {})
    const url = 'data:image/png;base64,same'
    sizes.set(url, { w: 20, h: 20 })
    loader.load([url])
    await settle()
    const first = decoded.length
    loader.load([url])
    await settle()
    expect(decoded.length).toBe(first)
  })

  it('reports the decoded bytes it is holding', async () => {
    const loader = createImageLoader(() => {})
    const url = 'data:image/png;base64,big'
    sizes.set(url, { w: 400, h: 400 })
    loader.load([url])
    await settle()
    expect(loader.bytes()).toBe(400 * 400 * 4)
  })

  it('drops media that left the window once the budget is exceeded', async () => {
    const evicted: string[] = []
    const loader = createImageLoader(() => {}, {
      budgetBytes: 40 * 40 * 4, // room for exactly one image
      onEvict: (url) => evicted.push(url),
    })
    const first = 'data:image/png;base64,one'
    const second = 'data:image/png;base64,two'
    for (const url of [first, second]) sizes.set(url, { w: 40, h: 40 })

    loader.load([first])
    await settle()
    expect(loader.bytes()).toBe(40 * 40 * 4)

    // the deck moved on: only the second slide is needed now
    loader.load([second])
    await settle()
    expect(evicted).toEqual([first])
    expect(loader.bytes()).toBeLessThanOrEqual(40 * 40 * 4)
  })

  it('never evicts what the current window still needs', async () => {
    const evicted: string[] = []
    const loader = createImageLoader(() => {}, {
      budgetBytes: 40 * 40 * 4,
      onEvict: (url) => evicted.push(url),
    })
    const a = 'data:image/png;base64,keep-a'
    const b = 'data:image/png;base64,keep-b'
    for (const url of [a, b]) sizes.set(url, { w: 40, h: 40 })
    loader.load([a, b])
    await settle()
    expect(evicted).toEqual([])
  })

  it('forgets an evicted image instead of counting it forever', async () => {
    const loader = createImageLoader(() => {}, { budgetBytes: 40 * 40 * 4 })
    const a = 'data:image/png;base64,gone-a'
    const b = 'data:image/png;base64,stays-b'
    for (const url of [a, b]) sizes.set(url, { w: 40, h: 40 })
    loader.load([a])
    await settle()
    expect(loader.stats().retainedImages).toBe(1)
    loader.load([b])
    await settle()
    const stats = loader.stats()
    expect(stats.evicted).toBe(1)
    expect(stats.retainedImages).toBe(1)
    expect(stats.retainedBytes).toBe(40 * 40 * 4)
  })

  it('decodes no more than maxConcurrent at a time', async () => {
    const loader = createImageLoader(() => {}, { maxConcurrent: 2 })
    const urls = [1, 2, 3, 4, 5].map((n) => `data:image/png;base64,c${n}`)
    for (const url of urls) sizes.set(url, { w: 40, h: 40 })

    loader.load(urls)
    expect(loader.pending()).toBe(2)
    expect(loader.stats().queued).toBe(3)
    await settle(20)
    expect(loader.pending()).toBe(0)
    expect(loader.stats().retainedImages).toBe(5)
  })
})
