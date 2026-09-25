// #763 follow-up: media is decoded on demand and retained at a capped size.
// These tests stub the browser pieces (Image, createImageBitmap, canvas, object
// URLs) so the loader's own logic — what it decodes, what it keeps, what it drops —
// is what is being asserted.
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { createImageLoader, decodeCapped } from '../src/renderer/image-loader'

/** Natural sizes the fake decoder reports, keyed by data URL. */
const sizes = new Map<string, { w: number; h: number }>()
const decoded: string[] = []
/** mime given to canvas.toBlob, i.e. what the cap re-encoded into */
const encodeMimes: string[] = []

class FakeImage {
  onload: (() => void) | null = null
  onerror: (() => void) | null = null
  naturalWidth = 0
  naturalHeight = 0
  width = 0
  height = 0
  set src(value: string) {
    const capped = /^blob:capped:(\d+)x(\d+)$/.exec(value)
    const size = capped
      ? { w: Number(capped[1]), h: Number(capped[2]) }
      : sizes.get(value) || { w: 10, h: 10 }
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

/** Every `blob:` url carries its size, so the fake <img> can report the capped size. */
function blobToUrl(blob: Blob): string {
  const marker = (blob as unknown as { __size?: string }).__size
  return marker ? `blob:capped:${marker}` : 'blob:unknown'
}

const settle = async (rounds = 6) => {
  for (let i = 0; i < rounds; i++) await new Promise((r) => setTimeout(r, 0))
}

beforeEach(() => {
  sizes.clear()
  decoded.length = 0
  encodeMimes.length = 0
  vi.stubGlobal('Image', FakeImage)
  vi.stubGlobal('fetch', async (input: unknown) => {
    lastSource = String(input)
    return { blob: async () => ({}) }
  })
  vi.stubGlobal('createImageBitmap', async () => {
    // the bitmap reports the *source* size; the loader must cap it
    const size = sizes.get(lastSource) || { w: 10, h: 10 }
    return { width: size.w, height: size.h, close: () => {} }
  })
  vi.spyOn(URL, 'createObjectURL').mockImplementation(blobToUrl as never)
  vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => {})
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
    () => ({ drawImage: () => {} }) as never,
  )
  vi.spyOn(HTMLCanvasElement.prototype, 'toBlob').mockImplementation(function (
    this: HTMLCanvasElement,
    cb: BlobCallback,
    mime?: string,
  ) {
    encodeMimes.push(String(mime))
    const blob = new Blob(['x']) as Blob & { __size?: string }
    blob.__size = `${this.width}x${this.height}`
    cb(blob)
  } as never)
})

afterEach(() => {
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

/** The url the loader last fetched; the fake decoder sizes itself from it. */
let lastSource = ''

describe('decodeCapped', () => {
  it('caps a large lossy source and returns the capped size', async () => {
    const url = 'data:image/jpeg;base64,AAAA'
    sizes.set(url, { w: 4000, h: 3000 })
    const image = await decodeCapped(url, 2560)
    expect(image).not.toBeNull()
    expect(image!.naturalWidth).toBe(2560)
    expect(image!.naturalHeight).toBe(1920)
  })

  it('leaves a source already within the cap alone', async () => {
    const url = 'data:image/jpeg;base64,BBBB'
    sizes.set(url, { w: 800, h: 600 })
    const image = await decodeCapped(url, 2560)
    expect(image!.naturalWidth).toBe(800)
    expect(decoded.some((u) => u.startsWith('blob:capped:'))).toBe(false)
  })

  it('caps PNG too and keeps it PNG (screenshots are the big ones on the measured deck)', async () => {
    const url = 'data:image/png;base64,CCCC'
    sizes.set(url, { w: 4000, h: 3000 })
    const image = await decodeCapped(url, 2560)
    expect(image!.naturalWidth).toBe(2560)
    expect(image!.naturalHeight).toBe(1920)
    expect(encodeMimes.at(-1)).toBe('image/png') // not re-encoded lossily
  })

  it('uses WebP for the lossy sources', async () => {
    const url = 'data:image/jpeg;base64,EEEE'
    sizes.set(url, { w: 4000, h: 3000 })
    await decodeCapped(url, 2560)
    expect(encodeMimes.at(-1)).toBe('image/webp')
  })

  it('falls back to the plain element when the bitmap pipeline fails', async () => {
    const url = 'data:image/jpeg;base64,DDDD'
    sizes.set(url, { w: 4000, h: 2000 })
    vi.stubGlobal('createImageBitmap', async () => {
      throw new Error('no intrinsic size')
    })
    const image = await decodeCapped(url, 2560)
    // uncapped, i.e. exactly what this module did before the cap existed
    expect(image!.naturalWidth).toBe(4000)
  })
})

describe('createImageLoader', () => {
  it('decodes only the urls it is asked for', async () => {
    const applied: string[] = []
    const loader = createImageLoader((entries) => {
      for (const [url] of entries) applied.push(url)
    })
    const a = 'data:image/png;base64,a'
    const b = 'data:image/png;base64,b'
    sizes.set(a, { w: 20, h: 20 })
    sizes.set(b, { w: 20, h: 20 })

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
    const loader = createImageLoader(() => {}, { maxSide: 100 })
    const url = 'data:image/jpeg;base64,big'
    sizes.set(url, { w: 400, h: 400 })
    loader.load([url])
    await settle()
    expect(loader.bytes()).toBe(100 * 100 * 4)
  })

  it('drops media that left the window once the budget is exceeded', async () => {
    const evicted: string[] = []
    const loader = createImageLoader(() => {}, {
      budgetBytes: 40 * 40 * 4, // one image
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

  it('decodes rail-only media at the rail size and stage media at the cap', async () => {
    const loader = createImageLoader(() => {}, { maxSide: 1000, thumbMaxSide: 100 })
    const railOnly = 'data:image/jpeg;base64,rail-only'
    const onStage = 'data:image/jpeg;base64,on-stage'
    for (const url of [railOnly, onStage]) sizes.set(url, { w: 4000, h: 3000 })

    loader.load([railOnly, onStage], [onStage])
    await settle()
    const stats = loader.stats()
    expect(stats.thumbRetained).toBe(1)
    expect(stats.largeRetained).toBe(1)
    // 100x75 for the rail row + 1000x750 for the stage, not 2x 4000x3000
    expect(stats.retainedBytes).toBe(100 * 75 * 4 + 1000 * 750 * 4)
  })

  it('re-decodes at the stage size when a rail image moves onto the stage', async () => {
    const evicted: string[] = []
    const loader = createImageLoader(() => {}, {
      maxSide: 1000,
      thumbMaxSide: 100,
      onEvict: (url) => evicted.push(url),
    })
    const url = 'data:image/jpeg;base64,promote-me'
    sizes.set(url, { w: 4000, h: 3000 })

    loader.load([url])
    await settle()
    expect(loader.stats().thumbRetained).toBe(1)

    loader.load([url], [url]) // now the slide is current: it is drawn on the stage
    await settle()
    const stats = loader.stats()
    expect(stats.largeRetained).toBe(1)
    expect(stats.thumbRetained).toBe(0)
    expect(stats.decoded).toBe(2)
    expect(stats.retainedBytes).toBe(1000 * 750 * 4)
    expect(evicted).toEqual([url])
  })

  it('forgets the decoded size of an image it evicts', async () => {
    const loader = createImageLoader(() => {}, {
      maxSide: 100,
      thumbMaxSide: 100,
      budgetBytes: 40 * 40 * 4,
    })
    const a = 'data:image/png;base64,gone-a'
    const b = 'data:image/png;base64,stays-b'
    for (const url of [a, b]) sizes.set(url, { w: 40, h: 40 })
    loader.load([a])
    await settle()
    expect(loader.stats().retainedImages).toBe(1)
    loader.load([b]) // a is no longer needed and the budget only fits one
    await settle()
    const stats = loader.stats()
    expect(stats.evicted).toBe(1)
    expect(stats.retainedImages).toBe(1)
    // the bookkeeping must not keep counting an image that is gone
    expect(stats.largeRetained + stats.thumbRetained).toBe(1)
  })

  it('decodes no more than maxConcurrent at a time', async () => {
    const loader = createImageLoader(() => {}, {
      maxConcurrent: 2,
      maxSide: 100,
      thumbMaxSide: 100,
    })
    const urls = [1, 2, 3, 4, 5].map((n) => `data:image/jpeg;base64,c${n}`)
    for (const url of urls) sizes.set(url, { w: 400, h: 400 })

    loader.load(urls)
    expect(loader.pending()).toBe(2)
    expect(loader.stats().queued).toBe(3)
    await settle(20)
    expect(loader.pending()).toBe(0)
    expect(loader.stats().retainedImages).toBe(5)
  })
})
