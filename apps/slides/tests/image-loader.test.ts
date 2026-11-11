import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { metafileToDataUrl } from '@genoffice/docx-engine/metafile'
import {
  createImageLoader,
  MAX_METAFILE_BASE64_CHARS,
  setImageSource,
} from '../src/renderer/image-loader'

vi.mock('@genoffice/docx-engine/metafile', () => ({
  metafileToDataUrl: vi.fn(async () => 'data:image/png;base64,AA=='),
}))

class FakeImage {
  static instances: FakeImage[] = []
  onload: (() => void) | null = null
  onerror: (() => void) | null = null
  src = ''
  constructor() {
    FakeImage.instances.push(this)
  }
}

const img = (src: string) => FakeImage.instances.find((i) => i.src === src)!
const flushed = (apply: ReturnType<typeof vi.fn>) =>
  apply.mock.calls.map((c) => (c[0] as [string, unknown][]).map(([k]) => k))

describe('createImageLoader', () => {
  beforeEach(() => {
    FakeImage.instances = []
    vi.stubGlobal('Image', FakeImage)
    vi.useFakeTimers()
  })
  afterEach(() => {
    vi.unstubAllGlobals()
    vi.useRealTimers()
  })

  it('surfaces loaded images by timer batches while others are still pending', () => {
    const apply = vi.fn()
    const loader = createImageLoader(apply, 16, 100)
    loader.load(['a', 'b', 'c'])
    img('a').onload!()
    expect(apply).not.toHaveBeenCalled()
    vi.advanceTimersByTime(100)
    expect(flushed(apply)).toEqual([['a']])
  })

  it('flushes immediately when the batch size is reached', () => {
    const apply = vi.fn()
    const loader = createImageLoader(apply, 2, 100)
    loader.load(['a', 'b', 'c'])
    img('a').onload!()
    img('b').onload!()
    expect(flushed(apply)).toEqual([['a', 'b']])
  })

  it('flushes the remainder when the last pending image settles, even on error', () => {
    const apply = vi.fn()
    const loader = createImageLoader(apply, 16, 100)
    loader.load(['a', 'b'])
    img('a').onload!()
    img('b').onerror!()
    expect(flushed(apply)).toEqual([['a']])
  })

  it('never reloads finished urls nor discards in-flight ones on a new load call', () => {
    const apply = vi.fn()
    const loader = createImageLoader(apply, 16, 100)
    loader.load(['a', 'b'])
    img('a').onload!()
    vi.advanceTimersByTime(100)
    loader.load(['a', 'b', 'c'])
    expect(FakeImage.instances.map((i) => i.src)).toEqual(['a', 'b', 'c'])
    img('b').onload!()
    img('c').onload!()
    expect(flushed(apply)).toEqual([['a'], ['b', 'c']])
  })

  it('stops applying after dispose', () => {
    const apply = vi.fn()
    const loader = createImageLoader(apply, 16, 100)
    loader.load(['a'])
    loader.dispose()
    img('a').onload!()
    vi.runAllTimers()
    expect(apply).not.toHaveBeenCalled()
  })

  it('never re-decodes a url whose decode failed', () => {
    const apply = vi.fn()
    const loader = createImageLoader(apply, 16, 100)
    loader.load(['broken'])
    img('broken').onerror!()
    // An edit/zoom/slide change re-collects the same url: it must not be re-queued
    loader.load(['broken', 'broken'])
    loader.load(['broken'])
    expect(FakeImage.instances.map((i) => i.src)).toEqual(['broken'])
    expect(loader.pending()).toBe(0)
    expect(apply).not.toHaveBeenCalled()
  })
})

describe('metafile rasterization waits for private fonts', () => {
  const flag = window as { __genofficeDocFontsSynced?: boolean }
  beforeEach(() => {
    FakeImage.instances = []
    vi.stubGlobal('Image', FakeImage)
    vi.useFakeTimers()
    vi.stubGlobal('atob', (b: string) => Buffer.from(b, 'base64').toString('binary'))
    vi.mocked(metafileToDataUrl).mockClear()
  })
  afterEach(() => {
    vi.unstubAllGlobals()
    vi.useRealTimers()
    delete flag.__genofficeDocFontsSynced
  })

  it('holds the EMF until the doc-fonts sync flag flips, then rasterizes', async () => {
    flag.__genofficeDocFontsSynced = false
    const loader = createImageLoader(vi.fn(), 16, 100)
    loader.load(['data:image/x-emf;base64,AQAAAA=='])
    await vi.advanceTimersByTimeAsync(300)
    expect(metafileToDataUrl).not.toHaveBeenCalled()
    flag.__genofficeDocFontsSynced = true
    await vi.advanceTimersByTimeAsync(100)
    expect(metafileToDataUrl).toHaveBeenCalledTimes(1)
  })

  it('does not wait when no sync has started', async () => {
    const loader = createImageLoader(vi.fn(), 16, 100)
    loader.load(['data:image/x-emf;base64,AQAAAA=='])
    await vi.advanceTimersByTimeAsync(0)
    expect(metafileToDataUrl).toHaveBeenCalledTimes(1)
  })

  it('fetches lazy EMF bytes and rasterizes them under the original URL key', async () => {
    const bytes = new Uint8Array([1, 0, 0, 0])
    const url = 'genoffice-slide-media://archive/test?mime=image%2Fx-emf'
    const fetchMedia = vi.fn(async () => ({
      ok: true,
      headers: new Headers({ 'Content-Length': String(bytes.length) }),
      arrayBuffer: async () => bytes.buffer,
    }))
    vi.stubGlobal('fetch', fetchMedia)
    const apply = vi.fn()
    const loader = createImageLoader(apply)
    loader.load([url])
    await vi.advanceTimersByTimeAsync(0)
    expect(fetchMedia).toHaveBeenCalledWith(url)
    expect(metafileToDataUrl).toHaveBeenCalledWith(bytes, 'image/x-emf', { maxSidePx: 2048 })
    img('data:image/png;base64,AA==').onload!()
    expect(flushed(apply)).toEqual([[url]])
  })

  it('refuses oversized lazy metafiles from the response length before reading bytes', async () => {
    const arrayBuffer = vi.fn()
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => ({
        ok: true,
        headers: new Headers({ 'Content-Length': String(MAX_METAFILE_BASE64_CHARS) }),
        arrayBuffer,
      })),
    )
    const loader = createImageLoader(vi.fn())
    loader.load(['genoffice-slide-media://archive/test?mime=image%2Fx-wmf'])
    await vi.advanceTimersByTimeAsync(0)
    expect(arrayBuffer).not.toHaveBeenCalled()
    expect(metafileToDataUrl).not.toHaveBeenCalled()
    expect(loader.pending()).toBe(0)
  })

  it('refuses oversized metafiles before base64 decoding', async () => {
    const apply = vi.fn()
    const loader = createImageLoader(apply, 16, 100)
    const huge = `data:image/x-emf;base64,${'A'.repeat(MAX_METAFILE_BASE64_CHARS + 1)}`
    loader.load([huge])
    await vi.advanceTimersByTimeAsync(0)
    await vi.advanceTimersByTimeAsync(0)
    expect(metafileToDataUrl).not.toHaveBeenCalled()
    expect(apply).not.toHaveBeenCalled()
  })
})

it('sets anonymous CORS before assigning slide media sources only', () => {
  const assignments: Array<[string, string | null]> = []
  const image = {
    crossOrigin: null as string | null,
    set src(url: string) {
      assignments.push([url, this.crossOrigin])
    },
  }
  setImageSource(image as HTMLImageElement, 'genoffice-slide-media://archive/image')
  const other = {
    crossOrigin: null as string | null,
    set src(url: string) {
      assignments.push([url, this.crossOrigin])
    },
  }
  setImageSource(other as HTMLImageElement, 'data:image/png;base64,AA==')
  expect(assignments).toEqual([
    ['genoffice-slide-media://archive/image', 'anonymous'],
    ['data:image/png;base64,AA==', null],
  ])
})
