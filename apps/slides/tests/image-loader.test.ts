import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { metafileToDataUrl } from '@genoffice/docx-engine/metafile'
import { createImageLoader } from '../src/renderer/image-loader'

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
})
