import { afterEach, describe, expect, it, vi } from 'vitest'
import UTIF from 'utif2'
import { tiffDimsOk, tiffToDataUrl } from '../src/tiff'

let maxCanvasDim = 0

function stubDom() {
  maxCanvasDim = 0
  const canvasStub = () => {
    let w = 0
    let h = 0
    return {
      get width() {
        return w
      },
      set width(v: number) {
        w = v
        maxCanvasDim = Math.max(maxCanvasDim, v)
      },
      get height() {
        return h
      },
      set height(v: number) {
        h = v
        maxCanvasDim = Math.max(maxCanvasDim, v)
      },
      getContext: () => ({
        putImageData: () => {},
      }),
      toDataURL: () => 'data:image/png;base64,AAAA',
    }
  }
  vi.stubGlobal('document', { createElement: canvasStub })
  vi.stubGlobal(
    'ImageData',
    class {
      constructor(
        public data: unknown,
        public w: number,
        public h: number,
      ) {}
    },
  )
}

afterEach(() => {
  vi.unstubAllGlobals()
})

function giantIfd(width: number, height: number): Uint8Array {
  const buf = new ArrayBuffer(8 + 2 + 2 * 12 + 4)
  const view = new DataView(buf)
  view.setUint8(0, 0x49)
  view.setUint8(1, 0x49)
  view.setUint16(2, 42, true)
  view.setUint32(4, 8, true)
  view.setUint16(8, 2, true)
  const entry = (i: number, tag: number, value: number) => {
    const off = 10 + i * 12
    view.setUint16(off, tag, true)
    view.setUint16(off + 2, 4, true)
    view.setUint32(off + 4, 1, true)
    view.setUint32(off + 8, value, true)
  }
  entry(0, 256, width)
  entry(1, 257, height)
  view.setUint32(10 + 2 * 12, 0, true)
  return new Uint8Array(buf)
}

/** Prepends an over-budget IFD so it becomes ifds[0] and the real page ifds[1]. */
function prependGiantIfd(tiff: Uint8Array, width: number, height: number): Uint8Array {
  const le = tiff[0] === 0x49
  const base = tiff.length + (tiff.length % 2)
  const out = new Uint8Array(base + 2 + 2 * 12 + 4)
  out.set(tiff)
  const view = new DataView(out.buffer)
  const firstIfd = view.getUint32(4, le)
  view.setUint32(4, base, le)
  view.setUint16(base, 2, le)
  const entry = (i: number, tag: number, value: number) => {
    const off = base + 2 + i * 12
    view.setUint16(off, tag, le)
    view.setUint16(off + 2, 4, le)
    view.setUint32(off + 4, 1, le)
    view.setUint32(off + 8, value, le)
  }
  entry(0, 256, width)
  entry(1, 257, height)
  view.setUint32(base + 2 + 2 * 12, firstIfd, le)
  return out
}

describe('tiffToDataUrl dimension guard', () => {
  it('transcodes a small TIFF', () => {
    stubDom()
    const rgba = new Uint8Array(4 * 3 * 4).fill(128)
    const bytes = new Uint8Array(UTIF.encodeImage(rgba, 4, 3))
    expect(tiffToDataUrl(bytes)).toBe('data:image/png;base64,AAAA')
    expect(maxCanvasDim).toBe(4)
  })

  it('rejects giant IFD dims without allocating canvas pixels', () => {
    stubDom()
    const start = Date.now()
    expect(tiffToDataUrl(giantIfd(30000, 30000))).toBeNull()
    expect(Date.now() - start).toBeLessThan(10000)
    expect(maxCanvasDim).toBe(0)
  })

  it('budget admits 1200 dpi Letter and 600 dpi A3 scans, rejects beyond the canvas limit', () => {
    expect(tiffDimsOk(10200, 13200)).toBe(true)
    expect(tiffDimsOk(7016, 9933)).toBe(true)
    expect(tiffDimsOk(16384, 9000)).toBe(true)
    expect(tiffDimsOk(16385, 1)).toBe(false)
    expect(tiffDimsOk(13000, 13000)).toBe(false)
    expect(tiffDimsOk(0, 10)).toBe(false)
    expect(tiffDimsOk(1.5, 10)).toBe(false)
  })

  it('falls through to a smaller in-budget page when the first IFD is over budget', () => {
    stubDom()
    const rgba = new Uint8Array(4 * 3 * 4).fill(128)
    const small = new Uint8Array(UTIF.encodeImage(rgba, 4, 3))
    expect(tiffToDataUrl(prependGiantIfd(small, 30000, 30000))).toBe('data:image/png;base64,AAAA')
    expect(maxCanvasDim).toBe(4)
  })
})
