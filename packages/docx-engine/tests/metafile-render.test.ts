/**
 * Replay-level regression tests for the vendored emf-converter, driven by real
 * metafile bytes from the POI corpus:
 *  - wrench.emf (61_VariousPictures): SETWINDOWEXTEX/SETVIEWPORTEXTEX mapping
 *    with a non-zero, negative-Y rclBounds origin — used to draw fully
 *    off-canvas (blank image).
 *  - ole-icon.wmf (91_drawing): OLE preview icon drawn with two
 *    META_DIBSTRETCHBLT records (AND mask + XOR color) — used to be dropped.
 * Node has no canvas, so OffscreenCanvas/ImageData/FileReader are stubbed with
 * a recording 2D context and the draw calls are asserted geometrically.
 */
import { readFileSync } from 'node:fs'
import { gzipSync } from 'node:zlib'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { convertEmfToDataUrl, convertWmfToDataUrl } from '../src/vendor/emf-converter/index.mjs'
import { isMetafileMime, metafileToDataUrl } from '../src/metafile'

interface Call {
  method: string
  args: unknown[]
}

const canvases: FakeOffscreenCanvas[] = []
let calls: Call[] = []

function makeRecordingCtx(canvas: FakeOffscreenCanvas) {
  const props: Record<string | symbol, unknown> = {
    canvas,
    globalCompositeOperation: 'source-over',
  }
  const fns = new Map<string, (...args: unknown[]) => unknown>()
  return new Proxy(props, {
    get(target, prop) {
      if (prop in target) return target[prop]
      const name = String(prop)
      let fn = fns.get(name)
      if (!fn) {
        fn = (...args: unknown[]) => {
          calls.push({ method: name, args })
          if (name === 'measureText') return { width: 10 }
          return undefined
        }
        fns.set(name, fn)
      }
      return fn
    },
    set(target, prop, value) {
      target[prop] = value
      if (prop === 'globalCompositeOperation') {
        calls.push({ method: 'set:globalCompositeOperation', args: [value] })
      }
      return true
    },
  })
}

class FakeOffscreenCanvas {
  width: number
  height: number
  private ctx: ReturnType<typeof makeRecordingCtx> | null = null
  constructor(width: number, height: number) {
    this.width = width
    this.height = height
    canvases.push(this)
  }
  getContext(type: string) {
    if (type !== '2d') return null
    return (this.ctx ??= makeRecordingCtx(this))
  }
  convertToBlob() {
    return Promise.resolve(
      new Blob([new Uint8Array([0x89, 0x50, 0x4e, 0x47])], { type: 'image/png' }),
    )
  }
}

class FakeImageData {
  data: Uint8ClampedArray
  width: number
  height: number
  constructor(data: Uint8ClampedArray, width: number, height: number) {
    this.data = data
    this.width = width
    this.height = height
  }
}

class FakeFileReader {
  result: string | null = null
  onload: (() => void) | null = null
  onerror: (() => void) | null = null
  error: Error | null = null
  readAsDataURL(blob: Blob) {
    void blob.arrayBuffer().then((buf) => {
      this.result = `data:${blob.type};base64,${Buffer.from(buf).toString('base64')}`
      this.onload?.()
    })
  }
}

const globals = globalThis as Record<string, unknown>
const saved: Record<string, unknown> = {}

beforeAll(() => {
  for (const [key, value] of Object.entries({
    OffscreenCanvas: FakeOffscreenCanvas,
    ImageData: FakeImageData,
    FileReader: FakeFileReader,
  })) {
    saved[key] = globals[key]
    globals[key] = value
  }
})

afterAll(() => {
  for (const [key, value] of Object.entries(saved)) {
    if (value === undefined) delete globals[key]
    else globals[key] = value
  }
})

function reset() {
  canvases.length = 0
  calls = []
}

function loadFixture(name: string): ArrayBuffer {
  const bytes = readFileSync(new URL(`./fixtures/${name}`, import.meta.url))
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
}

function pathPoints(): Array<{ x: number; y: number }> {
  return calls
    .filter((c) => c.method === 'moveTo' || c.method === 'lineTo')
    .map((c) => ({ x: c.args[0] as number, y: c.args[1] as number }))
}

describe('EMF window/viewport mapping (wrench.emf)', () => {
  it('draws the full figure inside the canvas at dpiScale', async () => {
    reset()
    const result = await convertEmfToDataUrl(loadFixture('wrench.emf'), { dpiScale: 2 })
    expect(result).toMatch(/^data:image\/png;base64,/)
    // bounds (300,-616)→(490,-501): logical 190×115, canvas 380×230
    expect(canvases[0]?.width).toBe(380)
    expect(canvases[0]?.height).toBe(230)
    const pts = pathPoints()
    expect(pts.length).toBeGreaterThan(100)
    for (const { x, y } of pts) {
      expect(x).toBeGreaterThanOrEqual(-1)
      expect(x).toBeLessThanOrEqual(381)
      expect(y).toBeGreaterThanOrEqual(-1)
      expect(y).toBeLessThanOrEqual(231)
    }
    const xs = pts.map((p) => p.x)
    const ys = pts.map((p) => p.y)
    // regression: pre-fix everything landed off-canvas (blank white image)
    expect(Math.max(...xs) - Math.min(...xs)).toBeGreaterThan(300)
    expect(Math.max(...ys) - Math.min(...ys)).toBeGreaterThan(180)
  })
})

describe('WMF DIB blts (ole-icon.wmf)', () => {
  it('renders the icon via two DIBSTRETCHBLT records with mask/xor composites', async () => {
    reset()
    const result = await convertWmfToDataUrl(loadFixture('ole-icon.wmf'), { dpiScale: 2 })
    expect(result).toMatch(/^data:image\/png;base64,/)
    // placeable bounds 90×50 at dpiScale 2
    expect(canvases[0]?.width).toBe(180)
    expect(canvases[0]?.height).toBe(100)
    const draws = calls.filter((c) => c.method === 'drawImage')
    expect(draws).toHaveLength(2)
    for (const draw of draws) {
      // 9-arg form: src rect crops the (double-height) icon DIB
      expect(draw.args).toHaveLength(9)
      const [, sx, sy, sw, sh, dx, dy, dw, dh] = draw.args as [unknown, ...number[]]
      expect([sx, sy, sw, sh]).toEqual([0, 0, 32, 32])
      // dest (29,0,32,32) logical → ×2 device
      expect(dx).toBeCloseTo(58)
      expect(dy).toBeCloseTo(0)
      expect(dw).toBeCloseTo(64)
      expect(dh).toBeCloseTo(64)
    }
    const gcos = calls
      .filter((c) => c.method === 'set:globalCompositeOperation')
      .map((c) => c.args[0])
    expect(gcos).toContain('multiply') // SRCAND mask
    expect(gcos).toContain('difference') // SRCINVERT color
  })

  it('skips a bitmap-less DIBSTRETCHBLT and keeps replaying (MS-WMF 2.3.1.3)', async () => {
    reset()
    // header + SETWINDOWORG(0,0) + SETWINDOWEXT(50,90) + bitmap-less
    // META_DIBSTRETCHBLT (RecordSize == (0x0B41 >> 8) + 3 = 14 words, one
    // reserved word in place of the DIB) + RECTANGLE + EOF
    const rec = (type: number, params: number[]) => {
      const bytes = new Uint8Array(6 + params.length * 2)
      const v = new DataView(bytes.buffer)
      v.setUint32(0, 3 + params.length, true)
      v.setUint16(4, type, true)
      params.forEach((p, i) => v.setInt16(6 + i * 2, p, true))
      return bytes
    }
    const records = [
      rec(0x020b, [0, 0]), // SETWINDOWORG (y, x)
      rec(0x020c, [50, 90]), // SETWINDOWEXT (cy, cx)
      rec(0x0b41, [0x20, 0xcc, 0, 32, 32, 0, 0, 32, 32, 0, 0]), // rop lo/hi, reserved, params
      rec(0x041b, [40, 80, 10, 10]), // RECTANGLE (b, r, t, l)
      rec(0, []), // EOF
    ]
    const body = records.reduce((n, r) => n + r.length, 0)
    const wmf = new Uint8Array(18 + body)
    const hv = new DataView(wmf.buffer)
    hv.setUint16(0, 1, true) // mtType
    hv.setUint16(2, 9, true) // mtHeaderSize (words)
    hv.setUint16(4, 0x0300, true) // mtVersion
    hv.setUint32(6, wmf.length / 2, true) // mtSize (words)
    let at = 18
    for (const r of records) {
      wmf.set(r, at)
      at += r.length
    }
    const result = await convertWmfToDataUrl(wmf.buffer, { dpiScale: 1 })
    expect(result).toMatch(/^data:image\/png;base64,/)
    // the bitmap-less blt draws nothing, and the record after it still renders
    expect(calls.filter((c) => c.method === 'drawImage')).toHaveLength(0)
    expect(calls.some((c) => c.method === 'strokeRect')).toBe(true)
  })

  it('derives bounds from SETWINDOWORG/EXT when the placeable header is missing', async () => {
    reset()
    const withHeader = new Uint8Array(loadFixture('ole-icon.wmf'))
    const stripped = withHeader.slice(22) // drop the 22-byte placeable header
    const result = await convertWmfToDataUrl(
      stripped.buffer.slice(stripped.byteOffset, stripped.byteOffset + stripped.byteLength),
      { dpiScale: 2 },
    )
    expect(result).toMatch(/^data:image\/png;base64,/)
    // pre-fix fallback was a fixed 800×600 guess (1600×1200 canvas)
    expect(canvases[0]?.width).toBe(180)
    expect(canvases[0]?.height).toBe(100)
  })
})

describe('EMR_ALPHABLEND (w-icon.emf)', () => {
  it('draws the OLE icon bitmap with its real alpha channel', async () => {
    reset()
    const result = await convertEmfToDataUrl(loadFixture('w-icon.emf'), { dpiScale: 2 })
    expect(result).toMatch(/^data:image\/png;base64,/)
    // regression: the record was unhandled — only the caption text rendered
    const draws = calls.filter((c) => c.method === 'drawImage')
    expect(draws).toHaveLength(1)
    expect(draws[0].args).toHaveLength(9)
    const [, sx, sy, sw, sh, , , dw, dh] = draws[0].args as [unknown, ...number[]]
    expect([sx, sy, sw, sh]).toEqual([0, 0, 32, 32])
    expect(dw).toBeGreaterThan(0)
    expect(dh).toBeGreaterThan(0)
    // AC_SRC_ALPHA source: the pixels around the icon stay transparent
    // (the shared decoder's zero-alpha-means-opaque heuristic must not apply)
    const put = calls.find((c) => c.method === 'putImageData')
    const img = put?.args[0] as { data: Uint8ClampedArray; width: number; height: number }
    expect(img.width).toBe(32)
    expect(img.height).toBe(32)
    let transparent = 0
    let opaque = 0
    for (let i = 3; i < img.data.length; i += 4) {
      if (img.data[i] === 0) transparent++
      else if (img.data[i] === 255) opaque++
    }
    expect(transparent).toBe(240)
    expect(opaque).toBe(784)
    // the caption text still replays after the blend
    const texts = calls.filter((c) => c.method === 'fillText').map((c) => c.args[0])
    expect(texts).toContain('Документ-в-докуме')
    expect(texts).toContain('нте')
  })
})

describe('gzipped metafiles (.emz/.wmz)', () => {
  it('accepts emz/wmz mimes', () => {
    for (const m of ['image/emz', 'image/x-emz', 'image/wmz', 'image/x-wmz']) {
      expect(isMetafileMime(m)).toBe(true)
    }
  })

  it('gunzips and converts a wmz payload', async () => {
    reset()
    const gz = gzipSync(Buffer.from(loadFixture('ole-icon.wmf')))
    const result = await metafileToDataUrl(new Uint8Array(gz), 'image/x-wmz')
    expect(result).toMatch(/^data:image\/png;base64,/)
    expect(canvases[0]?.width).toBe(180)
  })

  it('gunzips gzip-compressed bytes even under a plain emf/wmf mime', async () => {
    reset()
    const gz = gzipSync(Buffer.from(loadFixture('wrench.emf')))
    const result = await metafileToDataUrl(new Uint8Array(gz), 'image/x-emf')
    expect(result).toMatch(/^data:image\/png;base64,/)
    expect(canvases[0]?.width).toBe(380)
  })
})
