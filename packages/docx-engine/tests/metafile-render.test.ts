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
      if (prop === 'font') {
        calls.push({ method: 'set:font', args: [value] })
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

/**
 * Minimal EMF: header + EXTCREATEFONTINDIRECTW + SELECTOBJECT + EXTTEXTOUTW +
 * EOF, with GDI+-style non-zero OutPrecision/Quality/PitchAndFamily bytes in
 * the LOGFONTW (the facename-offset regression trigger).
 */
function buildEmfWithText(faceName: string, text: string): ArrayBuffer {
  const header = new Uint8Array(88)
  const hv = new DataView(header.buffer)
  hv.setUint32(0, 1, true) // EMR_HEADER
  hv.setUint32(4, 88, true)
  hv.setInt32(16, 200, true) // rclBounds right
  hv.setInt32(20, 100, true) // rclBounds bottom
  hv.setInt32(32, 5292, true) // rclFrame right (.01 mm, matches bounds)
  hv.setInt32(36, 2646, true)
  hv.setUint32(40, 0x464d4520, true) // ' EMF'
  hv.setUint32(44, 0x00010000, true)
  hv.setUint32(52, 5, true) // nRecords
  hv.setUint16(56, 2, true) // nHandles
  hv.setInt32(72, 1920, true) // szlDevice
  hv.setInt32(76, 1080, true)
  hv.setInt32(80, 508, true) // szlMillimeters
  hv.setInt32(84, 286, true)

  const font = new Uint8Array(332)
  const fv = new DataView(font.buffer)
  fv.setUint32(0, 82, true) // EMR_EXTCREATEFONTINDIRECTW
  fv.setUint32(4, 332, true)
  fv.setUint32(8, 1, true) // ihFont
  fv.setInt32(12, -15, true) // lfHeight
  fv.setInt32(28, 400, true) // lfWeight
  fv.setUint8(35, 0x80) // lfCharSet
  fv.setUint8(36, 4) // lfOutPrecision
  fv.setUint8(38, 5) // lfQuality
  fv.setUint8(39, 0x32) // lfPitchAndFamily
  for (let i = 0; i < faceName.length && i < 31; i++) {
    fv.setUint16(40 + i * 2, faceName.charCodeAt(i), true)
  }

  const select = new Uint8Array(12)
  const sv = new DataView(select.buffer)
  sv.setUint32(0, 37, true) // EMR_SELECTOBJECT
  sv.setUint32(4, 12, true)
  sv.setUint32(8, 1, true)

  const textRec = new Uint8Array(76 + text.length * 2)
  const tv = new DataView(textRec.buffer)
  tv.setUint32(0, 84, true) // EMR_EXTTEXTOUTW
  tv.setUint32(4, textRec.length, true)
  tv.setUint32(24, 1, true) // GM_COMPATIBLE
  tv.setInt32(36, 10, true) // reference x
  tv.setInt32(40, 20, true) // reference y
  tv.setUint32(44, text.length, true)
  tv.setUint32(48, 76, true) // offString
  for (let i = 0; i < text.length; i++) {
    tv.setUint16(76 + i * 2, text.charCodeAt(i), true)
  }

  const eof = new Uint8Array(20)
  const ev = new DataView(eof.buffer)
  ev.setUint32(0, 14, true) // EMR_EOF
  ev.setUint32(4, 20, true)

  const parts = [header, font, select, textRec, eof]
  const total = parts.reduce((n, p) => n + p.length, 0)
  hv.setUint32(48, total, true) // nBytes
  const emf = new Uint8Array(total)
  let at = 0
  for (const p of parts) {
    emf.set(p, at)
    at += p.length
  }
  return emf.buffer
}

function fontSets(): string[] {
  return calls.filter((c) => c.method === 'set:font').map((c) => c.args[0] as string)
}

describe('EMR_EXTCREATEFONTINDIRECTW facename', () => {
  it('reads the LOGFONTW FaceName at +32, past the precision/quality bytes', async () => {
    reset()
    const result = await convertEmfToDataUrl(buildEmfWithText('Segoe UI', 'Test'), { dpiScale: 2 })
    expect(result).toMatch(/^data:image\/png;base64,/)
    expect(calls.filter((c) => c.method === 'fillText').map((c) => c.args[0])).toContain('Test')
    // regression: read at +28 prefixed the family with the OutPrecision/
    // Quality bytes; the invalid CSS ident made the ctx.font assignment fail
    // silently in a real canvas, dropping the 30px size with it
    expect(fontSets()).toContain('30px "Segoe UI"')
    for (const font of fontSets()) {
      expect([...font].some((c) => c.charCodeAt(0) < 0x20 || c.charCodeAt(0) === 0x7f)).toBe(false)
    }
  })

  it('strips control chars from a corrupt facename instead of losing the size', async () => {
    reset()
    await convertEmfToDataUrl(buildEmfWithText('\u0004㈅Meiryo UI', 'Test'), { dpiScale: 2 })
    expect(fontSets()).toContain('30px "㈅Meiryo UI"')
  })

  it('maps localized facenames through the wrapper fontFamilyMap', async () => {
    reset()
    await metafileToDataUrl(new Uint8Array(buildEmfWithText('游ゴシック', 'Test')), 'image/x-emf')
    expect(fontSets()).toContain('30px "Yu Gothic"')
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
