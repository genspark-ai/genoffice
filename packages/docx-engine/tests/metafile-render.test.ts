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
import { isMetafileMime, MAX_METAFILE_GUNZIP_BYTES, metafileToDataUrl } from '../src/metafile'

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
          if (name === 'createPattern') return { pattern: args[0] }
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
      if (prop === 'fillStyle') {
        calls.push({ method: 'set:fillStyle', args: [value] })
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
    // the caption text still replays after the blend (glyph by glyph along its Dx advances)
    const texts = calls
      .filter((c) => c.method === 'fillText')
      .map((c) => c.args[0])
      .join('')
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

describe('EMR_EXTTEXTOUTW alignment (ole-caption.emf)', () => {
  it('centers a TA_CENTER caption instead of right-aligning it', async () => {
    reset()
    await convertEmfToDataUrl(loadFixture('ole-caption.emf'), { dpiScale: 2 })
    const canvas = canvases[0]
    // glyphs are placed individually along the Dx advances; the run is centered on the
    // reference point instead of hanging off it to the left
    const glyphs = calls.filter((c) => c.method === 'fillText')
    expect(glyphs.map((c) => c.args[0]).join('')).toBe('simple.txt')
    const xs = glyphs.map((c) => c.args[1] as number)
    expect(Math.min(...xs)).toBeLessThan(canvas.width / 2)
    expect(Math.max(...xs)).toBeGreaterThan(canvas.width / 2)
    expect(canvas.getContext('2d')?.textAlign).toBe('left')
  })
})

// ── Synthetic EMF/EMF+ builders ─────────────────────────────────────

function emfHeader(nBytes: number, nRecords: number): Uint8Array {
  const header = new Uint8Array(88)
  const hv = new DataView(header.buffer)
  hv.setUint32(0, 1, true)
  hv.setUint32(4, 88, true)
  hv.setInt32(16, 200, true) // rclBounds right
  hv.setInt32(20, 100, true) // rclBounds bottom
  hv.setInt32(32, 5292, true) // rclFrame (.01 mm)
  hv.setInt32(36, 2646, true)
  hv.setUint32(40, 0x464d4520, true)
  hv.setUint32(44, 0x00010000, true)
  hv.setUint32(48, nBytes, true)
  hv.setUint32(52, nRecords, true)
  hv.setUint16(56, 2, true)
  hv.setInt32(72, 1920, true)
  hv.setInt32(76, 1080, true)
  hv.setInt32(80, 508, true)
  hv.setInt32(84, 286, true)
  return header
}

function emfEof(): Uint8Array {
  const eof = new Uint8Array(20)
  const ev = new DataView(eof.buffer)
  ev.setUint32(0, 14, true)
  ev.setUint32(4, 20, true)
  return eof
}

function emfRectangle(l: number, t: number, r: number, b: number): Uint8Array {
  const rec = new Uint8Array(24)
  const v = new DataView(rec.buffer)
  v.setUint32(0, 43, true) // EMR_RECTANGLE
  v.setUint32(4, 24, true)
  v.setInt32(8, l, true)
  v.setInt32(12, t, true)
  v.setInt32(16, r, true)
  v.setInt32(20, b, true)
  return rec
}

/** EMR_EXTTEXTOUTW with an optional Dx advance array (logical units). */
function emfTextOut(text: string, dx?: number[], options = 0): Uint8Array {
  const dxBytes = dx ? dx.length * 4 : 0
  const rec = new Uint8Array(76 + text.length * 2 + dxBytes)
  const tv = new DataView(rec.buffer)
  tv.setUint32(0, 84, true)
  tv.setUint32(4, rec.length, true)
  tv.setUint32(24, 1, true) // GM_COMPATIBLE
  tv.setInt32(36, 10, true) // reference x
  tv.setInt32(40, 20, true) // reference y
  tv.setUint32(44, text.length, true)
  tv.setUint32(48, 76, true) // offString
  tv.setUint32(52, options, true)
  if (dx) tv.setUint32(72, 76 + text.length * 2, true) // offDx
  for (let i = 0; i < text.length; i++) tv.setUint16(76 + i * 2, text.charCodeAt(i), true)
  dx?.forEach((d, i) => tv.setInt32(76 + text.length * 2 + i * 4, d, true))
  return rec
}

/** EMR_CREATEPEN (ihPen 1) + EMR_SELECTOBJECT + EMR_MOVETOEX + EMR_LINETO: one stroked line. */
function emfPenLine(widthUnits: number): Uint8Array[] {
  const pen = new Uint8Array(28)
  const pv = new DataView(pen.buffer)
  pv.setUint32(0, 38, true)
  pv.setUint32(4, 28, true)
  pv.setUint32(8, 1, true) // ihPen
  pv.setUint32(12, 0, true) // PS_SOLID
  pv.setInt32(16, widthUnits, true)
  pv.setUint32(24, 0x000000, true)
  const sel = new Uint8Array(12)
  const sv = new DataView(sel.buffer)
  sv.setUint32(0, 37, true)
  sv.setUint32(4, 12, true)
  sv.setUint32(8, 1, true)
  const move = new Uint8Array(16)
  const mv = new DataView(move.buffer)
  mv.setUint32(0, 27, true)
  mv.setUint32(4, 16, true)
  mv.setInt32(8, 10, true)
  mv.setInt32(12, 30, true)
  const line = new Uint8Array(16)
  const lv = new DataView(line.buffer)
  lv.setUint32(0, 54, true)
  lv.setUint32(4, 16, true)
  lv.setInt32(8, 150, true)
  lv.setInt32(12, 30, true)
  return [pen, sel, move, line]
}

function emfSetTextAlign(value: number): Uint8Array {
  const rec = new Uint8Array(12)
  const v = new DataView(rec.buffer)
  v.setUint32(0, 22, true) // EMR_SETTEXTALIGN
  v.setUint32(4, 12, true)
  v.setUint32(8, value, true)
  return rec
}

/** One EMF+ record (12-byte header + payload). */
function plusRecord(type: number, flags: number, payload: Uint8Array): Uint8Array {
  const rec = new Uint8Array(12 + payload.length)
  const v = new DataView(rec.buffer)
  v.setUint16(0, type, true)
  v.setUint16(2, flags, true)
  v.setUint32(4, rec.length, true)
  v.setUint32(8, payload.length, true)
  rec.set(payload, 12)
  return rec
}

/** EMR_COMMENT carrying EMF+ records. */
function emfPlusComment(...records: Uint8Array[]): Uint8Array {
  const body = records.reduce((n, r) => n + r.length, 0)
  const size = 16 + body
  const padded = size + ((4 - (size % 4)) % 4)
  const rec = new Uint8Array(padded)
  const v = new DataView(rec.buffer)
  v.setUint32(0, 70, true) // EMR_COMMENT
  v.setUint32(4, padded, true)
  v.setUint32(8, 4 + body, true) // DataSize
  v.setUint32(12, 0x2b464d45, true) // 'EMF+'
  let at = 16
  for (const r of records) {
    rec.set(r, at)
    at += r.length
  }
  return rec
}

function plusHeader(dual: boolean): Uint8Array {
  const payload = new Uint8Array(16)
  const v = new DataView(payload.buffer)
  v.setUint32(0, 0xdbc01002, true)
  v.setUint32(4, dual ? 1 : 0, true)
  v.setUint32(8, 96, true)
  v.setUint32(12, 96, true)
  return plusRecord(0x4001, dual ? 1 : 0, payload)
}

function plusGetDC(): Uint8Array {
  return plusRecord(0x4004, 0, new Uint8Array(0))
}

function plusFontObject(id: number, family: string, emSize: number): Uint8Array {
  const payload = new Uint8Array(24 + family.length * 2)
  const v = new DataView(payload.buffer)
  v.setUint32(0, 0xdbc01002, true)
  v.setFloat32(4, emSize, true)
  v.setUint32(8, 2, true) // UnitTypePixel
  v.setUint32(12, 0, true)
  v.setUint32(20, family.length, true)
  for (let i = 0; i < family.length; i++) v.setUint16(24 + i * 2, family.charCodeAt(i), true)
  return plusRecord(0x4008, (6 << 8) | id, payload) // ObjectTypeFont
}

function plusDrawDriverString(fontId: number, text: string, pos: number[][]): Uint8Array {
  const payload = new Uint8Array(16 + text.length * 2 + pos.length * 8 + 24)
  const v = new DataView(payload.buffer)
  v.setUint32(0, 0xff000000, true) // brush: opaque black ARGB
  v.setUint32(4, 1, true) // DriverStringOptionsCmapLookup
  v.setUint32(8, 1, true) // MatrixPresent
  v.setUint32(12, text.length, true)
  let at = 16
  for (let i = 0; i < text.length; i++, at += 2) v.setUint16(at, text.charCodeAt(i), true)
  for (const [x, y] of pos) {
    v.setFloat32(at, x, true)
    v.setFloat32(at + 4, y, true)
    at += 8
  }
  for (const m of [1, 0, 0, 1, 0, 0]) {
    v.setFloat32(at, m, true)
    at += 4
  }
  return plusRecord(0x4036, 0x8000 | fontId, payload)
}

function assembleEmf(records: Uint8Array[]): ArrayBuffer {
  const eof = emfEof()
  const body = records.reduce((n, r) => n + r.length, 0) + 88 + eof.length
  const parts = [emfHeader(body, records.length + 2), ...records, eof]
  const emf = new Uint8Array(body)
  let at = 0
  for (const p of parts) {
    emf.set(p, at)
    at += p.length
  }
  return emf.buffer
}

describe('GDI pen width', () => {
  it('scales the logical pen width to device pixels (1 unit → 2 px at dpiScale 2)', async () => {
    reset()
    await convertEmfToDataUrl(assembleEmf(emfPenLine(1)), { dpiScale: 2 })
    const strokeIdx = calls.findIndex((c) => c.method === 'stroke')
    expect(strokeIdx).toBeGreaterThan(0)
    const page = canvases.find((c) => c.width === 400)
    expect(page?.getContext('2d')?.lineWidth).toBe(2)
  })

  it('keeps a hairline (width 0) at one device pixel', async () => {
    reset()
    await convertEmfToDataUrl(assembleEmf(emfPenLine(0)), { dpiScale: 2 })
    const page = canvases.find((c) => c.width === 400)
    expect(page?.getContext('2d')?.lineWidth).toBe(1)
  })
})

describe('EMR_EXTTEXTOUTW Dx advances', () => {
  it("places every glyph by the writer's advances, not the substitute font's widths", async () => {
    reset()
    await convertEmfToDataUrl(assembleEmf([emfTextOut('ab', [30, 40])]), { dpiScale: 2 })
    const glyphs = calls.filter((c) => c.method === 'fillText')
    expect(glyphs.map((c) => c.args[0])).toEqual(['a', 'b'])
    // bounds 200 wide → canvas 400 → 2 px per logical unit; 'b' sits 30 units after 'a'
    const [ax, bx] = glyphs.map((c) => c.args[1] as number)
    expect(bx - ax).toBeCloseTo(60, 3)
  })

  it('ETO_PDY interleaves y offsets: glyphs step vertically too', async () => {
    reset()
    await convertEmfToDataUrl(assembleEmf([emfTextOut('ab', [30, 5, 40, -5], 0x2000)]), {
      dpiScale: 2,
    })
    const glyphs = calls.filter((c) => c.method === 'fillText')
    expect(glyphs.map((c) => c.args[0])).toEqual(['a', 'b'])
    expect((glyphs[1]!.args[1] as number) - (glyphs[0]!.args[1] as number)).toBeCloseTo(60, 3)
    expect((glyphs[1]!.args[2] as number) - (glyphs[0]!.args[2] as number)).toBeCloseTo(10, 3)
  })

  it('TA_UPDATECP advances the current point by the text extent even without Dx', async () => {
    reset()
    await convertEmfToDataUrl(
      assembleEmf([emfSetTextAlign(1), emfTextOut('ab'), emfTextOut('cd')]),
      { dpiScale: 2 },
    )
    const runs = calls.filter((c) => c.method === 'fillText')
    expect(runs.map((c) => c.args[0])).toEqual(['ab', 'cd'])
    // the fake measureText reports 10 px: the second run starts one extent further right
    expect((runs[1]!.args[1] as number) - (runs[0]!.args[1] as number)).toBeCloseTo(10, 3)
  })

  it('keeps the whole-string draw when no Dx array is present', async () => {
    reset()
    await convertEmfToDataUrl(assembleEmf([emfTextOut('ab')]), { dpiScale: 2 })
    expect(calls.filter((c) => c.method === 'fillText').map((c) => c.args[0])).toEqual(['ab'])
  })
})

describe('EMF+ dual mode', () => {
  it('skips GDI drawing outside an EmfPlusGetDC window and paints it inside one', async () => {
    reset()
    await convertEmfToDataUrl(
      assembleEmf([
        emfPlusComment(plusHeader(true)),
        emfRectangle(0, 0, 50, 50), // GDI twin of an EMF+ fill: must not paint
        emfPlusComment(plusGetDC()),
        emfRectangle(100, 0, 150, 50), // GDI-only content after GetDC: paints
      ]),
      { dpiScale: 2 },
    )
    const rects = calls.filter((c) => /^(rect|fillRect|strokeRect)$/.test(c.method))
    expect(rects.length).toBeGreaterThan(0)
    expect(rects.every((c) => (c.args[0] as number) >= 200)).toBe(true)
  })

  it('a non-dual EMF+ file still replays its GDI records', async () => {
    reset()
    await convertEmfToDataUrl(
      assembleEmf([emfPlusComment(plusHeader(false)), emfRectangle(0, 0, 50, 50)]),
      { dpiScale: 2 },
    )
    expect(calls.some((c) => /^(rect|fillRect|strokeRect)$/.test(c.method))).toBe(true)
  })
})

/** EMF+ image object (compressed PNG bytes) split into continuation records, each chunk
 *  prefixed by the 4-byte TotalObjectSize like GDI+ writes them. */
function plusImageObjectChunks(
  id: number,
  png: Uint8Array,
  chunkSize: number,
  bareLastChunk = false,
): Uint8Array[] {
  const body = new Uint8Array(28 + png.length)
  const v = new DataView(body.buffer)
  v.setUint32(0, 0xdbc01002, true)
  v.setUint32(4, 1, true) // ImageDataTypeBitmap
  v.setUint32(24, 1, true) // BitmapDataTypeCompressed
  body.set(png, 28)
  const out: Uint8Array[] = []
  for (let at = 0; at < body.length; at += chunkSize) {
    const slice = body.subarray(at, Math.min(at + chunkSize, body.length))
    const last = at + chunkSize >= body.length
    if (bareLastChunk && last) {
      // spec-literal final chunk: continue flag clear, no size prefix
      out.push(plusRecord(0x4008, (5 << 8) | id, slice))
      continue
    }
    const payload = new Uint8Array(4 + slice.length)
    new DataView(payload.buffer).setUint32(0, body.length, true)
    payload.set(slice, 4)
    out.push(plusRecord(0x4008, 0x8000 | (5 << 8) | id, payload))
  }
  return out
}

function plusDrawImagePoints(id: number): Uint8Array {
  const payload = new Uint8Array(28 + 24)
  const v = new DataView(payload.buffer)
  v.setUint32(24, 3, true) // point count
  const pts = [0, 0, 100, 0, 0, 50]
  pts.forEach((p, i) => v.setFloat32(28 + i * 4, p, true))
  return plusRecord(0x401b, id, payload)
}

describe('EMF+ continued image objects', () => {
  it('a final chunk without the size prefix (continue flag clear) still completes the object', async () => {
    reset()
    const created: Blob[] = []
    const savedCIB = globals.createImageBitmap
    globals.createImageBitmap = async (blob: Blob) => {
      created.push(blob)
      return { width: 10, height: 5, close() {} }
    }
    try {
      const png = new Uint8Array(150)
      for (let i = 0; i < png.length; i++) png[i] = i & 0xff
      const chunks = plusImageObjectChunks(3, png, 64, true)
      await convertEmfToDataUrl(
        assembleEmf([emfPlusComment(plusHeader(false), ...chunks, plusDrawImagePoints(3))]),
        { dpiScale: 1 },
      )
      expect(created).toHaveLength(1)
      const bytes = new Uint8Array(await created[0]!.arrayBuffer())
      expect(bytes.length).toBe(150)
      expect(bytes[149]).toBe(149)
    } finally {
      if (savedCIB === undefined) delete globals.createImageBitmap
      else globals.createImageBitmap = savedCIB
    }
  })

  it('reassembles chunks that each repeat the size prefix, across comment records', async () => {
    reset()
    const created: Blob[] = []
    const savedCIB = globals.createImageBitmap
    globals.createImageBitmap = async (blob: Blob) => {
      created.push(blob)
      return { width: 10, height: 5, close() {} }
    }
    try {
      const png = new Uint8Array(150)
      png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
      for (let i = 8; i < png.length; i++) png[i] = i & 0xff
      const chunks = plusImageObjectChunks(3, png, 64)
      expect(chunks.length).toBe(3)
      await convertEmfToDataUrl(
        assembleEmf([
          emfPlusComment(plusHeader(false), chunks[0]!),
          emfPlusComment(chunks[1]!),
          emfPlusComment(chunks[2]!, plusDrawImagePoints(3)),
        ]),
        { dpiScale: 1 },
      )
      // the decoded bitmap is the exact PNG payload: byte-identical after reassembly
      expect(created).toHaveLength(1)
      const bytes = new Uint8Array(await created[0]!.arrayBuffer())
      expect(bytes.length).toBe(150)
      expect(Array.from(bytes.slice(0, 8))).toEqual([
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
      ])
      expect(bytes[100]).toBe(100)
      expect(bytes[149]).toBe(149)
      // painted in place (pre-decoded pass), not deferred past the replay
      expect(calls.some((c) => c.method === 'drawImage')).toBe(true)
    } finally {
      if (savedCIB === undefined) delete globals.createImageBitmap
      else globals.createImageBitmap = savedCIB
    }
  })
})

describe('EmfPlusDrawDriverString glyph positions', () => {
  it('reads an odd glyph count without padding and places each glyph', async () => {
    reset()
    await convertEmfToDataUrl(
      assembleEmf([
        emfPlusComment(
          plusHeader(false),
          plusFontObject(0, 'Arial', 12),
          plusDrawDriverString(0, 'abc', [
            [5, 10],
            [15, 10],
            [25, 10],
          ]),
        ),
      ]),
      { dpiScale: 2 },
    )
    const glyphs = calls.filter((c) => c.method === 'fillText')
    expect(glyphs.map((c) => c.args.slice(0, 3))).toEqual([
      ['a', 5, 10],
      ['b', 15, 10],
      ['c', 25, 10],
    ])
  })
})

describe('EMR_EXTCREATEFONTINDIRECTW facename', () => {
  it('reads the LOGFONTW FaceName at +32, past the precision/quality bytes', async () => {
    reset()
    const result = await convertEmfToDataUrl(buildEmfWithText('Segoe UI', 'Test'), { dpiScale: 2 })
    expect(result).toMatch(/^data:image\/png;base64,/)
    expect(calls.filter((c) => c.method === 'fillText').map((c) => c.args[0])).toContain('Test')
    // regression: read at +28 prefixed the family with the OutPrecision/
    // Quality bytes; the invalid CSS ident made the ctx.font assignment fail
    // silently in a real canvas, dropping the 30px size with it
    expect(fontSets()).toContain('30px "Segoe UI", sans-serif')
    for (const font of fontSets()) {
      expect([...font].some((c) => c.charCodeAt(0) < 0x20 || c.charCodeAt(0) === 0x7f)).toBe(false)
    }
  })

  it('strips control chars from a corrupt facename instead of losing the size', async () => {
    reset()
    await convertEmfToDataUrl(buildEmfWithText('\u0004㈅Meiryo UI', 'Test'), { dpiScale: 2 })
    expect(fontSets()).toContain('30px "㈅Meiryo UI", sans-serif')
  })

  it('maps localized facenames through the wrapper fontFamilyMap', async () => {
    reset()
    await metafileToDataUrl(new Uint8Array(buildEmfWithText('游ゴシック', 'Test')), 'image/x-emf')
    expect(fontSets()).toContain('30px "Yu Gothic", sans-serif')
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

  it('refuses gzip bombs instead of exhausting memory', async () => {
    reset()
    const bomb = gzipSync(Buffer.alloc(MAX_METAFILE_GUNZIP_BYTES + 1))
    await expect(metafileToDataUrl(new Uint8Array(bomb), 'image/x-emz')).resolves.toBeNull()
  })
})

describe('EMF+ bitmap image object (synthetic)', () => {
  // EMR_HEADER + one EMR_COMMENT(EMF+) holding EmfPlusHeader, EmfPlusObject
  // (image → bitmap → BitmapDataType Pixel = 0, 4×2 32bppARGB), EmfPlusDrawImage,
  // EmfPlusEndOfFile + EMR_EOF — the shape PowerPoint uses for a picture
  // background exported as EMF+ (prod master bg drew a blank white page)
  function buildEmfPlusBitmap(): ArrayBuffer {
    const w = 4
    const h = 2
    const pixels = new Uint8Array(w * h * 4).fill(0x80)
    const plus: number[] = []
    const u16 = (v: number) => plus.push(v & 0xff, (v >> 8) & 0xff)
    const u32s = (...vs: number[]) => {
      for (const v of vs) plus.push(v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >>> 24) & 0xff)
    }
    const f32s = (...vs: number[]) => plus.push(...new Uint8Array(new Float32Array(vs).buffer))
    // EmfPlusHeader
    u16(0x4001)
    u16(1)
    u32s(28, 16, 0xdbc01001, 0, 96, 96)
    // EmfPlusObject: type Image (5) << 8 | id 1
    const objData = 28 + pixels.length
    u16(0x4008)
    u16(0x0500 | 1)
    u32s(12 + objData, objData, 0xdbc01002, 1, w, h, w * 4, 0x26200a, 0)
    plus.push(...pixels)
    // EmfPlusDrawImage: image id 1, src unit pixel, float RectF src/dest
    u16(0x401a)
    u16(1)
    u32s(52, 40, 0, 2)
    f32s(0, 0, w, h, 0, 0, w, h)
    // EmfPlusEndOfFile
    u16(0x4002)
    u16(0)
    u32s(12, 0)

    const commentData = 4 + plus.length
    const commentSize = 12 + commentData
    const total = 108 + commentSize + 20
    const header: number[] = [
      1,
      108, // EMR_HEADER
      0,
      0,
      w,
      h, // bounds (px)
      0,
      0,
      Math.round((w * 2540) / 96),
      Math.round((h * 2540) / 96), // frame (0.01 mm)
      0x464d4520,
      0x10000,
      total,
      3, // signature, version, bytes, records
      0, // handles + reserved
      0,
      0,
      0, // description, palette
      1920,
      1080,
      508,
      286, // device px / mm
      0,
      0,
      0,
      508000,
      286000, // pixel format, OpenGL, micrometers
      70,
      commentSize,
      commentData,
      0x2b464d45, // EMR_COMMENT + 'EMF+'
    ]
    const buf = new ArrayBuffer(total)
    const v = new DataView(buf)
    header.forEach((x, i) => v.setUint32(i * 4, x >>> 0, true))
    new Uint8Array(buf).set(plus, header.length * 4)
    const eof = [14, 20, 0, 16, 20]
    eof.forEach((x, i) => v.setUint32(total - 20 + i * 4, x, true))
    return buf
  }

  it('decodes the pixel bitmap and draws it through DrawImage', async () => {
    reset()
    const created: Blob[] = []
    const savedCIB = globals.createImageBitmap
    globals.createImageBitmap = async (blob: Blob) => {
      created.push(blob)
      return { width: 4, height: 2, close() {} }
    }
    try {
      const result = await convertEmfToDataUrl(buildEmfPlusBitmap(), { dpiScale: 1 })
      expect(result).toMatch(/^data:image\/png;base64,/)
      expect(canvases[0]?.width).toBe(4)
      expect(canvases[0]?.height).toBe(2)
      // regression: BitmapDataType was compared against 1/2 instead of 0/1, so the
      // object never decoded and DrawImage had nothing to paint
      expect(created).toHaveLength(1)
      const bmp = new Uint8Array(await created[0].arrayBuffer())
      expect(String.fromCharCode(bmp[0], bmp[1])).toBe('BM')
      const draws = calls.filter((c) => c.method === 'drawImage')
      expect(draws).toHaveLength(1)
      expect(draws[0].args.slice(1)).toEqual([0, 0, 4, 2])
    } finally {
      if (savedCIB === undefined) delete globals.createImageBitmap
      else globals.createImageBitmap = savedCIB
    }
  })
})

describe('EMR_CREATEDIBPATTERNBRUSHPT (synthetic)', () => {
  // EMR_HEADER + CREATEDIBPATTERNBRUSHPT (8×8 1bpp checker DIB) + SELECTOBJECT + BITBLT
  // PATCOPY + EOF — how Excel OLE previews draw dotted cell borders
  function buildPatternBrushEmf(): ArrayBuffer {
    const dib = [
      40,
      8,
      8,
      0x00010001,
      0,
      32,
      0,
      0,
      2,
      0, // BITMAPINFOHEADER (planes=1, bpp=1)
      0x00000000,
      0x00ffffff, // color table: black, white
      0xaa55aa55,
      0xaa55aa55,
      0xaa55aa55,
      0xaa55aa55,
      0xaa55aa55,
      0xaa55aa55,
      0xaa55aa55,
      0xaa55aa55, // 8 rows × 4 bytes
    ]
    const brushRec = [94, 32 + dib.length * 4, 1, 0, 32, 48, 80, 32, ...dib]
    const select = [37, 12, 1]
    // BITBLT: bounds, dest 10,10 40×2, PATCOPY, no source DIB
    const bitblt = [
      76, 100, 0, 0, 100, 100, 10, 10, 40, 2, 0xf00021, 0, 0, 0, 0, 0, 0, 0x1000000, 0, 0, 0, 0, 0,
      0, 0,
    ]
    const body = [...brushRec, ...select, ...bitblt, 14, 20, 0, 16, 20]
    const total = 108 + body.length * 4
    const header = [
      1,
      108,
      0,
      0,
      100,
      100,
      0,
      0,
      2646,
      2646,
      0x464d4520,
      0x10000,
      total,
      5,
      2,
      0,
      0,
      0,
      1920,
      1080,
      508,
      286,
      0,
      0,
      0,
      508000,
      286000,
    ]
    const buf = new ArrayBuffer(total)
    const v = new DataView(buf)
    ;[...header, ...body].forEach((x, i) => v.setUint32(i * 4, x >>> 0, true))
    return buf
  }

  it('fills the PATCOPY blit with a repeating pattern built from the brush DIB', async () => {
    reset()
    const result = await convertEmfToDataUrl(buildPatternBrushEmf(), { dpiScale: 1 })
    expect(result).toMatch(/^data:image\/png;base64,/)
    // the 8×8 pattern tile is a second canvas next to the 100×100 page
    const tile = canvases.find((c) => c.width === 8 && c.height === 8)
    expect(tile).toBeDefined()
    const patterns = calls.filter((c) => c.method === 'createPattern')
    expect(patterns).toHaveLength(1)
    expect(patterns[0].args[1]).toBe('repeat')
    // regression: the record was unhandled, so the blit used the previous solid brush
    const fillIdx = calls.findIndex((c) => c.method === 'fillRect')
    expect(fillIdx).toBeGreaterThan(0)
    const lastFillStyle = calls
      .slice(0, fillIdx)
      .reverse()
      .find((c) => c.method === 'set:fillStyle')
    expect(lastFillStyle?.args[0]).toEqual({ pattern: tile })
    expect(calls[fillIdx].args).toEqual([10, 10, 40, 2])
  })

  it('tiles the brush in device pixels: the 8x8 DIB becomes a 16x16 tile at dpiScale 2', async () => {
    reset()
    await convertEmfToDataUrl(buildPatternBrushEmf(), { dpiScale: 2 })
    const patterns = calls.filter((c) => c.method === 'createPattern')
    expect(patterns).toHaveLength(1)
    const tile = patterns[0].args[0] as FakeOffscreenCanvas
    expect([tile.width, tile.height]).toEqual([16, 16])
    // nearest-neighbour blow-up of the 8x8 source, not a smoothed resample
    const blowUp = calls.find(
      (c) => c.method === 'drawImage' && c.args[3] === 16 && c.args[4] === 16,
    )
    expect(blowUp).toBeDefined()
    const bigCtx = canvases.find((c) => c.width === 16 && c.height === 16)
    expect(bigCtx?.getContext('2d')?.imageSmoothingEnabled).toBe(false)
  })
})

/**
 * Minimal EMF+ dual-mode file the way Office writes chart pictures: the plot is
 * clipped with SetClipRect, then SetClipRegion(region object) widens the clip
 * back to the whole chart before the title/legend are drawn. Region objects
 * carry ObjectType 4 (MS-EMFPLUS 2.1.1.22) and a single leaf node has
 * RegionNodeCount 0.
 */
function buildEmfPlusRegionClipFile(): ArrayBuffer {
  const plusRecord = (type: number, flags: number, data: number[]): number[] => {
    const size = 12 + data.length
    return [...u16(type), ...u16(flags), ...u32(size), ...u32(data.length), ...data]
  }
  const plus = [
    ...plusRecord(0x4001, 1, [...u32(0xdbc01002), ...u32(1), ...u32(96), ...u32(96)]),
    ...plusRecord(0x4032, 1 << 8, [...f32(10), ...f32(10), ...f32(20), ...f32(20)]),
    ...plusRecord(0x4008, (4 << 8) | 0, [
      ...u32(0xdbc01002),
      ...u32(0),
      ...u32(0x10000000),
      ...f32(0),
      ...f32(0),
      ...f32(100),
      ...f32(60),
    ]),
    ...plusRecord(0x4034, 1 << 8, []),
    ...plusRecord(0x400a, 0x8000, [
      ...u32(0xff0000ff),
      ...u32(1),
      ...f32(0),
      ...f32(0),
      ...f32(100),
      ...f32(60),
    ]),
    ...plusRecord(0x4002, 0, []),
  ]
  const commentData = [...u32(0x2b464d45), ...plus]
  const comment = [
    ...u32(70),
    ...u32(12 + commentData.length),
    ...u32(commentData.length),
    ...commentData,
  ]
  const eof = [...u32(14), ...u32(20), ...u32(0), ...u32(16), ...u32(20)]
  const total = 88 + comment.length + eof.length
  const header = [
    ...u32(1),
    ...u32(88),
    ...i32(0),
    ...i32(0),
    ...i32(100),
    ...i32(60),
    ...i32(0),
    ...i32(0),
    ...i32(2646),
    ...i32(1588),
    ...u32(0x464d4520),
    ...u32(0x10000),
    ...u32(total),
    ...u32(3),
    ...u16(1),
    ...u16(0),
    ...u32(0),
    ...u32(0),
    ...u32(0),
    ...i32(1000),
    ...i32(600),
    ...i32(265),
    ...i32(159),
  ]
  return new Uint8Array([...header, ...comment, ...eof]).buffer
}
function u16(v: number): number[] {
  return [v & 0xff, (v >>> 8) & 0xff]
}
function u32(v: number): number[] {
  return [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff]
}
function i32(v: number): number[] {
  return u32(v >>> 0)
}
function f32(v: number): number[] {
  const b = new Uint8Array(4)
  new DataView(b.buffer).setFloat32(0, v, true)
  return [...b]
}

describe('EMF+ SetClipRegion (synthetic dual-mode chart picture)', () => {
  it('restores the whole-chart clip from a region object before the fill', async () => {
    reset()
    const url = await convertEmfToDataUrl(buildEmfPlusRegionClipFile(), { dpiScale: 1 })
    expect(url).toMatch(/^data:image\/png/)
    const fillAt = calls.findIndex((c) => c.method === 'fillRect')
    expect(fillAt).toBeGreaterThan(0)
    const clipAt = calls
      .slice(0, fillAt)
      .map((c) => c.method)
      .lastIndexOf('clip')
    expect(clipAt).toBeGreaterThan(0)
    const pathStart = calls
      .slice(0, clipAt)
      .map((c) => c.method)
      .lastIndexOf('beginPath')
    const xs: number[] = []
    const ys: number[] = []
    for (const c of calls.slice(pathStart, clipAt)) {
      if (c.method === 'rect') {
        const [x, y, w, h] = c.args as number[]
        xs.push(x, x + w)
        ys.push(y, y + h)
      } else if (c.method === 'moveTo' || c.method === 'lineTo') {
        xs.push(c.args[0] as number)
        ys.push(c.args[1] as number)
      }
    }
    // the 20x20 plot clip must have been replaced by the 100x60 region
    expect(Math.max(...xs) - Math.min(...xs)).toBeGreaterThan(90)
    expect(Math.max(...ys) - Math.min(...ys)).toBeGreaterThan(50)
  })
})

describe('GDI TA_TOP text uses the Windows ascent of known faces', () => {
  it('Yu Gothic text sits 1.29 em below the top (canvas hhea metrics would put it at 0.88 em)', async () => {
    reset()
    await convertEmfToDataUrl(buildEmfWithText('\u6e38\u30b4\u30b7\u30c3\u30af', 'ab'), {
      dpiScale: 2,
    })
    const glyphs = calls.filter((c) => c.method === 'fillText')
    expect(glyphs.length).toBeGreaterThan(0)
    // reference y 20 logical → 40 px; lfHeight -15 → 30 px em; baseline = 40 + 1.292 × 30
    expect(glyphs[0].args[2] as number).toBeCloseTo(40 + 1.292 * 30, 1)
  })

  it('localized MS face names with fullwidth Latin match the table (MS PGothic 0.859 em)', async () => {
    reset()
    await convertEmfToDataUrl(
      buildEmfWithText('\uff2d\uff33 \uff30\u30b4\u30b7\u30c3\u30af', 'ab'),
      {
        dpiScale: 2,
      },
    )
    const glyphs = calls.filter((c) => c.method === 'fillText')
    expect(glyphs[0].args[2] as number).toBeCloseTo(40 + 0.859 * 30, 1)
  })

  it('unknown faces keep the canvas top baseline', async () => {
    reset()
    await convertEmfToDataUrl(buildEmfWithText('Some Unknown Face', 'ab'), { dpiScale: 2 })
    const glyphs = calls.filter((c) => c.method === 'fillText')
    expect(glyphs[0].args[2]).toBe(40)
  })
})
