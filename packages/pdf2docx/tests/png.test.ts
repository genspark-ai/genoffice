import { describe, expect, it } from 'vitest'
import { cropRgba, encodeRgbaPng, rotateRgbaQuarter } from '../src/extract'

function solidRgba(w: number, h: number, rgba: [number, number, number, number]): Uint8Array {
  const buf = new Uint8Array(w * h * 4)
  for (let i = 0; i < w * h; i++) buf.set(rgba, i * 4)
  return buf
}

/** deterministic noise: no two neighbours alike, more than 256 colours */
function noiseRgba(w: number, h: number, alpha: number | 'vary'): Uint8Array {
  const buf = new Uint8Array(w * h * 4)
  let seed = 12345
  for (let i = 0; i < w * h; i++) {
    for (let c = 0; c < 3; c++) {
      seed = (seed * 1103515245 + 12345) >>> 0
      buf[i * 4 + c] = seed >>> 24
    }
    buf[i * 4 + 3] = alpha === 'vary' ? i % 256 : alpha
  }
  return buf
}

const ihdr = (png: Uint8Array) => ({
  width: new DataView(png.buffer, png.byteOffset).getUint32(16),
  height: new DataView(png.buffer, png.byteOffset).getUint32(20),
  bitDepth: png[24],
  colorType: png[25],
})

async function decode(png: Uint8Array) {
  const { PDFDocument } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  return doc.embedPng(png) // parses IHDR/PLTE/tRNS/IDAT incl. zlib inflate
}

describe('encodeRgbaPng', () => {
  it('produces a well-formed PNG signature and IHDR', () => {
    const png = encodeRgbaPng(solidRgba(3, 2, [255, 0, 0, 255]), 3, 2)
    expect([...png.slice(0, 8)]).toEqual([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    expect(ihdr(png)).toMatchObject({ width: 3, height: 2 })
  })

  it('a two-colour stencil becomes a 1-bit indexed PNG with a tRNS entry', async () => {
    // black ink on transparent, the shape of a JBIG2 scan rendered as a mask
    const w = 64
    const h = 32
    const rgba = solidRgba(w, h, [0, 0, 0, 0])
    for (let x = 0; x < w; x += 2) rgba.set([0, 0, 0, 255], (5 * w + x) * 4)
    const png = encodeRgbaPng(rgba, w, h)
    expect(ihdr(png)).toMatchObject({ bitDepth: 1, colorType: 3 })
    expect(new TextDecoder('latin1').decode(png)).toContain('tRNS')
    expect(png.length).toBeLessThan(200)
    const image = await decode(png)
    expect(image.width).toBe(w)
    expect(image.height).toBe(h)
  })

  it('opaque greys use the greyscale type, greys with alpha the grey+alpha type', async () => {
    const w = 20
    const h = 20
    // 400 pixels cycling through all 256 greys: too many for a bit-packed
    // palette, so plain greyscale (no PLTE) is the cheaper exact encoding
    const grey = new Uint8Array(w * h * 4)
    for (let i = 0; i < w * h; i++) grey.set([i & 0xff, i & 0xff, i & 0xff, 255], i * 4)
    const pngGrey = encodeRgbaPng(grey, w, h)
    expect(ihdr(pngGrey).colorType).toBe(0)
    expect((await decode(pngGrey)).width).toBe(w)

    const greyAlpha = new Uint8Array(w * h * 4)
    for (let i = 0; i < w * h; i++)
      greyAlpha.set([i & 0xff, i & 0xff, i & 0xff, (i >> 1) & 0xff], i * 4) // 400 distinct pairs
    const pngGa = encodeRgbaPng(greyAlpha, w, h)
    expect(ihdr(pngGa).colorType).toBe(4)
    expect((await decode(pngGa)).height).toBe(h)
  })

  it('opaque colour noise is RGB, translucent colour noise stays RGBA', async () => {
    const w = 40
    const h = 30
    const pngRgb = encodeRgbaPng(noiseRgba(w, h, 255), w, h)
    expect(ihdr(pngRgb).colorType).toBe(2)
    expect((await decode(pngRgb)).width).toBe(w)
    const pngRgba = encodeRgbaPng(noiseRgba(w, h, 'vary'), w, h)
    expect(ihdr(pngRgba).colorType).toBe(6)
    expect((await decode(pngRgba)).height).toBe(h)
  })

  it('really deflates: a flat 200×120 image is far below its raw size', async () => {
    const w = 200
    const h = 120
    const png = encodeRgbaPng(solidRgba(w, h, [10, 20, 30, 255]), w, h)
    expect(png.length).toBeLessThan(1000)
    const image = await decode(png)
    expect(image.width).toBe(w)
    expect(image.height).toBe(h)
  })

  it('a 300-colour opaque image decodes at full size (palette overflow path)', async () => {
    const w = 30
    const h = 10
    const rgba = new Uint8Array(w * h * 4)
    for (let i = 0; i < w * h; i++) rgba.set([i & 0xff, (i >> 8) & 0xff, 0, 255], i * 4)
    const png = encodeRgbaPng(rgba, w, h)
    expect(ihdr(png).colorType).toBe(2)
    expect((await decode(png)).width).toBe(w)
  })

  it('rejects a size mismatch', () => {
    expect(() => encodeRgbaPng(new Uint8Array(10), 2, 2)).toThrow(/expected/)
  })
})

describe('rotateRgbaQuarter (P27)', () => {
  // 2×1 bitmap: left pixel 1, right pixel 2
  const px = { rgba: Uint8Array.from([1, 1, 1, 255, 2, 2, 2, 255]), width: 2, height: 1 }

  it('one clockwise turn puts the left pixel on top', () => {
    const out = rotateRgbaQuarter(px, 1)
    expect([out.width, out.height]).toEqual([1, 2])
    expect([out.rgba[0], out.rgba[4]]).toEqual([1, 2])
  })

  it('three turns (270°) puts the right pixel on top; two turns mirrors', () => {
    const out = rotateRgbaQuarter(px, 3)
    expect([out.width, out.height]).toEqual([1, 2])
    expect([out.rgba[0], out.rgba[4]]).toEqual([2, 1])
    const flipped = rotateRgbaQuarter(px, 2)
    expect([flipped.width, flipped.height]).toEqual([2, 1])
    expect([flipped.rgba[0], flipped.rgba[4]]).toEqual([2, 1])
  })

  it('zero turns returns the input untouched', () => {
    expect(rotateRgbaQuarter(px, 0)).toBe(px)
  })
})

describe('cropRgba (P34)', () => {
  const px = (w: number, h: number) => ({
    rgba: Uint8Array.from({ length: w * h * 4 }, (_, i) => (i >> 2) % 256),
    width: w,
    height: h,
  })

  it('crops a fractional window measured from the top-left', () => {
    const out = cropRgba(px(4, 4), { left: 0.25, top: 0.5, right: 0.25, bottom: 0 })
    expect(out).not.toBeNull()
    expect(out!.width).toBe(2)
    expect(out!.height).toBe(2)
    // row 2 (from top), col 1 → source pixel index 2*4+1 = 9
    expect(out!.rgba[0]).toBe(9)
    // second output row starts at source pixel 3*4+1 = 13
    expect(out!.rgba[8]).toBe(13)
  })

  it('returns null when the window collapses', () => {
    expect(cropRgba(px(4, 4), { left: 0.5, top: 0, right: 0.5, bottom: 0 })).toBeNull()
  })
})
