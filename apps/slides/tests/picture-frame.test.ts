import { describe, expect, it } from 'vitest'
import { imageDpiFromBytes } from '@genoffice/pptx-render'
import { pictureFrame } from '../src/main/picture-frame'

const PT = 12700
const SLIDE = { width: 960 * PT, height: 540 * PT }
const MAC_DPI = 144

// ── byte builders ────────────────────────────────────────────────────────

function be32(v: number): number[] {
  return [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255]
}

function be16(v: number): number[] {
  return [(v >> 8) & 255, v & 255]
}

function ascii(s: string): number[] {
  return [...s].map((c) => c.charCodeAt(0))
}

function png(phys?: { ppmX: number; ppmY: number; unit: number }): Uint8Array {
  const chunk = (type: string, data: number[]) => [
    ...be32(data.length),
    ...ascii(type),
    ...data,
    0,
    0,
    0,
    0,
  ]
  return Uint8Array.from([
    0x89,
    0x50,
    0x4e,
    0x47,
    0x0d,
    0x0a,
    0x1a,
    0x0a,
    ...chunk('IHDR', [...be32(2), ...be32(2), 8, 6, 0, 0, 0]),
    ...(phys ? chunk('pHYs', [...be32(phys.ppmX), ...be32(phys.ppmY), phys.unit]) : []),
    ...chunk('IDAT', [0, 0]),
    ...chunk('IEND', []),
  ])
}

function segment(marker: number, payload: number[]): number[] {
  return [0xff, marker, ...be16(payload.length + 2), ...payload]
}

function jfif(units: number, xd: number, yd: number): number[] {
  return segment(0xe0, [...ascii('JFIF'), 0, 1, 1, units, ...be16(xd), ...be16(yd), 0, 0])
}

/** APP1 Exif with a little-endian IFD0 carrying X/YResolution (RATIONAL) and ResolutionUnit. */
function exif(xRes: [number, number], yRes: [number, number], unit: number): number[] {
  const le16 = (v: number) => [v & 255, (v >> 8) & 255]
  const le32 = (v: number) => [v & 255, (v >> 8) & 255, (v >> 16) & 255, (v >>> 24) & 255]
  const entries = 3
  const valuesAt = 8 + 2 + entries * 12 + 4
  const entry = (tag: number, type: number, count: number, value: number[]) => [
    ...le16(tag),
    ...le16(type),
    ...le32(count),
    ...value,
  ]
  const tiff = [
    ...ascii('II'),
    ...le16(42),
    ...le32(8),
    ...le16(entries),
    ...entry(0x011a, 5, 1, le32(valuesAt)),
    ...entry(0x011b, 5, 1, le32(valuesAt + 8)),
    ...entry(0x0128, 3, 1, [...le16(unit), 0, 0]),
    ...le32(0),
    ...le32(xRes[0]),
    ...le32(xRes[1]),
    ...le32(yRes[0]),
    ...le32(yRes[1]),
  ]
  return segment(0xe1, [...ascii('Exif'), 0, 0, ...tiff])
}

function jpeg(...segments: number[][]): Uint8Array {
  const sof0 = segment(0xc0, [8, ...be16(2), ...be16(2), 1, 1, 0x11, 0])
  return Uint8Array.from([0xff, 0xd8, ...segments.flat(), ...sof0, 0xff, 0xda, 0, 2, 0xff, 0xd9])
}

// ── placement (PowerPoint for Mac, 960x540 pt slide) ────────────────────

describe('pictureFrame', () => {
  const rows: {
    name: string
    px: [number, number]
    dpi?: number
    pt: [number, number]
    at?: [number, number]
  }[] = [
    { name: '800x600 untagged', px: [800, 600], pt: [400, 300], at: [280, 120] },
    { name: '200x150 untagged', px: [200, 150], pt: [100, 75] },
    { name: '800x600 @96', px: [800, 600], dpi: 96, pt: [600, 450], at: [180, 45] },
    { name: '800x600 @300', px: [800, 600], dpi: 300, pt: [192, 144] },
    { name: '800x600 @72 shrinks to 8.5in', px: [800, 600], dpi: 72, pt: [612, 459] },
    { name: '1600x900 untagged shrinks to 8.5in', px: [1600, 900], pt: [612, 344.25] },
    { name: '2000x1000 untagged', px: [2000, 1000], pt: [612, 306] },
    { name: '3000x2000 untagged', px: [3000, 2000], pt: [612, 408] },
    { name: '1000x3000 untagged fits the slide height', px: [1000, 3000], pt: [180, 540] },
  ]
  for (const row of rows) {
    it(row.name, () => {
      const dpi = row.dpi ? { x: row.dpi, y: row.dpi } : undefined
      const f = pictureFrame(SLIDE, { width: row.px[0], height: row.px[1] }, dpi, MAC_DPI)
      expect(f.cx).toBe(Math.round(row.pt[0] * PT))
      expect(f.cy).toBe(Math.round(row.pt[1] * PT))
      const at = row.at ?? [(960 - row.pt[0]) / 2, (540 - row.pt[1]) / 2]
      expect(f.x).toBe(Math.round(at[0] * PT))
      expect(f.y).toBe(Math.round(at[1] * PT))
    })
  }

  it('uses 96 dpi for untagged bitmaps on Windows', () => {
    const f = pictureFrame(SLIDE, { width: 800, height: 600 }, undefined, 96)
    expect([f.cx, f.cy]).toEqual([600 * PT, 450 * PT])
  })

  it('honors a non-square declared dpi per axis', () => {
    const f = pictureFrame(SLIDE, { width: 800, height: 600 }, { x: 200, y: 100 }, MAC_DPI)
    expect([f.cx, f.cy]).toEqual([288 * PT, 432 * PT])
  })

  it('caps at the slide width when the slide is narrower than 8.5in', () => {
    const narrow = { width: 400 * PT, height: 300 * PT }
    const f = pictureFrame(narrow, { width: 1600, height: 900 }, undefined, MAC_DPI)
    expect([f.cx, f.cy]).toEqual([400 * PT, 225 * PT])
    expect([f.x, f.y]).toEqual([0, Math.round(37.5 * PT)])
  })

  it('centers on the drop point when one is given', () => {
    const f = pictureFrame(SLIDE, { width: 200, height: 150 }, undefined, MAC_DPI, {
      center: { x: 100 * PT, y: 100 * PT },
    })
    expect([f.x, f.y, f.cx, f.cy]).toEqual([50 * PT, 62.5 * PT, 100 * PT, 75 * PT])
  })

  it('grows a tiny banner proportionally to the minimum and centers the result', () => {
    // 24 renderer px on a 1280 px wide fit of this slide
    const minEmu = 24 * 9525
    const f = pictureFrame(SLIDE, { width: 40, height: 8 }, undefined, MAC_DPI, {
      center: { x: 100 * PT, y: 100 * PT },
      minEmu,
    })
    expect([f.cx, f.cy]).toEqual([120 * 9525, minEmu])
    expect([f.x, f.y]).toEqual([100 * PT - 60 * 9525, 100 * PT - 12 * 9525])
  })

  it('never lets the minimum push a capped strip past the fit box', () => {
    const f = pictureFrame(SLIDE, { width: 6000, height: 40 }, undefined, MAC_DPI, {
      minEmu: 24 * 9525,
    })
    expect(f.cx).toBe(612 * PT)
    expect(f.cy).toBe(Math.round(4.08 * PT))
    expect(f.cy).toBeLessThan(24 * 9525)
  })

  it('leaves a normal-size picture untouched by the minimum', () => {
    const f = pictureFrame(SLIDE, { width: 800, height: 600 }, undefined, MAC_DPI, {
      minEmu: 24 * 9525,
    })
    expect([f.cx, f.cy]).toEqual([400 * PT, 300 * PT])
  })

  it('falls back to a 4:3 frame at the cap when the pixel size is unreadable', () => {
    const f = pictureFrame(SLIDE, null, undefined, MAC_DPI)
    expect([f.cx, f.cy]).toEqual([612 * PT, 459 * PT])
  })
})

// ── dpi sniffing ─────────────────────────────────────────────────────────

describe('imageDpiFromBytes', () => {
  it('reads PNG pHYs in pixels per metre', () => {
    const dpi = imageDpiFromBytes(png({ ppmX: 11811, ppmY: 3780, unit: 1 }))
    expect(dpi?.x).toBeCloseTo(300, 1)
    expect(dpi?.y).toBeCloseTo(96, 1)
  })

  it('ignores PNG pHYs that only states an aspect ratio', () => {
    expect(imageDpiFromBytes(png({ ppmX: 1, ppmY: 1, unit: 0 }))).toBeUndefined()
    expect(imageDpiFromBytes(png())).toBeUndefined()
  })

  it('reads JFIF density in dpi and dpcm', () => {
    expect(imageDpiFromBytes(jpeg(jfif(1, 300, 300)))).toEqual({ x: 300, y: 300 })
    const cm = imageDpiFromBytes(jpeg(jfif(2, 118, 59)))
    expect(cm?.x).toBeCloseTo(299.72, 2)
    expect(cm?.y).toBeCloseTo(149.86, 2)
  })

  it('falls back to EXIF X/YResolution when JFIF declares no density', () => {
    expect(imageDpiFromBytes(jpeg(exif([300, 1], [300, 1], 2)))).toEqual({ x: 300, y: 300 })
    expect(imageDpiFromBytes(jpeg(jfif(0, 1, 1), exif([72, 1], [72, 1], 2)))).toEqual({
      x: 72,
      y: 72,
    })
    const cm = imageDpiFromBytes(jpeg(exif([1000, 10], [1000, 10], 3)))
    expect(cm?.x).toBeCloseTo(254, 5)
  })

  it('prefers JFIF density over EXIF', () => {
    expect(imageDpiFromBytes(jpeg(jfif(1, 96, 96), exif([300, 1], [300, 1], 2)))).toEqual({
      x: 96,
      y: 96,
    })
  })

  it('returns nothing for JPEGs without a tag and for GIF', () => {
    expect(imageDpiFromBytes(jpeg())).toBeUndefined()
    expect(
      imageDpiFromBytes(Uint8Array.from([...ascii('GIF89a'), 10, 0, 10, 0, ...Array(8).fill(0)])),
    ).toBeUndefined()
  })
})
