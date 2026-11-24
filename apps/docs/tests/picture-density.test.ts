import { describe, expect, it } from 'vitest'
import { originalPictureSize, pictureDensityOf } from '../src/renderer/editor/picture-density'

const u32 = (n: number) => [(n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255]
const u16 = (n: number) => [(n >> 8) & 255, n & 255]
const chars = (text: string) => [...text].map((ch) => ch.charCodeAt(0))

const chunk = (type: string, data: number[]) => [
  ...u32(data.length),
  ...chars(type),
  ...data,
  ...u32(0),
]

function png(...chunks: number[][]): Uint8Array {
  return Uint8Array.from([
    0x89,
    0x50,
    0x4e,
    0x47,
    0x0d,
    0x0a,
    0x1a,
    0x0a,
    ...chunk('IHDR', [...u32(600), ...u32(300), 8, 6, 0, 0, 0]),
    ...chunks.flat(),
    ...chunk('IDAT', [0]),
    ...chunk('IEND', []),
  ])
}

const segment = (marker: number, body: number[]) => [0xff, marker, ...u16(body.length + 2), ...body]

function jpeg(...segments: number[][]): Uint8Array {
  return Uint8Array.from([0xff, 0xd8, ...segments.flat(), 0xff, 0xda, 0, 2, 0xff, 0xd9])
}

const jfif = (unit: number, x: number, y: number) =>
  segment(0xe0, [...chars('JFIF\0'), 1, 2, unit, ...u16(x), ...u16(y), 0, 0])

/** an Exif segment with XResolution, YResolution and ResolutionUnit in either byte order */
function exif(little: boolean, x: number, y: number, unit: number): number[] {
  const s = (n: number) => (little ? u16(n).reverse() : u16(n))
  const l = (n: number) => (little ? u32(n).reverse() : u32(n))
  const entry = (tag: number, type: number, value: number[]) => [
    ...s(tag),
    ...s(type),
    ...l(1),
    ...value,
  ]
  const ifdEnd = 8 + 2 + 3 * 12 + 4
  return segment(0xe1, [
    ...chars('Exif\0\0'),
    ...chars(little ? 'II' : 'MM'),
    ...s(42),
    ...l(8),
    ...s(3),
    ...entry(0x011a, 5, l(ifdEnd)),
    ...entry(0x011b, 5, l(ifdEnd + 8)),
    ...entry(0x0128, 3, [...s(unit), 0, 0]),
    ...l(0),
    ...l(x),
    ...l(1),
    ...l(y),
    ...l(1),
  ])
}

describe('the density a picture file states', () => {
  it('reads a PNG pHYs chunk given in pixels per metre', () => {
    expect(pictureDensityOf(png(chunk('pHYs', [...u32(11811), ...u32(11811), 1])))).toEqual({
      xDpi: 300,
      yDpi: 300,
    })
    expect(pictureDensityOf(png(chunk('pHYs', [...u32(2835), ...u32(5669), 1])))).toEqual({
      xDpi: 72,
      yDpi: 144,
    })
  })

  it('finds pHYs behind other chunks, and reads none past the image data', () => {
    const phys = chunk('pHYs', [...u32(5906), ...u32(5906), 1])
    expect(pictureDensityOf(png(chunk('gAMA', u32(45455)), phys))).toEqual({
      xDpi: 150,
      yDpi: 150,
    })
    expect(pictureDensityOf(png(chunk('IDAT', [0]), phys))).toBeNull()
  })

  it('takes a PNG without pHYs, or with a pHYs that only gives the pixel shape, as stating none', () => {
    expect(pictureDensityOf(png())).toBeNull()
    expect(pictureDensityOf(png(chunk('pHYs', [...u32(2), ...u32(1), 0])))).toBeNull()
    expect(pictureDensityOf(png(chunk('pHYs', [...u32(0), ...u32(0), 1])))).toBeNull()
  })

  it('reads a JFIF density in dots per inch and per centimetre', () => {
    expect(pictureDensityOf(jpeg(jfif(1, 300, 300)))).toEqual({ xDpi: 300, yDpi: 300 })
    expect(pictureDensityOf(jpeg(jfif(2, 118, 118)))).toEqual({ xDpi: 299.72, yDpi: 299.72 })
  })

  it('reads the Exif resolution in either byte order when JFIF gives no unit', () => {
    expect(pictureDensityOf(jpeg(jfif(0, 1, 1), exif(false, 240, 240, 2)))).toEqual({
      xDpi: 240,
      yDpi: 240,
    })
    expect(pictureDensityOf(jpeg(exif(true, 200, 100, 2)))).toEqual({ xDpi: 200, yDpi: 100 })
    expect(pictureDensityOf(jpeg(exif(true, 100, 100, 3)))).toEqual({ xDpi: 254, yDpi: 254 })
  })

  it('prefers a JFIF density with a unit over Exif', () => {
    expect(pictureDensityOf(jpeg(jfif(1, 72, 72), exif(false, 300, 300, 2)))).toEqual({
      xDpi: 72,
      yDpi: 72,
    })
  })

  it('states none for a JPEG with neither, another format, or a cut-off file', () => {
    expect(pictureDensityOf(jpeg(jfif(0, 1, 1)))).toBeNull()
    expect(pictureDensityOf(jpeg())).toBeNull()
    expect(pictureDensityOf(Uint8Array.from(chars('GIF89a')))).toBeNull()
    expect(pictureDensityOf(new Uint8Array(0))).toBeNull()
    const whole = png(chunk('pHYs', [...u32(11811), ...u32(11811), 1]))
    expect(pictureDensityOf(whole.subarray(0, 40))).toBeNull()
    expect(pictureDensityOf(jpeg(exif(true, 200, 100, 2)).subarray(0, 30))).toBeNull()
  })
})

/** a BMP header: the file header, then a DIB header of `dib` bytes with pels per metre in an info header */
function bmp(dib: number, xPpm = 0, yPpm = 0): Uint8Array {
  const bytes = new Uint8Array(14 + Math.max(dib, 12))
  const view = new DataView(bytes.buffer)
  bytes.set(chars('BM'))
  view.setUint32(14, dib, true)
  if (dib >= 40) {
    view.setInt32(38, xPpm, true)
    view.setInt32(42, yPpm, true)
  }
  return bytes
}

describe('the density a BMP states', () => {
  it('is the pels per metre of its info header', () => {
    expect(pictureDensityOf(bmp(40, 11811, 11811))).toEqual({ xDpi: 300, yDpi: 300 })
    expect(pictureDensityOf(bmp(124, 2835, 5669))).toEqual({ xDpi: 72, yDpi: 144 })
  })

  it('is none when the header leaves it at zero, is the old core header, or is cut off', () => {
    expect(pictureDensityOf(bmp(40))).toBeNull()
    expect(pictureDensityOf(bmp(12))).toBeNull()
    expect(pictureDensityOf(bmp(40, 11811, 11811).subarray(0, 40))).toBeNull()
  })
})

describe('the original size of a picture', () => {
  it('is its pixels at the stated density, in CSS px', () => {
    expect(originalPictureSize({ width: 600, height: 300 }, { xDpi: 300, yDpi: 300 })).toEqual({
      width: 192,
      height: 96,
    })
    expect(originalPictureSize({ width: 72, height: 72 }, { xDpi: 72, yDpi: 144 })).toEqual({
      width: 96,
      height: 48,
    })
  })

  it('is its pixels at 96 dpi when the file states no density', () => {
    expect(originalPictureSize({ width: 600, height: 300 }, null)).toEqual({
      width: 600,
      height: 300,
    })
  })
})
