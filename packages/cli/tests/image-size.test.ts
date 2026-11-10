import { describe, expect, it } from 'vitest'
import { imageSize } from '../src/formats/image-size'

const PNG_1PX = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
  'base64',
)

function pngWithDims(width: number, height: number): Uint8Array {
  const out = Buffer.from(PNG_1PX)
  out.writeUInt32BE(width >>> 0, 16)
  out.writeUInt32BE(height >>> 0, 20)
  return new Uint8Array(out)
}

/** Minimal JPEG: SOI, APP0, RST0, SOF0 (w x h), EOI. */
function jpegWithDims(width: number, height: number, rstFirst = false): Uint8Array {
  const app0 = [
    0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01,
    0x00, 0x00,
  ]
  const rst0 = [0xff, 0xd0]
  const sof0 = [
    0xff,
    0xc0,
    0x00,
    0x0b,
    0x08,
    (height >> 8) & 0xff,
    height & 0xff,
    (width >> 8) & 0xff,
    width & 0xff,
    0x01,
    0x01,
    0x11,
    0x00,
  ]
  const bytes = [0xff, 0xd8, ...app0, ...(rstFirst ? rst0 : []), ...sof0, 0xff, 0xd9]
  return new Uint8Array(bytes)
}

describe('imageSize', () => {
  it('reads real PNG dims', () => {
    expect(imageSize(new Uint8Array(PNG_1PX))).toEqual({ width: 1, height: 1, mime: 'image/png' })
  })

  it('rejects zero and absurd PNG dims', () => {
    expect(imageSize(pngWithDims(0, 1))).toBeNull()
    expect(imageSize(pngWithDims(200000, 1))).toBeNull()
    expect(imageSize(pngWithDims(100, 100))).toEqual({ width: 100, height: 100, mime: 'image/png' })
  })

  it('skips length-less RST markers instead of misparsing', () => {
    expect(imageSize(jpegWithDims(40, 30, true))).toEqual({
      width: 40,
      height: 30,
      mime: 'image/jpeg',
    })
    expect(imageSize(jpegWithDims(40, 30, false))).toEqual({
      width: 40,
      height: 30,
      mime: 'image/jpeg',
    })
  })

  it('rejects zero JPEG frame dims', () => {
    expect(imageSize(jpegWithDims(0, 30))).toBeNull()
  })

  it('returns null for non-images', () => {
    expect(imageSize(new Uint8Array([1, 2, 3, 4]))).toBeNull()
    expect(imageSize(new Uint8Array(0))).toBeNull()
  })
})
