import { describe, expect, it } from 'vitest'
import { intrinsicSize, sniffImage } from '../src/units'

/** a 24-byte PNG prefix: signature, IHDR length/type, then the declared dimensions */
function png(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(24)
  bytes.set(
    [
      0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
      0x52,
    ],
    0,
  )
  const view = new DataView(bytes.buffer)
  view.setUint32(16, width)
  view.setUint32(20, height)
  return bytes
}

/** GIF89a header: six magic bytes, then two little-endian dimensions */
function gif(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(10)
  bytes.set([0x47, 0x49, 0x46, 0x38, 0x39, 0x61], 0)
  const view = new DataView(bytes.buffer)
  view.setUint16(6, width, true)
  view.setUint16(8, height, true)
  return bytes
}

/** SOI + one APP0 segment + a SOF0 carrying the dimensions */
function jpeg(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(32)
  const view = new DataView(bytes.buffer)
  bytes.set([0xff, 0xd8], 0)
  bytes.set([0xff, 0xe0], 2)
  view.setUint16(4, 4)
  bytes.set([0xff, 0xc0], 8)
  view.setUint16(10, 11)
  bytes[12] = 8
  view.setUint16(13, height)
  view.setUint16(15, width)
  return bytes
}

describe('intrinsicSize', () => {
  it('reads valid PNG, GIF and JPEG headers', () => {
    expect(sniffImage(png(640, 480))).toBe('png')
    expect(intrinsicSize(png(640, 480))).toEqual({ width: 640, height: 480 })
    expect(intrinsicSize(gif(320, 200))).toEqual({ width: 320, height: 200 })
    expect(intrinsicSize(jpeg(800, 600))).toEqual({ width: 800, height: 600 })
  })

  it('refuses zero and out-of-spec dimensions', () => {
    for (const bytes of [
      png(0, 480),
      png(640, 0),
      png(0xffffffff, 1),
      png(200_001, 1),
      gif(0, 200),
      gif(320, 0),
    ]) {
      expect(intrinsicSize(bytes)).toBeNull()
    }
  })

  it('applies the same rule to dimensions read from a JPEG SOF header', () => {
    expect(intrinsicSize(jpeg(0, 600))).toBeNull()
    expect(intrinsicSize(jpeg(800, 0))).toBeNull()
  })

  it('keeps a valid large photo at its own aspect', () => {
    expect(intrinsicSize(jpeg(8688, 5792))).toEqual({ width: 8688, height: 5792 })
    expect(intrinsicSize(png(10_000, 10_000))).toEqual({ width: 10_000, height: 10_000 })
  })
})
