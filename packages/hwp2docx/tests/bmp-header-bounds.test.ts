/** Malformed BMP headers must be dropped, not decoded into a wrong picture. */
import { describe, expect, it } from 'vitest'
import { bmpToPng } from '../src/bmp'

const W = 2
const H = 1
const COLORS = 2
const PIXELS = 4 // one padded 8-bit row

interface BmpOpts {
  headerSize?: number
  clrUsed?: number
  pixelIndex?: number
}

/** 8-bit BMP: entry 0 red, entry 1 blue, both in the palette. */
function indexedBmp(opts: BmpOpts = {}): Uint8Array {
  const headerSize = opts.headerSize ?? 40
  const offset = 14 + 40 + COLORS * 4
  const bytes = new Uint8Array(offset + PIXELS * H)
  const view = new DataView(bytes.buffer)
  bytes[0] = 0x42
  bytes[1] = 0x4d
  view.setUint32(10, offset, true)
  // The override is written as a bogus declared header size; the byte layout
  // stays the valid 40-byte one so only the bounds check can reject it.
  view.setUint32(14, headerSize, true)
  view.setInt32(18, W, true)
  view.setInt32(22, H, true)
  view.setUint16(26, 1, true)
  view.setUint16(28, 8, true)
  view.setUint32(46, opts.clrUsed ?? COLORS, true)
  // BGR entries at 14 + 40
  bytes[54] = 0
  bytes[55] = 0
  bytes[56] = 255
  bytes[58] = 255
  bytes[59] = 0
  bytes[60] = 0
  bytes[offset] = opts.pixelIndex ?? 1
  return bytes
}

const signature = (png: Uint8Array): boolean =>
  png.length > 8 && png[0] === 0x89 && png[1] === 0x50 && png[2] === 0x4e && png[3] === 0x47

describe('bmpToPng header bounds', () => {
  it('accepts a well-formed indexed BMP', () => {
    const png = bmpToPng(indexedBmp())
    expect(png).not.toBeNull()
    expect(signature(png!)).toBe(true)
  })

  it('rejects a declared header larger than the file', () => {
    const bytes = indexedBmp()
    new DataView(bytes.buffer).setUint32(14, bytes.length + 1, true)
    expect(bmpToPng(bytes)).toBeNull()
  })

  it('rejects a palette that does not fit in the file', () => {
    // claims five entries but carries only two; the pixel names entry 4, which
    // used to decode as an invented black palette colour
    expect(bmpToPng(indexedBmp({ clrUsed: 5, pixelIndex: 4 }))).toBeNull()
  })

  it('still accepts a 24-bit BMP with no palette', () => {
    const bytes = new Uint8Array(54 + 4)
    const view = new DataView(bytes.buffer)
    bytes[0] = 0x42
    bytes[1] = 0x4d
    view.setUint32(10, 54, true)
    view.setUint32(14, 40, true)
    view.setInt32(18, 1, true)
    view.setInt32(22, 1, true)
    view.setUint16(26, 1, true)
    view.setUint16(28, 24, true)
    bytes[54] = 255 // B
    bytes[56] = 255 // R
    expect(signature(bmpToPng(bytes)!)).toBe(true)
  })
})
