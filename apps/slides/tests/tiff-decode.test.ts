/** TIFF pictures must transcode to PNG for display (Chromium can't decode TIFF). */
import { describe, expect, it } from 'vitest'
import UTIF from 'utif2'
import { PNG } from 'pngjs'
import { tiffToPng } from '../src/main/tiff-decode'

function makeTiff(width: number, height: number): Uint8Array {
  const rgba = new Uint8Array(width * height * 4)
  for (let i = 0; i < rgba.length; i += 4) {
    rgba[i] = 200 // R
    rgba[i + 1] = 100 // G
    rgba[i + 2] = 50 // B
    rgba[i + 3] = 255
  }
  return new Uint8Array(UTIF.encodeImage(rgba, width, height))
}

describe('tiffToPng', () => {
  it('decodes a TIFF into a valid PNG with matching dimensions', () => {
    const decoded = tiffToPng(makeTiff(20, 12))
    expect(decoded).not.toBeNull()
    expect(decoded!.width).toBe(20)
    expect(decoded!.height).toBe(12)
    const png = PNG.sync.read(Buffer.from(decoded!.png))
    expect(png.width).toBe(20)
    expect(png.height).toBe(12)
    expect(png.data[0]).toBe(200)
    expect(png.data[1]).toBe(100)
    expect(png.data[2]).toBe(50)
  })

  it('returns null on garbage bytes instead of throwing', () => {
    expect(tiffToPng(new Uint8Array([1, 2, 3, 4, 5]))).toBeNull()
  })

  it('rejects giant IFD dimensions without allocating pixels', () => {
    // Minimal little-endian TIFF whose header claims 100000x100000 px:
    // must return null fast instead of attempting a 40GB RGBA buffer.
    const buf = new ArrayBuffer(8 + 2 + 2 * 12 + 4)
    const view = new DataView(buf)
    view.setUint8(0, 0x49) // 'II'
    view.setUint8(1, 0x49)
    view.setUint16(2, 42, true)
    view.setUint32(4, 8, true) // IFD offset
    view.setUint16(8, 2, true) // 2 entries
    const entry = (i: number, tag: number, value: number) => {
      const off = 10 + i * 12
      view.setUint16(off, tag, true)
      view.setUint16(off + 2, 4, true) // LONG
      view.setUint32(off + 4, 1, true)
      view.setUint32(off + 8, value, true)
    }
    entry(0, 256, 100000) // ImageWidth
    entry(1, 257, 100000) // ImageLength
    view.setUint32(10 + 2 * 12, 0, true) // next IFD
    const start = Date.now()
    expect(tiffToPng(new Uint8Array(buf))).toBeNull()
    expect(Date.now() - start).toBeLessThan(10000)
  })

  it('accepts a 600 dpi A3 scan (~69MP) header', () => {
    // Header-only TIFF: passes the budget check, then fails in decodeImage
    // (no strips) rather than being rejected up front by the pixel cap.
    const dims = tiffHeaderOnly(7016, 9933)
    expect(tiffToPng(dims)).toBeNull()
    expect(tiffToPng(tiffHeaderOnly(40000, 40000))).toBeNull()
  })

  it('skips an over-budget first IFD and decodes the next in-budget page', () => {
    const w = 6
    const h = 4
    const rgba = new Uint8Array(w * h * 4).fill(90)
    const hostile = { t256: [100000], t257: [100000] }
    const real = {
      t256: [w],
      t257: [h],
      t258: [8, 8, 8, 8],
      t259: [1],
      t262: [2],
      t273: [1000],
      t277: [4],
      t278: [h],
      t279: [w * h * 4],
      t284: [1],
      t338: [1],
    }
    const prefix = new Uint8Array(UTIF.encode([hostile, real] as unknown as UTIF.IFD[]))
    expect(prefix.length).toBeLessThan(1000)
    const file = new Uint8Array(1000 + rgba.length)
    file.set(prefix, 0)
    file.set(rgba, 1000)
    expect(UTIF.decode(Buffer.from(file))).toHaveLength(2)
    const decoded = tiffToPng(file)
    expect(decoded).not.toBeNull()
    expect(decoded!.width).toBe(w)
    expect(decoded!.height).toBe(h)
    expect(PNG.sync.read(Buffer.from(decoded!.png)).data[0]).toBe(90)
  })
})

function tiffHeaderOnly(width: number, height: number): Uint8Array {
  return new Uint8Array(UTIF.encode([{ t256: [width], t257: [height] }] as unknown as UTIF.IFD[]))
}
