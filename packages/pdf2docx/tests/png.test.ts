import { describe, expect, it } from 'vitest'
import { encodeRgbaPng } from '../src/extract'

function solidRgba(w: number, h: number, rgba: [number, number, number, number]): Uint8Array {
  const buf = new Uint8Array(w * h * 4)
  for (let i = 0; i < w * h; i++) buf.set(rgba, i * 4)
  return buf
}

describe('encodeRgbaPng', () => {
  it('produces a well-formed PNG signature and IHDR', () => {
    const png = encodeRgbaPng(solidRgba(3, 2, [255, 0, 0, 255]), 3, 2)
    expect([...png.slice(0, 8)]).toEqual([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    // IHDR: width 3, height 2, bit depth 8, color type 6 (RGBA)
    const view = new DataView(png.buffer, png.byteOffset)
    expect(view.getUint32(16)).toBe(3)
    expect(view.getUint32(20)).toBe(2)
    expect(png[24]).toBe(8)
    expect(png[25]).toBe(6)
  })

  it('round-trips through a real PNG decoder (pdf-lib embedPng)', async () => {
    const { PDFDocument } = await import('pdf-lib')
    const png = encodeRgbaPng(solidRgba(16, 8, [0, 128, 255, 255]), 16, 8)
    const doc = await PDFDocument.create()
    const image = await doc.embedPng(png) // parses IHDR/IDAT incl. zlib inflate
    expect(image.width).toBe(16)
    expect(image.height).toBe(8)
  })

  it('handles buffers that need multiple stored deflate blocks (>64KB)', async () => {
    const { PDFDocument } = await import('pdf-lib')
    const w = 200
    const h = 120 // 96000 bytes raw > 65535 → 2 stored blocks
    const png = encodeRgbaPng(solidRgba(w, h, [10, 20, 30, 255]), w, h)
    const doc = await PDFDocument.create()
    const image = await doc.embedPng(png)
    expect(image.width).toBe(w)
    expect(image.height).toBe(h)
  })

  it('rejects a size mismatch', () => {
    expect(() => encodeRgbaPng(new Uint8Array(10), 2, 2)).toThrow(/expected/)
  })
})
