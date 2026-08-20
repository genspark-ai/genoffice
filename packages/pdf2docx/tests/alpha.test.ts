/**
 * P9 A: image transparency survives extraction. A JPEG whose alpha lives in a
 * separate SMask must NOT take the raw-JPEG passthrough (the bare DCT stream
 * shows a black matte where the PDF was transparent) — it renders through
 * PDFium and lands as a PNG that keeps the transparent pixels.
 */
import { describe, expect, it } from 'vitest'
import { extractPage, withPdfDocument } from '../src/extract'
import { buildJpegPdf, buildMaskedJpegPdf } from './helpers/fixtures'
import { loadPdfium } from './helpers/wasm'

/** decode our own encoder's output (stored-deflate zlib, filter 0 rows) */
function decodeStoredPng(png: Uint8Array): { rgba: Uint8Array; width: number; height: number } {
  const view = new DataView(png.buffer, png.byteOffset, png.byteLength)
  const width = view.getUint32(16)
  const height = view.getUint32(20)
  // concatenate IDAT payloads
  let idat = new Uint8Array(0)
  let pos = 8
  while (pos < png.length) {
    const len = view.getUint32(pos)
    const type = String.fromCharCode(png[pos + 4]!, png[pos + 5]!, png[pos + 6]!, png[pos + 7]!)
    if (type === 'IDAT') {
      const merged = new Uint8Array(idat.length + len)
      merged.set(idat)
      merged.set(png.subarray(pos + 8, pos + 8 + len), idat.length)
      idat = merged
    }
    pos += 12 + len
  }
  // zlib stored blocks: 2-byte header, then 5-byte block headers + raw data
  const raw = new Uint8Array(height * (1 + width * 4))
  let src = 2
  let dst = 0
  for (;;) {
    const final = idat[src]! & 1
    const len = idat[src + 1]! | (idat[src + 2]! << 8)
    raw.set(idat.subarray(src + 5, src + 5 + len), dst)
    src += 5 + len
    dst += len
    if (final) break
  }
  // strip the per-row filter byte (always 0)
  const rgba = new Uint8Array(width * height * 4)
  for (let y = 0; y < height; y++) {
    rgba.set(raw.subarray(y * (1 + width * 4) + 1, (y + 1) * (1 + width * 4)), y * width * 4)
  }
  return { rgba, width, height }
}

describe('image alpha extraction (P9 A)', () => {
  it('a JPEG with an SMask becomes a PNG that keeps transparent pixels', async () => {
    const m = await loadPdfium()
    const pdf = await buildMaskedJpegPdf()
    const page = withPdfDocument(m, pdf, (doc) => extractPage(m, doc, 0))

    expect(page.images).toHaveLength(1)
    const img = page.images[0]!
    expect(img.mime).toBe('image/png')

    const { rgba, width, height } = decodeStoredPng(img.data)
    let transparent = 0
    let opaque = 0
    for (let i = 3; i < rgba.length; i += 4) {
      if (rgba[i]! < 128) transparent++
      else if (rgba[i]! > 230) opaque++
    }
    // the SMask blanks the left half — both kinds must be present
    expect(transparent).toBeGreaterThan((width * height) / 4)
    expect(opaque).toBeGreaterThan((width * height) / 4)
    // an opaque pixel keeps the JPEG's red
    for (let i = 0; i < rgba.length; i += 4) {
      if (rgba[i + 3]! > 230) {
        expect(rgba[i]!).toBeGreaterThan(150)
        expect(rgba[i + 1]!).toBeLessThan(120)
        break
      }
    }
  })

  it('a plain opaque JPEG still passes through as image/jpeg', async () => {
    const m = await loadPdfium()
    const pdf = await buildJpegPdf()
    const page = withPdfDocument(m, pdf, (doc) => extractPage(m, doc, 0))

    expect(page.images).toHaveLength(1)
    const img = page.images[0]!
    expect(img.mime).toBe('image/jpeg')
    // JPEG signature intact (raw stream, not re-encoded)
    expect([img.data[0], img.data[1]]).toEqual([0xff, 0xd8])
    expect(img.pixelWidth).toBe(8)
    expect(img.pixelHeight).toBe(8)
  })
})
