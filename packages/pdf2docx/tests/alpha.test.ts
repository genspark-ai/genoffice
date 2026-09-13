/**
 * P9 A: image transparency survives extraction. A JPEG whose alpha lives in a
 * separate SMask must NOT take the raw-JPEG passthrough (the bare DCT stream
 * shows a black matte where the PDF was transparent) — it renders through
 * PDFium and lands as a PNG that keeps the transparent pixels.
 */
import { describe, expect, it } from 'vitest'
import { extractPage, withPdfDocument } from '../src/extract'
import { buildJpegPdf, buildMaskedJpegPdf } from './helpers/fixtures'
import { decodePngRgba } from './helpers/png-decode'
import { loadPdfium } from './helpers/wasm'

describe('image alpha extraction (P9 A)', () => {
  it('a JPEG with an SMask becomes a PNG that keeps transparent pixels', async () => {
    const m = await loadPdfium()
    const pdf = await buildMaskedJpegPdf()
    const page = withPdfDocument(m, pdf, (doc) => extractPage(m, doc, 0))

    expect(page.images).toHaveLength(1)
    const img = page.images[0]!
    expect(img.mime).toBe('image/png')

    const { rgba, width, height } = decodePngRgba(img.data)
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
