/**
 * Redaction image census cost: applying N rectangles used to re-walk every page
 * and re-read + re-SHA-256 every image N times over (one full census per
 * region). The census now runs once per apply, so the number of document walks
 * (PDFium page loads) drops from regions * pages to pages + regions — counted
 * here deterministically through a loadPdfium wrapper.
 */
import { describe, expect, it, vi } from 'vitest'
import { PDFDocument, PDFDict, PDFName, PDFRawStream, StandardFonts } from 'pdf-lib'
import { redactPdf } from '../src/main/redaction'

const pdfiumCounts = vi.hoisted(() => ({ loadPage: 0 }))

vi.mock('../src/main/text-edit', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/main/text-edit')>()
  return {
    ...actual,
    loadPdfium: async () => {
      const m = await actual.loadPdfium()
      return new Proxy(m, {
        get(target, prop) {
          if (prop === '_FPDF_LoadPage') {
            return (...args: unknown[]) => {
              pdfiumCounts.loadPage++
              return target._FPDF_LoadPage(...args)
            }
          }
          return Reflect.get(target, prop)
        },
      })
    },
  }
})

const PAGES = 4
const REGIONS = 12

async function scannedDoc(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  for (let p = 0; p < PAGES; p++) {
    const page = doc.addPage([612, 792])
    page.drawText(`CONFIDENTIAL SCAN PAGE ${p + 1}`, { x: 40, y: 700, size: 14, font })
    // unique 512x512 DeviceRGB pixels per page: every page has its own image object
    const pixels = new Uint8Array(512 * 512 * 3)
    for (let i = 0; i < pixels.length; i++) pixels[i] = (i * (p + 3)) & 0xff
    const image = doc.context.register(
      PDFRawStream.of(
        doc.context.obj({
          Type: 'XObject',
          Subtype: 'Image',
          Width: 512,
          Height: 512,
          ColorSpace: 'DeviceRGB',
          BitsPerComponent: 8,
        }),
        pixels,
      ),
    )
    const name = page.node.newXObject('Image', image)
    page.node.addContentStream(
      doc.context.register(
        PDFRawStream.of(
          PDFDict.withContext(doc.context),
          Buffer.from(`q\n500 0 0 500 56 100 cm\n${name.toString()} Do\nQ\n`),
        ),
      ),
    )
  }
  return doc.save({ useObjectStreams: false })
}

describe('redaction image census', () => {
  it('walks the document once per apply, not once per region', async () => {
    const bytes = await scannedDoc()
    const regions = Array.from({ length: REGIONS }, (_, i) => ({
      pageIndex: 0,
      rect: [40 + i * 8, 695, 46 + i * 8, 712] as [number, number, number, number],
    }))
    pdfiumCounts.loadPage = 0
    const out = await redactPdf(bytes, regions)
    expect(out.length).toBeGreaterThan(0)
    // one census walk over all pages, then one preflight + one apply load per region
    expect(pdfiumCounts.loadPage).toBe(PAGES + 2 * REGIONS)
  })

  it('still refuses a region over a shared image', async () => {
    const doc = await PDFDocument.create()
    const pixels = new Uint8Array(8 * 8 * 3).fill(7)
    const stream = doc.context.register(
      PDFRawStream.of(
        doc.context.obj({
          Type: 'XObject',
          Subtype: 'Image',
          Width: 8,
          Height: 8,
          ColorSpace: 'DeviceRGB',
          BitsPerComponent: 8,
        }),
        pixels,
      ),
    )
    for (const _ of [0, 1]) {
      const page = doc.addPage([100, 100])
      const name = page.node.newXObject('Image', stream)
      page.node.addContentStream(
        doc.context.register(
          PDFRawStream.of(
            PDFDict.withContext(doc.context),
            Buffer.from(`q\n80 0 0 80 10 10 cm\n${name.toString()} Do\nQ\n`),
          ),
        ),
      )
    }
    await expect(
      redactPdf(await doc.save({ useObjectStreams: false }), [
        { pageIndex: 0, rect: [20, 20, 60, 60] },
      ]),
    ).rejects.toThrow(/shared image/)
  })
})
