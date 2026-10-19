import { describe, expect, it } from 'vitest'
import { PDFArray, PDFDict, PDFDocument, PDFName } from 'pdf-lib'
import { addPdfInternalLinks, sanitizePdfLinks } from '../src/main/pdf-links'

describe('sanitizePdfLinks', () => {
  it('drops malformed, out-of-range and out-of-page entries', () => {
    const good = { page: 0, x: 0.1, y: 0.2, w: 0.3, h: 0.05, toPage: 1, toY: 0.5 }
    expect(
      sanitizePdfLinks(
        [
          good,
          { ...good, page: 2 },
          { ...good, toPage: -1 },
          { ...good, x: 0.9, w: 0.2 },
          { ...good, h: 0 },
          { ...good, toY: Number.NaN },
          null,
          'x',
        ],
        2,
      ),
    ).toEqual([good])
    expect(sanitizePdfLinks('nope', 2)).toEqual([])
  })
})

describe('addPdfInternalLinks', () => {
  it('stamps GoTo annotations in PDF coordinates and keeps existing annots', async () => {
    const pdf = await PDFDocument.create()
    const p0 = pdf.addPage([200, 400])
    const p1 = pdf.addPage([300, 600])
    const existing = pdf.context.register(
      pdf.context.obj({ Type: 'Annot', Subtype: 'Link', Rect: [0, 0, 1, 1] }),
    )
    p0.node.set(PDFName.of('Annots'), pdf.context.obj([existing]))

    const n = addPdfInternalLinks(pdf, [
      { page: 0, x: 0.1, y: 0.25, w: 0.5, h: 0.05, toPage: 1, toY: 0.5 },
      { page: 5, x: 0, y: 0, w: 1, h: 1, toPage: 0, toY: 0 },
    ])
    expect(n).toBe(1)

    const annots = p0.node.lookup(PDFName.of('Annots'), PDFArray)
    expect(annots.size()).toBe(2)
    const link = annots.lookup(1, PDFDict)
    expect(link.get(PDFName.of('Subtype'))?.toString()).toBe('/Link')
    const rect = link
      .lookup(PDFName.of('Rect'), PDFArray)
      .asArray()
      .map((v) => Number(v.toString()))
    expect(rect).toEqual([20, 280, 120, 300])
    const dest = link.lookup(PDFName.of('Dest'), PDFArray)
    expect(dest.get(0)).toBe(p1.ref)
    expect(dest.get(1).toString()).toBe('/XYZ')
    expect(Number(dest.get(3).toString())).toBe(300)
    expect(p1.node.lookup(PDFName.of('Annots'))).toBeUndefined()
  })
})
