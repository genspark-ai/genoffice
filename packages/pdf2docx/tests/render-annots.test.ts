import { describe, expect, it } from 'vitest'
import { renderPageByIndexPng, withPdfDocument } from '../src/extract'
import { decodePngRgba } from './helpers/png-decode'
import { loadPdfium } from './helpers/wasm'

/** 300x400 page whose only painted content is a red /Square annotation */
async function annotatedPdf(): Promise<Uint8Array> {
  const { PDFDocument, PDFName } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([300, 400])
  // the appearance stream a viewer paints: solid red over the annotation's box
  const appearance = doc.context.stream('1 0 0 rg 0 0 100 100 re f', {
    Type: 'XObject',
    Subtype: 'Form',
    FormType: 1,
    BBox: [0, 0, 100, 100],
    Resources: {},
  })
  const annot = doc.context.obj({
    Type: 'Annot',
    Subtype: 'Square',
    Rect: [50, 200, 150, 300],
    F: 4,
    C: [1, 0, 0],
    Border: [0, 0, 0],
    AP: doc.context.obj({ N: doc.context.register(appearance) }),
  })
  page.node.set(PDFName.of('Annots'), doc.context.obj([doc.context.register(annot)]))
  return doc.save({ useObjectStreams: false })
}

describe('full-page PNG render', () => {
  it('paints the drawing a saved annotation carries', async () => {
    // A mark-up drawn in the app is saved as an annotation with an appearance
    // stream, not as page content. Rendering the page without FPDF_ANNOT left
    // the bitmap blank, so a rasterized page lost every authored drawing.
    const m = await loadPdfium()
    const png = withPdfDocument(m, await annotatedPdf(), (doc) =>
      renderPageByIndexPng(m, doc, 0, 1),
    )
    expect(png).not.toBeNull()
    const { rgba, width, height } = decodePngRgba(png!.data)
    expect(width).toBe(300)
    expect(height).toBe(400)
    let red = 0
    for (let i = 0; i < rgba.length; i += 4) {
      if (rgba[i]! > 200 && rgba[i + 1]! < 80 && rgba[i + 2]! < 80) red++
    }
    expect(red).toBeGreaterThan(1000)
  })
})
