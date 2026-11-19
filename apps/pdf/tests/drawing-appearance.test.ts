import { describe, expect, it } from 'vitest'
import { PDFArray, PDFDict, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { applySaveRequest } from '../src/main/save-pdf'
import type { DrawingInput, SavePdfRequest } from '../src/shared/ipc'

/**
 * Every drawn shape has to carry its own appearance stream.
 *
 * A viewer draws an annotation from its /AP when there is one, and from the
 * annotation's own entries when there is not. pdf.js has no default appearance
 * for a Caret and pdfium's generator covers neither Caret nor Polygon, so
 * both came back blank — the mark was in the file and nothing drew it, which
 * breaks the app's own round-trip.
 */
const request = (drawings: DrawingInput[]): SavePdfRequest => ({
  path: '/tmp/test.pdf',
  markups: [],
  drawings,
  formValues: [],
  stamps: [],
})

async function onePage(): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  doc.addPage([612, 792])
  return doc.save({ useObjectStreams: false })
}

/** The saved annotation of the given subtype, and the context it was read in. */
async function annotOf(
  bytes: Uint8Array,
  subtype: string,
): Promise<{ context: PDFDocument['context']; annot: PDFDict }> {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false })
  const { context } = doc
  const annots = context.lookup(doc.getPages()[0]!.node.get(PDFName.of('Annots')), PDFArray)
  for (let index = 0; index < annots.size(); index += 1) {
    const annot = context.lookup(annots.get(index), PDFDict)
    const found = context.lookup(annot.get(PDFName.of('Subtype')), PDFName).decodeText()
    if (found === subtype) return { context, annot }
  }
  throw new Error(`no ${subtype} annotation in the saved file`)
}

/** The normal appearance stream of an annotation. */
function appearanceOf(
  context: PDFDocument['context'],
  annot: PDFDict,
): { stream: PDFRawStream; contents: string } {
  const ap = context.lookup(annot.get(PDFName.of('AP')), PDFDict)
  const stream = context.lookup(ap.get(PDFName.of('N')), PDFRawStream)
  return { stream, contents: stream.getContentsString() }
}

describe('the appearance of a drawn shape', () => {
  it('gives the caret a wedge of its own to draw', async () => {
    const { bytes } = await applySaveRequest(
      await onePage(),
      request([{ kind: 'caret', pageIndex: 0, color: [0, 0, 0], point: [100, 200] }]),
    )
    const { context, annot } = await annotOf(bytes, 'Caret')
    const { stream, contents } = appearanceOf(context, annot)
    // a form the viewer can place, sized to the annotation's own rect
    expect(stream.dict.get(PDFName.of('Subtype'))?.toString()).toBe('/Form')
    expect(stream.dict.get(PDFName.of('BBox'))).toBeDefined()
    // the wedge itself: a filled triangle, not an empty box
    expect(contents).toContain(' f')
    expect(contents).toContain(' l')
  })

  it('gives the polygon a closed stroke of its own to draw', async () => {
    const { bytes } = await applySaveRequest(
      await onePage(),
      request([
        {
          kind: 'polygon',
          pageIndex: 0,
          color: [1, 0, 0],
          width: 2,
          vertices: [
            [10, 10],
            [50, 10],
            [30, 40],
          ],
        },
      ]),
    )
    const { context, annot } = await annotOf(bytes, 'Polygon')
    const { stream, contents } = appearanceOf(context, annot)
    expect(stream.dict.get(PDFName.of('Subtype'))?.toString()).toBe('/Form')
    expect(stream.dict.get(PDFName.of('BBox'))).toBeDefined()
    // the path is stroked and closed
    expect(contents).toContain('S')
    expect(contents).toContain('h')
  })

  it('leaves the shapes visible to the read path', async () => {
    const { bytes } = await applySaveRequest(
      await onePage(),
      request([
        { kind: 'caret', pageIndex: 0, color: [0, 0, 0], point: [100, 200] },
        {
          kind: 'polygon',
          pageIndex: 0,
          color: [1, 0, 0],
          width: 2,
          vertices: [
            [10, 10],
            [50, 10],
            [30, 40],
          ],
        },
      ]),
    )
    const loadingTask = getDocument({ data: bytes.slice() })
    try {
      const doc = await loadingTask.promise
      const page = await doc.getPage(1)
      const annots = await page.getAnnotations({ intent: 'display' })
      const bySubtype = new Map(annots.map((a) => [a.subtype, a]))
      // pdf.js hands both back to the viewer — with no /AP it would draw the
      // caret as nothing and rely on pdfium's (absent) polygon generator
      expect(bySubtype.has('Caret')).toBe(true)
      expect(bySubtype.has('Polygon')).toBe(true)
    } finally {
      await loadingTask.destroy()
    }
  })
})
