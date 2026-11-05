import { describe, expect, it } from 'vitest'
import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFString } from 'pdf-lib'
import { applySaveRequest } from '../src/main/save-pdf'
import type { DrawingInput, MarkupInput, SavePdfRequest } from '../src/shared/ipc'

/** A 1×1 PNG, for the image stamp. */
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=='

const highlight: MarkupInput = {
  pageIndex: 0,
  type: 'highlight',
  color: [1, 1, 0],
  quads: [[100, 220, 200, 220, 100, 200, 200, 200]],
}

const drawings: DrawingInput[] = [
  { kind: 'rect', pageIndex: 0, color: [1, 0, 0], width: 1, rect: [10, 10, 60, 60] },
  { kind: 'ink', pageIndex: 0, color: [0, 0, 1], width: 1, paths: [[10, 100, 40, 130]] },
  { kind: 'caret', pageIndex: 0, color: [0, 0, 0], point: [300, 300] },
  {
    kind: 'polygon',
    pageIndex: 0,
    color: [0, 1, 0],
    width: 1,
    vertices: [
      [400, 400],
      [450, 400],
      [425, 440],
    ],
  },
  { kind: 'image', pageIndex: 0, image: PNG, rect: [500, 500, 520, 520] },
  { kind: 'note', pageIndex: 0, color: [1, 1, 0], at: [50, 700], contents: 'mine' },
  {
    kind: 'note',
    pageIndex: 0,
    color: [1, 1, 0],
    at: [80, 700],
    contents: 'theirs',
    author: 'AI Assistant',
  },
]

async function savedAuthors(annotationAuthor?: string): Promise<Map<string, unknown[]>> {
  const doc = await PDFDocument.create()
  doc.addPage([612, 792])
  const request: SavePdfRequest = {
    path: '/tmp/test.pdf',
    markups: [highlight],
    drawings,
    formValues: [],
    stamps: [],
    ...(annotationAuthor === undefined ? {} : { annotationAuthor }),
  }
  const { bytes } = await applySaveRequest(await doc.save(), request)
  const saved = await PDFDocument.load(bytes, { updateMetadata: false })
  const { context } = saved
  const annots = context.lookup(saved.getPages()[0]!.node.get(PDFName.of('Annots')), PDFArray)
  const bySubtype = new Map<string, unknown[]>()
  for (let index = 0; index < annots.size(); index += 1) {
    const annot = context.lookup(annots.get(index), PDFDict)
    const subtype = context.lookup(annot.get(PDFName.of('Subtype')), PDFName).decodeText()
    const t = annot.get(PDFName.of('T'))
    // /T is a text string (ISO 32000-1, 12.5.6.2); anything else is reported as is
    const author = t instanceof PDFHexString || t instanceof PDFString ? t.decodeText() : t
    bySubtype.set(subtype, [...(bySubtype.get(subtype) ?? []), author])
  }
  return bySubtype
}

describe('the author of a saved annotation', () => {
  it('is the requested author, as a text string, on every kind the save adds', async () => {
    const authors = await savedAuthors('Ann')
    expect(Object.fromEntries(authors)).toEqual({
      Highlight: ['Ann'],
      Square: ['Ann'],
      Ink: ['Ann'],
      Caret: ['Ann'],
      Polygon: ['Ann'],
      Stamp: ['Ann'],
      Text: ['Ann', 'AI Assistant'],
    })
  })

  it('falls back to GenOffice, still as a text string, when no author is known', async () => {
    const authors = await savedAuthors()
    expect(authors.get('Highlight')).toEqual(['GenOffice'])
    expect(authors.get('Ink')).toEqual(['GenOffice'])
  })
})
