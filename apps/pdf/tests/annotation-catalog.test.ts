import type { PDFDocumentProxy } from 'pdfjs-dist'
import { describe, expect, it, vi } from 'vitest'
import { createSavedAnnotCountsLoader, loadSavedAnnots } from '../src/renderer/annotation-catalog'
import type { SavedMarkupAnnot } from '../src/renderer/edit-state'
import type { PdfJsAnnotData, SavedNoteAnnot } from '../src/renderer/note-threads'

/** pdf.js hands out rect/quadPoints; a malformed file can leave either one unusable */
type StubAnnot = Omit<PdfJsAnnotData, 'rect'> & {
  rect?: number[]
  quadPoints?: Float32Array
}

const doc = { numPages: 2 } as PDFDocumentProxy
const markup = (pageIndex: number, objNum: number): SavedMarkupAnnot => ({
  pageIndex,
  objNum,
  type: 'highlight',
  quads: [],
  rect: [0, 0, 1, 1],
})
const note = (pageIndex: number, objNum: number, inReplyTo: number | null): SavedNoteAnnot => ({
  pageIndex,
  objNum,
  type: 'note',
  rect: [0, 0, 1, 1],
  color: null,
  author: '',
  contents: '',
  timeMs: null,
  inReplyTo,
})

describe('createSavedAnnotCountsLoader', () => {
  it('does not inspect any page until its result is requested', async () => {
    const loadPage = vi.fn(async () => ({ markups: [], notes: [] }))

    const loadCounts = createSavedAnnotCountsLoader(doc, loadPage)

    expect(loadPage).not.toHaveBeenCalled()
    await loadCounts()
    expect(loadPage).toHaveBeenCalledTimes(2)
  })

  it('shares one whole-document scan between concurrent requests', async () => {
    const loadPage = vi.fn(async (_doc: PDFDocumentProxy, pageIndex: number) => ({
      markups: pageIndex === 0 ? [markup(0, 1)] : [markup(1, 2), markup(1, 3)],
      notes: pageIndex === 0 ? [note(0, 4, null), note(0, 5, 4)] : [note(1, 6, null)],
    }))
    const loadCounts = createSavedAnnotCountsLoader(doc, loadPage)

    const first = loadCounts()
    const second = loadCounts()

    expect(second).toBe(first)
    await expect(first).resolves.toEqual({ threads: [1, 1], markups: [1, 2] })
    expect(loadPage).toHaveBeenCalledTimes(2)
  })

  it('stops before the next page when the document scan is cancelled', async () => {
    const controller = new AbortController()
    const loadPage = vi.fn(async () => {
      controller.abort()
      return { markups: [], notes: [] }
    })
    const loadCounts = createSavedAnnotCountsLoader(doc, loadPage, controller.signal)

    await loadCounts()

    expect(loadPage).toHaveBeenCalledTimes(1)
  })
})

function annotDoc(annots: StubAnnot[]): PDFDocumentProxy {
  return {
    numPages: 1,
    getPage: async () => ({ getAnnotations: async () => annots }),
  } as unknown as PDFDocumentProxy
}

/** Highlight 9 (MARKUP_TYPE_BY_ANNOT) with one quad and a usable rect */
const VALID_HIGHLIGHT: StubAnnot = {
  id: '7R',
  annotationType: 9,
  rect: [10, 700, 90, 712],
  quadPoints: new Float32Array([10, 712, 90, 712, 10, 700, 90, 700]),
}

/** Text 1 (PDFJS_ANNOT_TEXT) — a note, so the page is non-empty even if the catch fires */
const NOTE: StubAnnot = {
  id: '9R',
  annotationType: 1,
  rect: [0, 0, 20, 20],
}

const VALID_HIGHLIGHT_SAVED = {
  pageIndex: 0,
  objNum: 7,
  type: 'highlight',
  quads: [[10, 712, 90, 712, 10, 700, 90, 700]],
  rect: [10, 700, 90, 712],
}

describe('loadSavedAnnots', () => {
  it('skips a markup with a short rect and keeps the valid ones on the page', async () => {
    const malformed: StubAnnot = { ...VALID_HIGHLIGHT, id: '8R', rect: [0, 0] }

    const page = await loadSavedAnnots(annotDoc([VALID_HIGHLIGHT, malformed, NOTE]), 0)

    expect(page.markups).toEqual([VALID_HIGHLIGHT_SAVED])
    expect(page.notes).toHaveLength(1)
  })

  it('skips a markup with no rect at all instead of failing the whole page', async () => {
    const malformed: StubAnnot = { ...VALID_HIGHLIGHT, id: '8R', rect: undefined }

    const page = await loadSavedAnnots(annotDoc([VALID_HIGHLIGHT, malformed, NOTE]), 0)

    expect(page.markups).toEqual([VALID_HIGHLIGHT_SAVED])
    expect(page.notes).toHaveLength(1)
  })
})
