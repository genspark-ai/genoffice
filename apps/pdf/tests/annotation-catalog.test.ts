import type { PDFDocumentProxy } from 'pdfjs-dist'
import { describe, expect, it, vi } from 'vitest'
import { createSavedAnnotCountsLoader } from '../src/renderer/annotation-catalog'
import type { SavedMarkupAnnot } from '../src/renderer/edit-state'
import type { SavedNoteAnnot } from '../src/renderer/note-threads'

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
