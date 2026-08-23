import type { PDFDocumentProxy } from 'pdfjs-dist'
import { MARKUP_TYPE_BY_ANNOT, type SavedMarkupAnnot } from './edit-state'
import { toSavedNote, type PdfJsAnnotData, type SavedNoteAnnot } from './note-threads'

export interface PageSavedAnnots {
  markups: SavedMarkupAnnot[]
  notes: SavedNoteAnnot[]
}

/**
 * Saved markup + note annotations of one page, read straight from pdf.js.
 * Single source for both the lazy visible-page cache and the AI tools (which
 * need arbitrary pages, not just the ones scrolled into view).
 */
export async function loadSavedAnnots(
  doc: PDFDocumentProxy,
  origIdx: number,
): Promise<PageSavedAnnots> {
  try {
    const page = await doc.getPage(origIdx + 1)
    const annots = (await page.getAnnotations()) as (PdfJsAnnotData & {
      quadPoints?: Float32Array | null
    })[]
    const markups = annots.flatMap((a) => {
      const type = MARKUP_TYPE_BY_ANNOT[a.annotationType]
      // Only ref-backed annots can be addressed for deletion (id "123R" → object 123)
      const objNum = /^(\d+)R$/.exec(a.id)
      if (!type || !objNum || !a.quadPoints || a.quadPoints.length < 8) return []
      const quads: number[][] = []
      for (let q = 0; q + 8 <= a.quadPoints.length; q += 8)
        quads.push([...a.quadPoints.slice(q, q + 8)])
      return [
        {
          pageIndex: origIdx,
          objNum: Number(objNum[1]),
          type,
          quads,
          rect: [a.rect[0]!, a.rect[1]!, a.rect[2]!, a.rect[3]!] as [
            number,
            number,
            number,
            number,
          ],
        },
      ]
    })
    const notes = annots.flatMap((a) => {
      const note = toSavedNote(a, origIdx)
      return note ? [note] : []
    })
    return { markups, notes }
  } catch {
    return { markups: [], notes: [] } // page unreadable; no saved annotations to offer
  }
}
