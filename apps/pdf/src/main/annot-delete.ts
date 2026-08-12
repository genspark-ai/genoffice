import { chainPdfium, loadPdfium, saveDoc, withDocument } from './text-edit'
import type { Pdfium } from './text-edit'
import type { AnnotDeleteInput, MarkupType } from '../shared/ipc'

/** pdfium FPDF_ANNOT_* subtype codes (fpdf_annot.h; same values pdf.js reports) */
const SUBTYPE_CODE: Record<MarkupType, number> = {
  highlight: 9,
  underline: 10,
  strikeout: 12,
}

/** Match tolerance for the rect fallback (float32 round-trips only; rects travel unchanged) */
const RECT_TOL = 2

function annotRect(m: Pdfium, annot: number): [number, number, number, number] | null {
  // FS_RECTF: four float32 {left, top, right, bottom}
  const ptr = m._malloc(16)
  try {
    if (!m._FPDFAnnot_GetRect(annot, ptr)) return null
    const f = m.HEAPF32
    return [f[ptr >> 2]!, f[(ptr >> 2) + 1]!, f[(ptr >> 2) + 2]!, f[(ptr >> 2) + 3]!]
  } finally {
    m._free(ptr)
  }
}

const rectsClose = (
  a: readonly [number, number, number, number],
  b: readonly [number, number, number, number],
): boolean => {
  // FS_RECTF is {left, top, right, bottom} while our rects are [x1,y1,x2,y2] with
  // y1 < y2 — compare the normalized bounds
  const norm = (r: readonly [number, number, number, number]) =>
    [
      Math.min(r[0], r[2]),
      Math.min(r[1], r[3]),
      Math.max(r[0], r[2]),
      Math.max(r[1], r[3]),
    ] as const
  const na = norm(a)
  const nb = norm(b)
  return Math.max(...na.map((v, i) => Math.abs(v - nb[i]!))) <= RECT_TOL
}

/** Remove by object number only when all stable identity fields still match. A full
    rewrite may reassign the number to another annotation, including one of the same subtype. */
function removeByObjNum(m: Pdfium, page: number, d: AnnotDeleteInput): boolean {
  const annot = m._EPDFPage_GetAnnotByObjectNumber(page, d.objNum)
  if (!annot) return false
  const subtypeOk = m._FPDFAnnot_GetSubtype(annot) === SUBTYPE_CODE[d.subtype]
  const rect = subtypeOk ? annotRect(m, annot) : null
  m._FPDFPage_CloseAnnot(annot)
  return (
    !!rect && rectsClose(rect, d.rect) && !!m._EPDFPage_RemoveAnnotByObjectNumber(page, d.objNum)
  )
}

/** Fallback: scan the page for an annotation with the same subtype and rect */
function removeByMatch(m: Pdfium, page: number, d: AnnotDeleteInput): boolean {
  const count = m._FPDFPage_GetAnnotCount(page)
  for (let i = 0; i < count; i++) {
    const annot = m._FPDFPage_GetAnnot(page, i)
    if (!annot) continue
    const subtypeOk = m._FPDFAnnot_GetSubtype(annot) === SUBTYPE_CODE[d.subtype]
    const rect = subtypeOk ? annotRect(m, annot) : null
    m._FPDFPage_CloseAnnot(annot)
    if (rect && rectsClose(rect, d.rect)) return !!m._FPDFPage_RemoveAnnot(page, i)
  }
  return false
}

/**
 * Remove saved markup annotations from the file bytes (in memory). Runs before every
 * other save stage so the object numbers still address the on-disk file; unmatched
 * entries are skipped fail-soft (same policy as image ops).
 */
export function applyAnnotDeletes(
  bytes: Uint8Array,
  deletes: AnnotDeleteInput[],
): Promise<Uint8Array> {
  return chainPdfium(async () => {
    const m = await loadPdfium()
    return withDocument(m, bytes, async (doc) => {
      const pageCount = m._FPDF_GetPageCount(doc)
      const byPage = new Map<number, AnnotDeleteInput[]>()
      for (const d of deletes) {
        if (d.pageIndex < 0 || d.pageIndex >= pageCount) continue
        byPage.set(d.pageIndex, [...(byPage.get(d.pageIndex) ?? []), d])
      }
      let removed = 0
      for (const [pageIndex, list] of byPage) {
        const page = m._FPDF_LoadPage(doc, pageIndex)
        if (!page) continue
        try {
          for (const d of list) {
            if (removeByObjNum(m, page, d) || removeByMatch(m, page, d)) removed++
          }
        } finally {
          m._FPDF_ClosePage(page)
        }
      }
      return removed > 0 ? saveDoc(m, doc) : bytes
    })
  })
}

/** Remove annotations on an already-loaded page for the live preview. The preview may
    run after a save rewrote the file, so it uses the same guarded lookup and fallback. */
export function removeMatchingAnnots(m: Pdfium, page: number, deletes: AnnotDeleteInput[]): void {
  for (const d of deletes) {
    if (!removeByObjNum(m, page, d)) removeByMatch(m, page, d)
  }
}
