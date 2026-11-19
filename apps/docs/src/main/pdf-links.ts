/**
 * GoTo link annotations for exported PDFs. Chromium only emits an internal
 * link when the fragment resolves to an element with a matching `id` inside
 * the same print job; the paginated preview prints clones without ids and in
 * page chunks, so the renderer measures the links and they are stamped here
 * after the (merged) document exists. Electron-free for unit tests.
 */
import { PDFArray, PDFName, type PDFDocument } from 'pdf-lib'
import type { PdfInternalLink } from '../shared/ipc'

const MAX_LINKS = 20_000

const frac = (v: unknown): v is number =>
  typeof v === 'number' && Number.isFinite(v) && v >= 0 && v <= 1

const pageIndex = (v: unknown): v is number =>
  typeof v === 'number' && Number.isInteger(v) && v >= 0

/** renderer-supplied links filtered to well-formed, in-range entries */
export function sanitizePdfLinks(links: unknown, pageCount: number): PdfInternalLink[] {
  if (!Array.isArray(links)) return []
  const out: PdfInternalLink[] = []
  for (const l of links as Array<Record<string, unknown>>) {
    if (out.length >= MAX_LINKS) break
    if (!l || typeof l !== 'object') continue
    const { page, x, y, w, h, toPage, toY } = l
    if (!pageIndex(page) || !pageIndex(toPage) || page >= pageCount || toPage >= pageCount) continue
    if (!frac(x) || !frac(y) || !frac(w) || !frac(h) || !frac(toY)) continue
    if (w <= 0 || h <= 0 || x + w > 1.0001 || y + h > 1.0001) continue
    out.push({ page, x, y, w, h, toPage, toY })
  }
  return out
}

export function addPdfInternalLinks(pdf: PDFDocument, links: PdfInternalLink[]): number {
  const pages = pdf.getPages()
  let added = 0
  for (const l of sanitizePdfLinks(links, pages.length)) {
    const page = pages[l.page]
    const target = pages[l.toPage]
    const pw = page.getWidth()
    const ph = page.getHeight()
    const annot = pdf.context.obj({
      Type: 'Annot',
      Subtype: 'Link',
      Rect: [l.x * pw, ph - (l.y + l.h) * ph, (l.x + l.w) * pw, ph - l.y * ph],
      Border: [0, 0, 0],
      Dest: [target.ref, 'XYZ', null, target.getHeight() * (1 - l.toY), null],
    })
    let annots = page.node.lookup(PDFName.of('Annots'))
    if (!(annots instanceof PDFArray)) {
      annots = pdf.context.obj([])
      page.node.set(PDFName.of('Annots'), annots)
    }
    ;(annots as PDFArray).push(pdf.context.register(annot))
    added += 1
  }
  return added
}
