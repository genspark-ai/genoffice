import type { PDFDocumentProxy, PDFPageProxy } from 'pdfjs-dist'

/** Baseline raster density (previous behavior — quality never drops below this). */
const BASE_DPI = 150
/** Target raster density for sharper glyph edges on high-DPI printers. */
const TARGET_DPI = 200
/**
 * Upper bound on total raster pixels across all pages. A 200 DPI US-Letter
 * page is ~3.7 MP, so ~40 such pages print at full target quality; larger
 * documents scale down toward (never below) the 150 DPI baseline. This keeps
 * the JPEG data URLs held alive until the dialog closes within ~tens of MB
 * instead of pushing the renderer out of memory on long scanned PDFs.
 */
const MAX_PRINT_PIXELS = 150_000_000

/**
 * Render scale for a document given each page's area at scale 1 (PDF points²):
 * full 200 DPI while the whole document fits the pixel budget, otherwise the
 * largest scale that fits, floored at the 150 DPI baseline.
 */
export function printScaleForAreas(areas: number[]): number {
  const target = TARGET_DPI / 72
  const total = areas.reduce((sum, area) => sum + area, 0)
  if (!(total > 0)) return target
  const atTarget = total * target * target
  if (atTarget <= MAX_PRINT_PIXELS) return target
  return Math.max(BASE_DPI / 72, target * Math.sqrt(MAX_PRINT_PIXELS / atTarget))
}

/**
 * Sequentially render pages as JPEG images into a print-only container (canvas discarded
 * immediately to avoid keeping full-doc hi-res bitmaps in memory), then hand off to the
 * system print dialog; clean up after it closes (including cancel).
 * Caller flushes unsaved changes and re-getDocument first — rotations/deleted pages are
 * already in the file.
 */
export async function printPdf(doc: PDFDocumentProxy): Promise<void> {
  const root = document.createElement('div')
  root.className = 'pdf-print-root'
  const canvas = document.createElement('canvas')
  const pages: PDFPageProxy[] = []
  const areas: number[] = []
  for (let n = 1; n <= doc.numPages; n++) {
    const page = await doc.getPage(n)
    pages.push(page)
    const unit = page.getViewport({ scale: 1 })
    areas.push(unit.width * unit.height)
  }
  const scale = printScaleForAreas(areas)
  for (const page of pages) {
    const viewport = page.getViewport({ scale })
    canvas.width = Math.floor(viewport.width)
    canvas.height = Math.floor(viewport.height)
    await page.render({ canvas, viewport }).promise
    const img = document.createElement('img')
    img.src = canvas.toDataURL('image/jpeg', 0.92)
    root.appendChild(img)
  }
  canvas.width = 0
  canvas.height = 0
  document.body.appendChild(root)
  try {
    await Promise.all([...root.querySelectorAll('img')].map((img) => img.decode()))
    await new Promise<void>((resolve) => {
      const done = () => {
        window.removeEventListener('afterprint', done)
        resolve()
      }
      window.addEventListener('afterprint', done)
      window.print()
    })
  } finally {
    root.remove()
  }
}
