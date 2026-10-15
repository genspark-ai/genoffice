import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { PrintJobOptions } from '../shared/ipc'

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

/** How long to wait for the dialog to report back before giving up. Chromium
    and Electron both fire `afterprint`, but a suppressed print (blocked popup,
    kiosk policy, stubbed `window.print`) never does: unbounded, the promise
    never settles, the print root keeps every rasterized data URL alive, and
    the caller's busy flag stays set so no later print is accepted. Sized to
    outlast a real dialog the user is still reading. */
export const PRINT_DIALOG_TIMEOUT_MS = 300_000

/**
 * Render scale for a document given each page's area at scale 1 (PDF points²):
 * full 200 DPI while the whole document fits the pixel budget, otherwise the
 * largest scale that fits, floored at the 150 DPI baseline.
 */
export function printScaleForAreas(areas: number[]): number {
  const target = TARGET_DPI / 72
  const total = areas.reduce((sum, area) => sum + area, 0)
  if (!(total > 0) || !Number.isFinite(total)) return target
  const atTarget = total * target * target
  if (!Number.isFinite(atTarget) || atTarget <= 0) return target
  if (atTarget <= MAX_PRINT_PIXELS) return target
  return Math.max(BASE_DPI / 72, target * Math.sqrt(MAX_PRINT_PIXELS / atTarget))
}

/**
 * Sequentially render pages as JPEG images into a print-only container (canvas discarded
 * immediately to avoid keeping full-doc hi-res bitmaps in memory), then hand off to the
 * system print dialog; clean up after it closes (including cancel).
 * Caller flushes unsaved changes and re-getDocument first — rotations/deleted pages are
 * already in the file.
 * @param pages 1-based file pages to render; defaults to the whole document.
 * The 200 DPI pixel budget is computed over the selected pages only, so a
 * small range out of a huge document still prints at full target quality.
 */
/** Settles when the print dialog reports back, when `window.print` fails
    outright, or when the wait exceeds the bound. Always detaches the listener
    and the timer, so a wedged dialog cannot leak either. */
function waitForPrintDialog(timeoutMs: number): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    let timer: ReturnType<typeof setTimeout> | null = null
    const detach = () => {
      if (timer !== null) clearTimeout(timer)
      timer = null
      window.removeEventListener('afterprint', onAfterPrint)
    }
    function onAfterPrint(): void {
      detach()
      resolve()
    }
    timer = setTimeout(() => {
      detach()
      reject(new Error('the print dialog never reported completion; released the rendered pages'))
    }, timeoutMs)
    window.addEventListener('afterprint', onAfterPrint)
    try {
      window.print()
    } catch (err) {
      detach()
      reject(err instanceof Error ? err : new Error(String(err)))
    }
  })
}

/** Keep every other page of a file-page list, Word's "odd pages / even pages" subset. */
export function filterPageParity(pages: number[], parity: PageParity): number[] {
  if (parity === 'all') return pages
  return pages.filter((n) => n % 2 === (parity === 'odd' ? 1 : 0))
}

export type PageParity = 'all' | 'odd' | 'even'

/** `job`: a chosen printer prints silently through the main process with its
    copies/duplex options; omitted, the system print dialog takes over. */
export async function printPdf(
  doc: PDFDocumentProxy,
  pages?: number[],
  job?: PrintJobOptions,
): Promise<void> {
  const root = document.createElement('div')
  root.className = 'pdf-print-root'
  const canvas = document.createElement('canvas')
  // Integer-only: a float (1.5) passes the range check but pdf.js getPage
  // throws on it, aborting the whole job in the measure pass below.
  const targets =
    pages && pages.length > 0
      ? [...new Set(pages)]
          .filter((n) => Number.isInteger(n) && n >= 1 && n <= doc.numPages)
          .sort((a, b) => a - b)
      : Array.from({ length: doc.numPages }, (_x, i) => i + 1)
  // A caller-supplied range that filters to nothing used to print an empty document:
  // both loops no-op and window.print() runs against an empty root. Refuse instead.
  if (pages && pages.length > 0 && targets.length === 0) {
    throw new Error(
      `none of the requested pages (${pages.join(', ')}) exist in this ${doc.numPages}-page document`,
    )
  }
  // First pass: measure each page at unit scale to budget the shared scale.
  // Stream one page at a time so no PDFPageProxy outlives its render — the
  // previous two-pass held all pages alive plus 2× viewports.
  const areas: number[] = []
  for (const n of targets) {
    const page = await doc.getPage(n)
    try {
      const unit = page.getViewport({ scale: 1 })
      const area = unit.width * unit.height
      areas.push(Number.isFinite(area) && area > 0 ? area : 0)
    } finally {
      try {
        page.cleanup()
      } catch {
        /* older pdf.js without cleanup */
      }
    }
  }
  const scale = printScaleForAreas(areas)
  // Second pass: render at the budgeted scale, bounding each canvas to the
  // pixel budget so a single huge drawing (e.g. 200in² at 150 DPI → >1Bpx)
  // cannot OOM the renderer even at the floor scale.
  const MAX_PAGE_PIXELS = MAX_PRINT_PIXELS
  for (const n of targets) {
    const page = await doc.getPage(n)
    try {
      let viewport = page.getViewport({ scale })
      let w = Math.floor(viewport.width)
      let h = Math.floor(viewport.height)
      if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) {
        // Skip a single corrupt page rather than aborting the whole print.
        continue
      }
      if (w * h > MAX_PAGE_PIXELS) {
        // A missing page is worse than a softer one: render this page at its
        // own reduced scale instead of dropping it.
        const pageScale = scale * Math.sqrt(MAX_PAGE_PIXELS / (w * h))
        viewport = page.getViewport({ scale: pageScale })
        w = Math.floor(viewport.width)
        h = Math.floor(viewport.height)
        if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) {
          continue
        }
      }
      canvas.width = w
      canvas.height = h
      await page.render({ canvas, viewport }).promise
      const img = document.createElement('img')
      img.src = canvas.toDataURL('image/jpeg', 0.92)
      root.appendChild(img)
    } finally {
      try {
        page.cleanup()
      } catch {
        /* older pdf.js without cleanup */
      }
    }
  }
  canvas.width = 0
  canvas.height = 0
  document.body.appendChild(root)
  try {
    await Promise.all([...root.querySelectorAll('img')].map((img) => img.decode()))
    if (job) {
      const r = await window.pdfApi.print(job)
      if (!r.ok && r.error) throw new Error(r.error)
    } else {
      await waitForPrintDialog(PRINT_DIALOG_TIMEOUT_MS)
    }
  } finally {
    root.remove()
  }
}
