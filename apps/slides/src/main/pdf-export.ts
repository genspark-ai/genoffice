import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { ExportPdfLink, ExportPdfPage } from '../shared/ipc'

export interface PdfExportWindow {
  loadFile(path: string): Promise<void>
  webContents: {
    executeJavaScript(script: string, userGesture?: boolean): Promise<unknown>
    printToPDF(options: Electron.PrintToPDFOptions): Promise<Buffer>
  }
  destroy(): void
}

export interface ExportSlidesPdfOptions {
  pages: ExportPdfPage[]
  widthPx: number
  heightPx: number
  filePath: string
  /** Per-page clickable link overlays (fractions of the page box), same order as pages */
  links?: ExportPdfLink[][]
  /** @font-face rules for the SVG pages' text (data: URLs) */
  fontCss?: string
  createWindow(): PdfExportWindow
  openExportedPdf(path: string): void
}

export interface ExportSlidesPdfResult {
  ok: boolean
  path?: string
  error?: string
}

function escapeAttr(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

/** Fraction (0..1) → CSS percentage; NaN/Infinity would corrupt the markup, so reject the rect. */
function pct(v: number): string | null {
  return Number.isFinite(v) ? `${Math.round(v * 100000) / 1000}%` : null
}

/**
 * Transparent <a> boxes over the page: printToPDF converts them to PDF
 * link annotations (URI actions for URLs, in-document destinations for the
 * "#pgN" page anchors — pages carry matching ids), keeping element and text
 * hyperlinks clickable in the exported PDF like a PowerPoint export.
 */
export function buildPdfLinkOverlays(links: ExportPdfLink[] | undefined): string {
  if (!links?.length) return ''
  return links
    .map((l) => {
      const [x, y, w, h] = [pct(l.x), pct(l.y), pct(l.w), pct(l.h)]
      if (x == null || y == null || w == null || h == null || !l.href) return ''
      return `<a href="${escapeAttr(l.href)}" style="left:${x};top:${y};width:${w};height:${h}"></a>`
    })
    .join('')
}

/** A `</style>` inside the font CSS would end the block early; it cannot occur in valid rules. */
function safeCss(css: string | undefined): string {
  return css ? css.replace(/<\/style/gi, '') : ''
}

function pageMarkup(page: ExportPdfPage): string {
  return 'svg' in page ? page.svg : `<img src="data:image/png;base64,${page.png}">`
}

export function buildPdfExportHtml(
  pages: ExportPdfPage[],
  widthIn: number,
  heightIn: number,
  links?: ExportPdfLink[][],
  fontCss?: string,
): string {
  return `<!doctype html><html><head><meta charset="utf-8"><style>
${safeCss(fontCss)}
@page { size: ${widthIn}in ${heightIn}in; margin: 0; }
html, body { margin: 0; padding: 0; }
.page { position: relative; width: ${widthIn}in; height: ${heightIn}in; overflow: hidden; page-break-after: always; }
.page:last-child { page-break-after: auto; }
.page img, .page > svg { display: block; width: 100%; height: 100%; }
.page a { position: absolute; display: block; }
</style></head><body>${pages
    .map(
      (page, i) =>
        `<div class="page" id="pg${i + 1}">${pageMarkup(page)}${buildPdfLinkOverlays(links?.[i])}</div>`,
    )
    .join('')}</body></html>`
}

/**
 * Fonts and every bitmap decoded before printing, or pages print blank. SVG
 * `<image>` elements are not in document.images: decoding the same data URL
 * through an Image primes the shared cache, and two frames let the SVG pick
 * it up.
 */
export const PRINT_READY_SCRIPT = `Promise.all([
  document.fonts.ready,
  ...Array.from(document.images).map((i) => i.decode().catch(() => {})),
  ...Array.from(document.querySelectorAll('svg image')).map((el) => {
    const href = el.getAttribute('href') || el.getAttribute('xlink:href')
    if (!href) return Promise.resolve()
    const img = new Image()
    img.src = href
    return img.decode().catch(() => {})
  }),
]).then(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))))`

/** PDF page size: fixed 7.5in height, width by slide ratio (16:9 -> 13.333in, 4:3 -> 10in).
    Rendered dimensions can be zeroed by a failed capture or non-finite from a
    degenerate transform; dividing them raw yields NaN/Infinity/0in pages, so
    the ratio is normalized to the same bounds as the print path. */
export const PDF_EXPORT_HEIGHT_IN = 7.5

export function exportPageWidthIn(widthPx: number, heightPx: number): number {
  const ratio =
    Number.isFinite(widthPx) && Number.isFinite(heightPx) && heightPx > 0
      ? widthPx / heightPx
      : 16 / 9
  const safe = Number.isFinite(ratio) && ratio > 0 ? Math.min(Math.max(ratio, 0.2), 5) : 16 / 9
  return Math.round(safe * PDF_EXPORT_HEIGHT_IN * 1000) / 1000
}

/** Export rendered slide pages via an app-owned temporary HTML file. */
export async function exportSlidesPdf({
  pages,
  widthPx,
  heightPx,
  filePath,
  links,
  fontCss,
  createWindow,
  openExportedPdf,
}: ExportSlidesPdfOptions): Promise<ExportSlidesPdfResult> {
  // PDF page size: fixed 7.5in height, width by slide ratio (16:9 -> 13.333in, 4:3 -> 10in)
  const heightIn = PDF_EXPORT_HEIGHT_IN
  const widthIn = exportPageWidthIn(widthPx, heightPx)
  const win = createWindow()
  let tempDir: string | null = null
  try {
    tempDir = await mkdtemp(join(tmpdir(), 'genoffice-slides-pdf-'))
    const htmlPath = join(tempDir, 'slides.html')
    await writeFile(htmlPath, buildPdfExportHtml(pages, widthIn, heightIn, links, fontCss), 'utf8')
    await win.loadFile(htmlPath)
    await win.webContents.executeJavaScript(PRINT_READY_SCRIPT, true)
    const pdf = await win.webContents.printToPDF({
      landscape: false, // The page size is already landscape (width > height); passing landscape would rotate a second time
      printBackground: true,
      pageSize: { width: widthIn, height: heightIn },
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      preferCSSPageSize: false,
    })
    await writeFile(filePath, pdf)
    openExportedPdf(filePath)
    return { ok: true, path: filePath }
  } catch (err) {
    return { ok: false, error: String(err) }
  } finally {
    try {
      win.destroy()
    } finally {
      if (tempDir) await rm(tempDir, { recursive: true, force: true })
    }
  }
}
