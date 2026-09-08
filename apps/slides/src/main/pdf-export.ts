import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

export interface PdfExportWindow {
  loadFile(path: string): Promise<void>
  webContents: {
    executeJavaScript(script: string, userGesture?: boolean): Promise<unknown>
    printToPDF(options: Electron.PrintToPDFOptions): Promise<Buffer>
  }
  destroy(): void
}

export interface ExportSlidesPdfOptions {
  pngsBase64: string[]
  widthPx: number
  heightPx: number
  filePath: string
  createWindow(): PdfExportWindow
  openExportedPdf(path: string): void
}

export interface ExportSlidesPdfResult {
  ok: boolean
  path?: string
  error?: string
}

function buildPdfExportHtml(pngsBase64: string[], widthIn: number, heightIn: number): string {
  return `<!doctype html><html><head><meta charset="utf-8"><style>
@page { size: ${widthIn}in ${heightIn}in; margin: 0; }
html, body { margin: 0; padding: 0; }
.page { width: ${widthIn}in; height: ${heightIn}in; overflow: hidden; page-break-after: always; }
.page:last-child { page-break-after: auto; }
.page img { display: block; width: 100%; height: 100%; }
</style></head><body>${pngsBase64
    .map((b64) => `<div class="page"><img src="data:image/png;base64,${b64}"></div>`)
    .join('')}</body></html>`
}

/** Export rendered slide PNGs via an app-owned temporary HTML file. */
export async function exportSlidesPdf({
  pngsBase64,
  widthPx,
  heightPx,
  filePath,
  createWindow,
  openExportedPdf,
}: ExportSlidesPdfOptions): Promise<ExportSlidesPdfResult> {
  // PDF page size: fixed 7.5in height, width by slide ratio (16:9 -> 13.333in, 4:3 -> 10in)
  const heightIn = 7.5
  const widthIn = Math.round((widthPx / heightPx) * heightIn * 1000) / 1000
  const win = createWindow()
  let tempDir: string | null = null
  try {
    tempDir = await mkdtemp(join(tmpdir(), 'genoffice-slides-pdf-'))
    const htmlPath = join(tempDir, 'slides.html')
    await writeFile(htmlPath, buildPdfExportHtml(pngsBase64, widthIn, heightIn), 'utf8')
    await win.loadFile(htmlPath)
    // Wait for fonts and all images to decode before printing, avoiding blank pages
    await win.webContents.executeJavaScript(
      'Promise.all([document.fonts.ready, ...Array.from(document.images).map((i) => i.decode().catch(() => {}))])',
      true,
    )
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
