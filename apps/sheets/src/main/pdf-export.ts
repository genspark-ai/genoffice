/// PDF export: renders the print HTML (laid out by the renderer) in a hidden
/// scripting-disabled window and writes webContents.printToPDF's output where
/// the save dialog points.

import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { BrowserWindow, dialog } from 'electron'

import { showSaveDialogWithMemory } from '@genoffice/electron-utils'

import { evenPageRanges, stitchPlan, type PageVariant } from './pdf-page-variants'

import type { IpcMainInvokeEvent, WebContents } from 'electron'
import type { PDFDocument } from 'pdf-lib'
import type { WorkbookExportPdfRequest, WorkbookExportPdfResult } from '../shared/desktop-api'

export async function exportPdf(
  event: IpcMainInvokeEvent,
  request: WorkbookExportPdfRequest,
): Promise<WorkbookExportPdfResult> {
  const parent = BrowserWindow.fromWebContents(event.sender)
  const dialogOptions = {
    defaultPath: request.fileName,
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
  }
  const selection = await showSaveDialogWithMemory(dialog, parent, dialogOptions)
  if (selection.canceled || !selection.filePath) return { canceled: true }

  const workDir = await mkdtemp(join(tmpdir(), 'ai-excel-pdf-'))
  const htmlPath = join(workDir, 'print.html')
  const window = new BrowserWindow({
    show: false,
    webPreferences: { sandbox: true, javascript: false },
  })
  try {
    await writeFile(htmlPath, request.html, 'utf8')
    await window.loadFile(htmlPath)
    const pdf = await renderPdf(window.webContents, request)
    await writeFile(selection.filePath, pdf)
    return { canceled: false, path: selection.filePath }
  } finally {
    window.destroy()
    await rm(workDir, { recursive: true, force: true })
  }
}

interface TemplatePair {
  readonly headerTemplate?: string | undefined
  readonly footerTemplate?: string | undefined
}

/// One printToPDF pass; `pageRanges` undefined prints every page. Chromium
/// falls back to its own date/title header when a template is missing, so
/// both templates are always passed once headers/footers are shown.
async function printPass(
  contents: WebContents,
  request: WorkbookExportPdfRequest,
  templates: TemplatePair | undefined,
  pageRanges?: string,
): Promise<Buffer> {
  return contents.printToPDF({
    landscape: request.landscape,
    pageSize: request.pageSize,
    margins: request.margins,
    scale: request.scale,
    printBackground: true,
    ...(pageRanges === undefined ? {} : { pageRanges }),
    ...(templates
      ? {
          displayHeaderFooter: true,
          headerTemplate: templates.headerTemplate ?? '<span></span>',
          footerTemplate: templates.footerTemplate ?? '<span></span>',
        }
      : {}),
  })
}

/// Chromium prints one header/footer template pair for every page. Excel's
/// differentFirst / differentOddEven need extra passes — page 1 with the
/// first-page templates, the even pages with the even ones — stitched into
/// the odd-page print by page index (pdf-lib). Without variants the odd pass
/// is the whole export (single-pass fast path).
async function renderPdf(
  contents: WebContents,
  request: WorkbookExportPdfRequest,
): Promise<Buffer> {
  const oddTemplates: TemplatePair | undefined =
    request.headerTemplate !== undefined || request.footerTemplate !== undefined
      ? { headerTemplate: request.headerTemplate, footerTemplate: request.footerTemplate }
      : undefined
  const flags = {
    hasFirst: request.firstPage !== undefined,
    hasEven: request.evenPages !== undefined,
  }
  const showHeaderFooter = oddTemplates !== undefined || flags.hasFirst || flags.hasEven
  const odd = await printPass(
    contents,
    request,
    showHeaderFooter ? (oddTemplates ?? {}) : undefined,
  )
  if (!flags.hasFirst && !flags.hasEven) return odd

  const { PDFDocument: PdfDocument } = await import('pdf-lib')
  const oddDocument = await PdfDocument.load(odd)
  const total = oddDocument.getPageCount()
  const passes: Partial<Record<PageVariant, PDFDocument>> = { odd: oddDocument }
  if (request.firstPage !== undefined && total >= 1) {
    const first = await printPass(contents, request, request.firstPage, '1')
    passes.first = await PdfDocument.load(first)
  }
  const evenRanges = evenPageRanges(total)
  if (request.evenPages !== undefined && evenRanges !== '') {
    const even = await printPass(contents, request, request.evenPages, evenRanges)
    passes.even = await PdfDocument.load(even)
  }
  const merged = await PdfDocument.create()
  for (const step of stitchPlan(total, flags)) {
    // A pass that came back with fewer pages than planned (Chromium and
    // pdf-lib disagreeing about a range) falls back to the odd print of
    // that page rather than failing the export.
    const source = passes[step.source]
    const [page] =
      source !== undefined && step.index < source.getPageCount()
        ? await merged.copyPages(source, [step.index])
        : await merged.copyPages(oddDocument, [step.page - 1])
    if (page) merged.addPage(page)
  }
  return Buffer.from(await merged.save())
}
