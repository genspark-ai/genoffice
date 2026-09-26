/**
 * OCR fallback for content indexing: scanned PDFs (no text layer) get their
 * pages rendered through the shared pdfium wasm and recognized by the platform
 * OCR helpers shared with pdf2docx (macOS Vision / Windows.Media.Ocr). Purely
 * local — page bitmaps never leave the process. Loaded lazily by the extract
 * worker only when a scanned PDF actually needs it, and only when the user
 * enabled scanned-PDF OCR in settings.
 *
 * Runs inside the extraction worker thread: the sequential page loop blocks
 * that thread by design (extraction is background work, one file at a time).
 */
import { existsSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { ensurePdfium } from '../pdf2docx-local'
import { renderPageByIndexPng, withPdfDocument } from '../../../../../packages/pdf2docx/src/extract'
import {
  createVisionOcrEngine,
  createWindowsOcrEngine,
} from '../../../../../packages/pdf2docx/src/ocr-vision'
import type { OcrEngine } from '../../../../../packages/pdf2docx/src/ocr'

/** a 500-page scan must not monopolize the extraction queue: cap the pages
 *  that get recognized (the rest of the file stays filename-only) */
const OCR_MAX_PAGES = 24
/** render scale for recognition: 2× matches the pdf2docx convert pipeline */
const OCR_RENDER_SCALE = 2

let engine: OcrEngine | null | undefined

function ensureEngine(): OcrEngine | null {
  if (engine !== undefined) return engine
  const here = dirname(fileURLToPath(import.meta.url))
  const helper = process.platform === 'darwin' ? 'vision-ocr' : 'win-ocr.exe'
  const create = process.platform === 'darwin' ? createVisionOcrEngine : createWindowsOcrEngine
  const candidates = [
    ...(process.resourcesPath ? [join(process.resourcesPath, 'ocr', helper)] : []),
    // dev/packaged-module layout: out/main/file-index → 5 levels up is the repo root
    join(here, '../../../../..', 'packages/pdf2docx/ocr-helper', helper),
  ]
  engine = null
  for (const path of candidates) {
    if (!existsSync(path)) continue
    engine = create(path)
    if (engine) break
  }
  return engine
}

/** OCR a scanned PDF into plain text (page text joined in order); null when the
 *  platform has no engine or nothing was recognized. Errors are swallowed by
 *  design — a failed OCR must never fail the indexing of the file. */
export async function ocrPdfToText(pdfPath: string): Promise<string | null> {
  const ocr = ensureEngine()
  if (!ocr) return null
  const pdfium = await ensurePdfium()
  const bytes = readFileSync(pdfPath)
  const pdf = new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const pages: string[] = []
  withPdfDocument(pdfium, pdf, (doc) => {
    const total = Math.min(pdfium._FPDF_GetPageCount(doc), OCR_MAX_PAGES)
    for (let i = 0; i < total; i++) {
      const render = renderPageByIndexPng(pdfium, doc, i, OCR_RENDER_SCALE)
      if (!render) continue
      const res = ocr(render.data, {
        widthPt: render.pixelWidth / OCR_RENDER_SCALE,
        heightPt: render.pixelHeight / OCR_RENDER_SCALE,
      })
      if (res && res.lines.length > 0) {
        pages.push(res.lines.map((l) => l.text).join(' '))
      }
    }
    return null
  })
  return pages.length > 0 ? pages.join('\n\n') : null
}
