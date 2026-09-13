import { readFileSync } from 'node:fs'
import {
  convertPdfToDocx,
  convertPdfToPptx,
  convertPdfToXlsx,
  extract,
  PdfLoadError,
  type OcrEngine,
  type PdfiumModule,
} from '@genoffice/pdf2docx'
import { createVisionOcrEngine, createWindowsOcrEngine } from '../../../pdf2docx/src/ocr-vision'
import { ocrHelperPath, pdfiumWasmPath } from '../resources'

let pdfiumPromise: Promise<PdfiumModule> | null = null

/** Same bootstrap as the app (apps/shell/src/main/pdf2docx-local.ts): wasm bytes read by us, thisProgram pinned so CJK argv cannot abort emscripten init. */
export function loadPdfium(): Promise<PdfiumModule> {
  pdfiumPromise ??= (async () => {
    const { init } = (await import('@embedpdf/pdfium')) as unknown as {
      init(overrides: object): Promise<object>
    }
    const raw = readFileSync(pdfiumWasmPath())
    const wasmBinary = raw.buffer.slice(raw.byteOffset, raw.byteOffset + raw.byteLength)
    const wrapped = (await init({ wasmBinary, thisProgram: 'genoffice' })) as { pdfium?: unknown }
    const m = (wrapped.pdfium ?? wrapped) as PdfiumModule & { _PDFiumExt_Init(): void }
    m._PDFiumExt_Init()
    return m
  })()
  return pdfiumPromise
}

let ocrEngine: OcrEngine | null | undefined
function ocr(): OcrEngine | null {
  if (ocrEngine !== undefined) return ocrEngine
  const helper = ocrHelperPath()
  const create = process.platform === 'darwin' ? createVisionOcrEngine : createWindowsOcrEngine
  ocrEngine = helper ? create(helper) : null
  return ocrEngine
}

export interface PdfInfo {
  pages: number | null
  encrypted: boolean
  producer?: string
  creator?: string
}

/** Opens without a password first, so `encrypted` reflects the document rather than the caller's flags. */
export async function pdfInfo(bytes: Uint8Array, password?: string): Promise<PdfInfo> {
  const m = await loadPdfium()
  const read = (pw: string | undefined, encrypted: boolean): PdfInfo =>
    extract.withPdfDocument(
      m,
      bytes,
      (doc) => ({
        pages: m._FPDF_GetPageCount(doc),
        encrypted,
        ...extract.readDocMetadata(m, doc),
      }),
      pw,
    )
  try {
    return read(undefined, false)
  } catch (err) {
    if (!(err instanceof PdfLoadError) || err.code !== 'password-required') throw err
    if (password === undefined) return { pages: null, encrypted: true }
    return read(password, true)
  }
}

export type PdfTarget = 'docx' | 'pptx' | 'xlsx'

export interface PdfConvertOutcome {
  bytes: Uint8Array
  pages: number
  warnings: string[]
}

export async function convertPdf(
  bytes: Uint8Array,
  target: PdfTarget,
  opts: { password?: string; onProgress?: (page: number, total: number) => void },
): Promise<PdfConvertOutcome> {
  const pdfium = await loadPdfium()
  const engine = ocr()
  const convertOpts = {
    pdfium,
    ...(engine ? { ocr: engine } : {}),
    ...(opts.onProgress ? { onProgress: opts.onProgress } : {}),
    ...(opts.password !== undefined ? { password: opts.password } : {}),
  }
  switch (target) {
    case 'docx': {
      const r = await convertPdfToDocx(bytes, convertOpts)
      return { bytes: r.docx, pages: r.pages, warnings: r.warnings }
    }
    case 'pptx': {
      const r = await convertPdfToPptx(bytes, convertOpts)
      return { bytes: r.pptx, pages: r.pages, warnings: r.warnings }
    }
    case 'xlsx': {
      const r = await convertPdfToXlsx(bytes, convertOpts)
      return { bytes: r.xlsx, pages: r.pages, warnings: r.warnings }
    }
  }
}
