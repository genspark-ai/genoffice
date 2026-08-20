/**
 * Local PDF → Word conversion for the shell's pdf tabs (pdf2docx P4).
 * Loads the shared PDFium wasm with the same lazy-singleton pattern as
 * apps/pdf/src/main/text-edit.ts and runs the pure @genoffice/pdf2docx
 * pipeline in the main process. Imported by relative path (like the other
 * sibling app modules) so the bundled shell main carries the package inline.
 */
import { readFileSync } from 'node:fs'
import { convertPdfToDocx, PdfLoadError } from '../../../../packages/pdf2docx/src'
import type { ConvertResult, PdfiumModule } from '../../../../packages/pdf2docx/src'
import { pdfiumWasmPath } from '../../../pdf/src/main/wasm-path'

export type { ConvertResult, PageResult } from '../../../../packages/pdf2docx/src'
export { PdfLoadError } from '../../../../packages/pdf2docx/src'

let pdfiumPromise: Promise<PdfiumModule> | null = null

/** Load the wasm bytes ourselves: the bundled main process must not rely on
 *  the package's own file resolution (see apps/pdf text-edit.ts). Exported so
 *  the pptx exporter (pdf2pptx-local.ts) shares the same wasm singleton. */
export function ensurePdfium(): Promise<PdfiumModule> {
  pdfiumPromise ??= (async () => {
    const { init } = (await import('@embedpdf/pdfium')) as unknown as {
      init(overrides: object): Promise<object>
    }
    const raw = readFileSync(pdfiumWasmPath())
    // exact slice: Buffer.buffer may be a shared pool larger than the file
    const wasmBinary = raw.buffer.slice(raw.byteOffset, raw.byteOffset + raw.byteLength)
    const wrapped = (await init({ wasmBinary })) as { pdfium?: unknown }
    const m = (wrapped.pdfium ?? wrapped) as PdfiumModule & { _PDFiumExt_Init(): void }
    m._PDFiumExt_Init()
    return m
  })()
  return pdfiumPromise
}

/** Convert a PDF file on disk to docx bytes, fully locally. */
export async function convertPdfFileToDocxLocal(
  pdfPath: string,
  onProgress?: (page: number, total: number) => void,
  password?: string,
): Promise<ConvertResult> {
  const pdfium = await ensurePdfium()
  const bytes = readFileSync(pdfPath)
  const pdf = new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  return convertPdfToDocx(pdf, {
    pdfium,
    ...(onProgress !== undefined ? { onProgress } : {}),
    ...(password !== undefined ? { password } : {}),
  })
}

/**
 * Password retry loop around a conversion attempt (P23). Runs `convert`
 * without a password first; on PdfLoadError('password-required') asks
 * `promptPassword` (retry=true once a submitted password was rejected) and
 * re-runs until it succeeds, a different error is thrown, or the prompt
 * returns null (user cancelled) → resolves null. Pure state machine, exported
 * separately from the UI so tests can drive it with fakes or the real
 * converter.
 */
export async function convertWithPasswordRetry<T>(
  convert: (password: string | undefined) => Promise<T>,
  promptPassword: (retry: boolean) => Promise<string | null>,
): Promise<T | null> {
  let password: string | undefined
  for (;;) {
    try {
      return await convert(password)
    } catch (err) {
      if (!(err instanceof PdfLoadError) || err.code !== 'password-required') throw err
      const entered = await promptPassword(password !== undefined)
      if (entered === null) return null
      password = entered
    }
  }
}

/**
 * Local conversion with an interactive password prompt: like
 * convertPdfFileToDocxLocal but encrypted PDFs ask the user for the password
 * (looping on wrong entries) instead of failing. Resolves null when the user
 * cancels the prompt.
 */
export function convertPdfFileToDocxLocalWithPrompt(
  pdfPath: string,
  promptPassword: (retry: boolean) => Promise<string | null>,
  onProgress?: (page: number, total: number) => void,
): Promise<ConvertResult | null> {
  return convertWithPasswordRetry(
    (password) => convertPdfFileToDocxLocal(pdfPath, onProgress, password),
    promptPassword,
  )
}
