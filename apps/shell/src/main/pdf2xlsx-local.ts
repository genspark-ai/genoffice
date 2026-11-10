/**
 * Local PDF → Excel conversion for the shell's pdf tabs (pdf2xlsx P26).
 * Mirrors pdf2pptx-local.ts: same lazy PDFium wasm singleton (shared via
 * ensurePdfium), same password retry loop, pure @genoffice/pdf2docx pipeline
 * in the main process. Imported by relative path so the bundled shell main
 * carries the package inline.
 */
import { readFileSync } from 'node:fs'
import { convertPdfToXlsx } from '../../../../packages/pdf2docx/src'
import type { ConvertXlsxResult } from '../../../../packages/pdf2docx/src'
import { convertWithPasswordRetry, ensurePdfium } from './pdf2docx-local'

export type { ConvertXlsxResult } from '../../../../packages/pdf2docx/src'

/** Convert a PDF file on disk to xlsx bytes, fully locally. */
export async function convertPdfFileToXlsxLocal(
  pdfPath: string,
  onProgress?: (page: number, total: number) => void,
  password?: string,
): Promise<ConvertXlsxResult> {
  const pdfium = await ensurePdfium()
  const bytes = readFileSync(pdfPath)
  const pdf = new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  return convertPdfToXlsx(pdf, {
    pdfium,
    ...(onProgress !== undefined ? { onProgress } : {}),
    ...(password !== undefined ? { password } : {}),
  })
}

/**
 * Local conversion with an interactive password prompt: encrypted PDFs ask
 * the user for the password (looping on wrong entries) instead of failing.
 * Resolves null when the user cancels the prompt.
 */
export function convertPdfFileToXlsxLocalWithPrompt(
  pdfPath: string,
  promptPassword: (retry: boolean) => Promise<string | null>,
  onProgress?: (page: number, total: number) => void,
): Promise<ConvertXlsxResult | null> {
  return convertWithPasswordRetry(
    (password) => convertPdfFileToXlsxLocal(pdfPath, onProgress, password),
    promptPassword,
  )
}
