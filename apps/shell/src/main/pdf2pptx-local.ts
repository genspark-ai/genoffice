/**
 * Local PDF → PowerPoint conversion for the shell's pdf tabs (pdf2pptx P25).
 * Mirrors pdf2docx-local.ts: same lazy PDFium wasm singleton (shared via
 * ensurePdfium), same password retry loop, pure @genoffice/pdf2docx pipeline
 * in the main process. Imported by relative path so the bundled shell main
 * carries the package inline.
 */
import { readFileSync } from 'node:fs'
import { convertPdfToPptx } from '../../../../packages/pdf2docx/src'
import type { ConvertPptxResult } from '../../../../packages/pdf2docx/src'
import { convertWithPasswordRetry, ensurePdfium } from './pdf2docx-local'

export type { ConvertPptxResult } from '../../../../packages/pdf2docx/src'

/** Convert a PDF file on disk to pptx bytes, fully locally. */
export async function convertPdfFileToPptxLocal(
  pdfPath: string,
  onProgress?: (page: number, total: number) => void,
  password?: string,
): Promise<ConvertPptxResult> {
  const pdfium = await ensurePdfium()
  const bytes = readFileSync(pdfPath)
  const pdf = new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  return convertPdfToPptx(pdf, {
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
export function convertPdfFileToPptxLocalWithPrompt(
  pdfPath: string,
  promptPassword: (retry: boolean) => Promise<string | null>,
  onProgress?: (page: number, total: number) => void,
): Promise<ConvertPptxResult | null> {
  return convertWithPasswordRetry(
    (password) => convertPdfFileToPptxLocal(pdfPath, onProgress, password),
    promptPassword,
  )
}
