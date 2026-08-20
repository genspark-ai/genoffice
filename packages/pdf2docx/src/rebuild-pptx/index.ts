/**
 * PDF → PPTX conversion entry (P25). Shares the extract→analyze pipeline
 * with convertPdfToDocx (pipeline.ts) and rebuilds each page as one slide of
 * the page's size, all blocks absolutely positioned.
 *
 * Furniture (repeated headers/footers/page numbers) stays dropped: slides
 * carry no header/footer text boxes; the dedup warning still surfaces.
 */
import {
  extractIrDocument,
  isScannedDocument,
  type ConvertOptions,
  type PageResult,
} from '../pipeline'
import { applyOutputFontSubstitutions } from '../rebuild/fontmap'
import { rebuildPptx } from './rebuild'

export interface ConvertPptxResult {
  pptx: Uint8Array
  pages: number
  /** human-readable notes about degraded/scanned pages (1-based page numbers) */
  warnings: string[]
  /** per-page outcomes, in page order */
  pageResults: PageResult[]
  /** most pages are scans → the caller should steer the user to an OCR flow */
  scannedDocument: boolean
}

export async function convertPdfToPptx(
  pdf: Uint8Array,
  opts: ConvertOptions,
): Promise<ConvertPptxResult> {
  const { irPages, warnings, pageResults } = extractIrDocument(pdf, opts)
  // uninstalled embedded families map to metric-compatible stand-ins (P21)
  applyOutputFontSubstitutions(irPages, [])
  const pptx = await rebuildPptx(irPages)
  return {
    pptx,
    pages: irPages.length,
    warnings,
    pageResults,
    scannedDocument: isScannedDocument(pageResults, irPages.length),
  }
}

export { rebuildPptx } from './rebuild'
export { textBlockParagraph, blockRuns } from './text'
export { tableGridOptions, type PageMapper } from './table'
