/**
 * PDF → XLSX conversion entry (P26). Shares the extract→analyze pipeline
 * with convertPdfToDocx/convertPdfToPptx (pipeline.ts) and rebuilds each page
 * as one worksheet: tables as cell grids, non-table text as column-A rows.
 *
 * Furniture (repeated headers/footers/page numbers) stays dropped, same as
 * the other exporters; the dedup warning still surfaces.
 */
import {
  extractIrDocument,
  isScannedDocument,
  type ConvertOptions,
  type PageResult,
} from '../pipeline'
import { applyOutputFontSubstitutions } from '../rebuild/fontmap'
import { rebuildXlsx } from './rebuild'

export interface ConvertXlsxResult {
  xlsx: Uint8Array
  pages: number
  /** human-readable notes about degraded/scanned pages (1-based page numbers) */
  warnings: string[]
  /** per-page outcomes, in page order */
  pageResults: PageResult[]
  /** most pages are scans → the caller should steer the user to an OCR flow */
  scannedDocument: boolean
}

export async function convertPdfToXlsx(
  pdf: Uint8Array,
  opts: ConvertOptions,
): Promise<ConvertXlsxResult> {
  const { irPages, warnings, pageResults } = extractIrDocument(pdf, opts)
  // uninstalled embedded families map to metric-compatible stand-ins (P21)
  applyOutputFontSubstitutions(irPages, [])
  const { xlsx, warnings: xlsxWarnings } = await rebuildXlsx(irPages)
  return {
    xlsx,
    pages: irPages.length,
    warnings: [...warnings, ...xlsxWarnings],
    pageResults,
    scannedDocument: isScannedDocument(pageResults, irPages.length),
  }
}

export { rebuildXlsx } from './rebuild'
export { parseCellValue, type ParsedCell } from './numbers'
export { ptToColumnChars, flattenBlockText } from './rebuild'
