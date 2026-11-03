import { openXlsContainer } from './cfb'
import { parseXlsWorkbook } from './workbook'
import { writeXlsxWorkbook } from './xlsx-writer'

export { LegacyXlsError, LEGACY_XLS_ERROR_CODES, type LegacyXlsErrorCode } from './errors'
export { openXlsContainer, type XlsContainer } from './cfb'
export type { XlsCell, XlsSheet, XlsWorkbook } from './workbook'
export { parseXlsWorkbook } from './workbook'

export interface XlsConversionResult {
  /** the .xlsx bytes, ready to write atomically like any other workbook */
  bytes: Uint8Array
  sheetCount: number
  /** what stage 1 did not carry over, so callers (and later stages) can say so */
  warnings: string[]
}

/**
 * Legacy .xls → .xlsx, stage 1: cell values, text, booleans, errors, merged
 * ranges and column widths. Formulas land as their cached values (the BIFF8
 * RPN translation is a later stage), number formats stay unformatted, and
 * charts, images and styles are not carried. The JS/headless counterpart to the
 * xlsx-engine sidecar the Sheets app uses; the CLI and MCP processes are what
 * wire this package up (the Sheets open flow keeps using the sidecar).
 */
export async function convertXlsToXlsx(bytes: Uint8Array): Promise<XlsConversionResult> {
  const container = openXlsContainer(bytes)
  const workbook = parseXlsWorkbook(container.workbook)
  const out = await writeXlsxWorkbook(workbook)
  const warnings: string[] = []
  warnings.push('formulas are carried as their cached values')
  return { bytes: out, sheetCount: workbook.sheets.length, warnings }
}
