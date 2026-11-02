import { openXlsContainer } from './cfb'
import { parseXlsWorkbook } from './workbook'
import { writeXlsxWorkbook } from './xlsx-writer'

export { LegacyXlsError, LEGACY_XLS_ERROR_CODES, type LegacyXlsErrorCode } from './errors'
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
 * charts, images and styles are not carried. The foundation the CLI and the
 * Sheets open flow wire up in the following stages.
 */
export async function convertXlsToXlsx(bytes: Uint8Array): Promise<XlsConversionResult> {
  const container = openXlsContainer(bytes)
  const workbook = parseXlsWorkbook(container.workbook)
  const out = await writeXlsxWorkbook(workbook)
  const warnings: string[] = []
  if (workbook.date1904) {
    warnings.push('date system: serial numbers use the 1904 epoch and are carried as numbers')
  }
  warnings.push('formulas are carried as their cached values')
  return { bytes: out, sheetCount: workbook.sheets.length, warnings }
}
