/// CSV export of the active sheet: display strings from the live Univer model
/// (number formats applied, formula results, values only), Excel-style
/// quoting, CRLF rows. The main process runs the loss warning + save dialog
/// and writes UTF-8 with a BOM so Excel reopens it correctly.

import { isDefaultFormat, numfmt } from '@univerjs/core'
import { INumfmtService } from '@univerjs/sheets'

import { ConditionalFormattingService } from '@univerjs/preset-sheets-conditional-formatting'

import { MAX_CSV_EXPORT_CHARS } from '../shared/ipc-channels'
import { formulaViewSheets } from './formula-view'
import { t } from './i18n/locale'
import { DATE_1904_OFFSET, isCalendarDatePattern, patternType } from './numfmt-fix'
import { showToast } from './toast-bus'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/// Rows fetched per getDisplayValues call, bounding the facade's allocation.
const EXPORT_ROW_BLOCK = 4096

/** The App refs/state the CSV export needs; built fresh per call. */
export interface CsvExportContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  setMessage: (message: string) => void
  /// Routes into the regular Save As flow when the user picks "Save as .xlsx"
  /// in the formula-loss warning.
  requestSaveAs: () => void
}

/// The slice of the Univer facade the export reads (structural, so the
/// caller passes the FWorksheet through a cast — same pattern as print-html).
export interface CsvWorksheet {
  getLastRow(): number
  getLastColumn(): number
  getSheetId(): string
  getSheetName(): string
  getRange(
    row: number,
    column: number,
    numRows: number,
    numColumns: number,
  ): { getDisplayValues(): string[][]; getRawValues(): unknown[][] }
  getSheet(): {
    getCellMatrix(): {
      forValue(
        callback: (row: number, column: number, cell: { f?: string | null } | null) => unknown,
      ): void
    }
  }
}

export function csvField(text: string): string {
  const normalized = text.replace(/\r\n|\r/g, '\n')
  return /[",\n]/.test(normalized) ? `"${normalized.replace(/"/g, '""')}"` : normalized
}

export function csvFromDisplayRows(rows: readonly (readonly string[])[]): string {
  return rows.map((row) => row.map(csvField).join(',') + '\r\n').join('')
}

/// Runs `read` with the sheet dropped from the formula-view set, restoring it
/// afterwards: CSV always exports calculated values, and in formula view the
/// display interceptor would hand getDisplayValues the formula text instead.
/// The read is synchronous, so the grid never repaints while the flag is off.
export function withoutFormulaView<T>(sheets: Set<string>, sheetId: string, read: () => T): T {
  const hadFormulaView = sheets.delete(sheetId)
  try {
    return read()
  } finally {
    if (hadFormulaView) sheets.add(sheetId)
  }
}

/// The non-General format a numeric cell renders through: the pattern (a
/// conditional-formatting rule's pattern wins over the cell's own, as on
/// screen) and the serial numfmt should format — shifted for file-loaded
/// calendar dates in a 1904 workbook, null when the hash display is genuine
/// (a negative serial has no 1900 date at any width). Null for General.
export interface CellNumberFormat {
  pattern: string
  renderValue: number | null
}
export type CellPatternLookup = (
  row: number,
  column: number,
  raw: number,
) => CellNumberFormat | null

export function cellPatternLookup(
  runtime: UniverRuntime,
  sheetId: string,
  date1904 = false,
): CellPatternLookup | undefined {
  try {
    const injector = runtime.univer.__getInjector()
    const numfmtService = injector.get(INumfmtService) as {
      getValue(
        unitId: string,
        subUnitId: string,
        row: number,
        column: number,
      ): { pattern?: string } | null
    }
    let cfService: Pick<ConditionalFormattingService, 'composeStyle'> | null = null
    try {
      cfService = injector.get(ConditionalFormattingService)
    } catch {
      cfService = null
    }
    const workbook = runtime.univerAPI.getActiveWorkbook()
    const unitId = workbook?.getId() ?? ''
    const worksheet = workbook?.getSheetBySheetId(sheetId)?.getSheet()
    return (row, column, raw) => {
      const cfPattern = cfService?.composeStyle(unitId, sheetId, row, column)?.style?.n?.pattern
      const pattern =
        cfPattern !== undefined && !isDefaultFormat(cfPattern)
          ? cfPattern
          : (numfmtService.getValue(unitId, sheetId, row, column)?.pattern ?? null)
      if (pattern === null || isDefaultFormat(pattern)) return null
      const type = patternType(pattern)
      const isDate = type === 'date' || type === 'datetime' || type === 'time'
      if (raw < 0 && !date1904 && isDate) return { pattern, renderValue: null }
      const cell = date1904 ? worksheet?.getCellRaw(row, column) : null
      const staticSerial =
        date1904 && isCalendarDatePattern(pattern) && cell?.f == null && cell?.si == null
      return { pattern, renderValue: staticSerial ? raw + DATE_1904_OFFSET : raw }
    }
  } catch {
    // No injector (torn-down runtime): keep the display strings as-is.
    return undefined
  }
}

/// A display string Univer substituted for a value that does not fit the
/// column ("########") — never a real cell text worth exporting.
const isHashOverflow = (text: string): boolean => text.length > 0 && /^#+$/.test(text)

/// Serializes the active sheet's display grid; 'too-large' when the text
/// exceeds the IPC cap. Callers gate on preloadComplete first.
///
/// `patternOf` guards against width-fitted displays. A 645577.79 in a narrow
/// General column renders (and used to export) as "645577.8", and a formatted
/// number that does not fit at all renders as "########" — both are silent
/// data corruption in the written file. General cells emit the full model
/// value; explicitly formatted cells keep their formatted text (Excel's
/// "values as displayed"), re-rendered without the width limit when the
/// display is the hash overflow.
export function serializeActiveSheetCsv(
  sheet: CsvWorksheet,
  state: LazyWorkbookState | null,
  patternOf?: CellPatternLookup,
): string | 'too-large' {
  const rowCount = Math.max(sheet.getLastRow(), 0) + 1
  const columnCount = Math.max(sheet.getLastColumn(), 0) + 1
  const parts: string[] = []
  let length = 0
  const tooLarge = withoutFormulaView(formulaViewSheets(state), sheet.getSheetId(), () => {
    for (let start = 0; start < rowCount; start += EXPORT_ROW_BLOCK) {
      const blockRows = Math.min(EXPORT_ROW_BLOCK, rowCount - start)
      const range = sheet.getRange(start, 0, blockRows, columnCount)
      const displays = range.getDisplayValues()
      if (patternOf) {
        const raws = range.getRawValues()
        for (let rowOffset = 0; rowOffset < displays.length; rowOffset += 1) {
          const displayRow = displays[rowOffset]
          if (!displayRow) continue
          for (let column = 0; column < displayRow.length; column += 1) {
            const raw = raws[rowOffset]?.[column]
            if (typeof raw !== 'number' || !Number.isFinite(raw)) continue
            const display = displayRow[column] ?? ''
            const full = String(raw)
            if (display === full) continue
            const format = patternOf(start + rowOffset, column, raw)
            if (format === null) displayRow[column] = full
            else if (isHashOverflow(display) && format.renderValue !== null) {
              try {
                displayRow[column] = numfmt.format(format.pattern, format.renderValue)
              } catch {
                displayRow[column] = full
              }
            }
          }
        }
      }
      const part = csvFromDisplayRows(displays)
      length += part.length
      if (length > MAX_CSV_EXPORT_CHARS) return true
      parts.push(part)
    }
    return false
  })
  return tooLarge ? 'too-large' : parts.join('')
}

/// The active FWorksheet through the structural CsvWorksheet slice, plus the
/// live sheet count — shared by the menu export and the save flows.
export function activeCsvSheet(
  runtime: UniverRuntime | null,
): { sheet: CsvWorksheet; sheetCount: number } | null {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!workbook || !worksheet) return null
  return {
    sheet: worksheet as unknown as CsvWorksheet,
    sheetCount: workbook.getSheets().length,
  }
}

/// A worksheet by id through the CsvWorksheet slice (the active sheet when
/// omitted) — the AI create_document path's analog of activeCsvSheet.
export function csvSheetById(runtime: UniverRuntime | null, sheetId?: string): CsvWorksheet | null {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  const worksheet =
    sheetId === undefined ? workbook?.getActiveSheet() : workbook?.getSheetBySheetId(sheetId)
  return worksheet ? (worksheet as unknown as CsvWorksheet) : null
}

export function sheetHasFormulas(sheet: CsvWorksheet): boolean {
  let found = false
  sheet
    .getSheet()
    .getCellMatrix()
    .forValue((_row, _column, cell) => {
      if (!cell?.f) return undefined
      found = true
      return false
    })
  return found
}

export async function handleExportCsv(ctx: CsvExportContext, targetPath?: string): Promise<void> {
  const runtime = ctx.univerRef.current
  const active = activeCsvSheet(runtime)
  if (!runtime || !active) return
  const { sheet, sheetCount } = active
  const state = ctx.lazyWorkbookRef.current
  if (state && !state.flags.preloadComplete) {
    ctx.setMessage(t('appCsvExportNeedsFullLoad'))
    return
  }
  try {
    const content = serializeActiveSheetCsv(
      sheet,
      state,
      cellPatternLookup(runtime, sheet.getSheetId()),
    )
    if (content === 'too-large') {
      ctx.setMessage(t('appCsvExportTooLarge'))
      return
    }
    const baseName = (state?.file.name ?? 'Book1').replace(/\.[^.]+$/, '')
    const result = await window.desktopApi.exportCsv({
      fileName: `${baseName}.csv`,
      content,
      hasFormulas: sheetHasFormulas(sheet),
      ...(sheetCount > 1 ? { activeSheetName: sheet.getSheetName() } : {}),
      ...(targetPath === undefined ? {} : { targetPath }),
    })
    if (result.canceled) {
      if (result.saveAsXlsxInstead) ctx.requestSaveAs()
      else ctx.setMessage(t('appCsvExportCanceled'))
      return
    }
    // Toast like the regular save paths: the status bar alone is easy to
    // miss, and a CSV Save As keeps the window on the source workbook — with
    // no prominent confirmation the export reads as "nothing was saved".
    const exported =
      sheetCount > 1
        ? t('appCsvExportedActiveOnly', { name: sheet.getSheetName(), path: result.path })
        : t('appCsvExported', { path: result.path })
    ctx.setMessage(exported)
    showToast(exported)
  } catch (error: unknown) {
    const failed = error instanceof Error ? error.message : t('appCsvExportFailed')
    ctx.setMessage(failed)
    showToast(failed, 'error')
  }
}
