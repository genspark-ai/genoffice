import type { FWorkbook, FWorksheet } from '@univerjs/sheets/facade'

/** Request/reply the shell relays from `genoffice open --range` and `genoffice selection`. */
export type ControlRequest =
  { cmd: 'goto'; target: { kind: string; sheet?: string; range?: string } } | { cmd: 'selection' }

export type ControlReply =
  | { status: 'ok'; result: Record<string, unknown> }
  | { status: 'not_ready' }
  | {
      status: 'error'
      error: { reason: string; message: string; detail?: Record<string, unknown> }
    }

const A1_RANGE = /^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/
const MAX_SELECTION_VALUES = 200

/** `Sheet1!B2:D5` → { sheet: 'Sheet1', range: 'B2:D5' }; quotes around the sheet name are dropped. */
export function splitSheetRange(spec: string): { sheet?: string; range: string } | null {
  const bang = spec.lastIndexOf('!')
  const sheet = bang >= 0 ? spec.slice(0, bang).replace(/^'(.*)'$/, '$1') : undefined
  const range = (bang >= 0 ? spec.slice(bang + 1) : spec).trim().toUpperCase()
  if (!A1_RANGE.test(range)) return null
  return sheet ? { sheet, range } : { range }
}

export function handleSheetsControl(
  req: ControlRequest,
  workbook: FWorkbook | null | undefined,
  fileLoaded: boolean,
): ControlReply {
  // a pre-warmed spare view has a placeholder workbook mounted before the file is in
  if (!workbook || !fileLoaded) return { status: 'not_ready' }
  if (req.cmd === 'selection') {
    const sheet = workbook.getActiveSheet()
    try {
      sheet.getVisibleRange()
    } catch {
      return { status: 'not_ready' }
    }
    const range = workbook.getActiveRange()
    if (!range) return { status: 'ok', result: { none: true, sheet: sheet.getSheetName() } }
    const cells = range.getWidth() * range.getHeight()
    return {
      status: 'ok',
      result: {
        sheet: sheet.getSheetName(),
        range: range.getA1Notation(),
        ...(cells <= MAX_SELECTION_VALUES ? { values: range.getValues() } : { cells }),
      },
    }
  }
  const { sheet: sheetName, range: rangeSpec } = req.target
  const parsed = rangeSpec ? splitSheetRange(rangeSpec) : null
  if (!parsed) {
    return {
      status: 'error',
      error: {
        reason: 'invalid_argument',
        message: `--range must look like B2:D5 or Sheet1!B2:D5, got ${rangeSpec}`,
      },
    }
  }
  const wanted = sheetName ?? parsed.sheet
  const names = workbook.getSheets().map((s) => s.getSheetName())
  const sheet = wanted
    ? (workbook.getSheetByName(wanted) ??
      workbook.getSheets().find((s) => s.getSheetName().toLowerCase() === wanted.toLowerCase()))
    : workbook.getActiveSheet()
  if (!sheet) {
    return {
      status: 'error',
      error: {
        reason: 'sheet_not_found',
        message: `no worksheet named ${wanted}`,
        detail: { sheets: names },
      },
    }
  }
  try {
    workbook.setActiveSheet(sheet)
    // throws until the sheet's render controllers exist; before that Univer's own
    // mount-time selection would land on top of ours
    sheet.getVisibleRange()
    const range = sheet.getRange(parsed.range)
    range.activate()
    const active = workbook.getActiveRange()
    if (!active || active.getA1Notation() !== range.getA1Notation()) return { status: 'not_ready' }
    scrollIntoView(sheet, range.getRow(), range.getColumn())
    return {
      status: 'ok',
      result: { sheet: sheet.getSheetName(), range: range.getA1Notation() },
    }
  } catch {
    // Univer services are still registering while the workbook boots; the shell asks again
    return { status: 'not_ready' }
  }
}

/** Best effort: the scroll render controller is not registered while a sheet is still booting. */
function scrollIntoView(sheet: FWorksheet, row: number, column: number, attempt = 0): void {
  try {
    sheet.scrollToCell(row, column)
  } catch {
    if (attempt < 10) setTimeout(() => scrollIntoView(sheet, row, column, attempt + 1), 100)
  }
}
