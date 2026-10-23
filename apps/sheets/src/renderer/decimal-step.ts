/**
 * Excel-parity Increase/Decrease Decimal for the ribbon buttons.
 *
 * Univer's Add/SubtractDecimalCommand derive a General-format cell's current
 * decimal count with `getDecimalFromPattern`, which asks the numfmt engine for
 * the pattern's `maxDecimals` — and `getFormatInfo('General').maxDecimals` is
 * 9. A cell with no numfmt record takes the raw-value branch (0.77 → 2), but a
 * cell carrying an explicit "General" record (what the file gateway writes for
 * numFmtId 0) reads as 9, so one Increase Decimal click jumps 341504.77 to
 * 341504.7700000000. This reimplements both commands verbatim except that an
 * explicit default-format record goes through the same raw-value branch as a
 * missing one.
 */
import { CellValueType, isDefaultFormat, Range } from '@univerjs/core'
import { INumfmtService, SheetsSelectionsService } from '@univerjs/sheets'
import { getDecimalFromPattern, setPatternDecimal } from '@univerjs/sheets-numfmt'

import type { UniverRuntime } from './univer-state'

interface NumfmtCellValue {
  row: number
  col: number
  pattern: string
}

interface RawCellLike {
  v?: unknown
  t?: unknown
}

/** Decimal digits of the raw value's shortest string form (0.77 → 2). */
export function rawValueDecimals(cell: RawCellLike | null | undefined): number {
  if (!cell || cell.t !== CellValueType.NUMBER || !cell.v) return 0
  const match = /\.(\d*)$/.exec(String(cell.v))
  return match?.[1]?.length ?? 0
}

export function stepSelectionDecimals(runtime: UniverRuntime, step: 1 | -1): void {
  const injector = runtime.univer.__getInjector()
  const selectionsService = injector.get(SheetsSelectionsService)
  const numfmtService = injector.get(INumfmtService)
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const facadeSheet = workbook?.getActiveSheet()
  if (!workbook || !facadeSheet) return
  const unitId = workbook.getId()
  const subUnitId = facadeSheet.getSheetId()
  const worksheet = facadeSheet.getSheet()
  const selections = selectionsService.getCurrentSelections()
  if (!selections || selections.length === 0) return

  // An explicit "General" record carries no decimal information; route it
  // through the raw-value branch like a missing record.
  const explicitPattern = (row: number, col: number): string | null => {
    const value = numfmtService.getValue(unitId, subUnitId, row, col)
    return value && !isDefaultFormat(value.pattern) ? value.pattern : null
  }

  let maxDecimals = 0
  for (const selection of selections) {
    Range.foreach(selection.range, (row, col) => {
      const pattern = explicitPattern(row, col)
      if (!pattern) {
        // Univer quirk kept as-is: raw values only seed the count while it is
        // still zero, and a decimal-less number leaves it untouched.
        if (!maxDecimals) {
          const decimals = rawValueDecimals(worksheet.getCellRaw(row, col) as RawCellLike | null)
          if (decimals) maxDecimals = Math.max(maxDecimals, decimals)
        }
        return
      }
      maxDecimals = Math.max(maxDecimals, getDecimalFromPattern(pattern))
    })
  }

  const decimals = maxDecimals + step
  const defaultPattern =
    step > 0
      ? setPatternDecimal(`0${decimals > 0 ? '.0' : ''}`, decimals)
      : setPatternDecimal(`0${decimals > 0 ? '.0' : '.'}`, decimals)
  const values: NumfmtCellValue[] = []
  for (const selection of selections) {
    Range.foreach(selection.range, (row, col) => {
      const pattern = explicitPattern(row, col)
      if (!pattern) {
        values.push({ row, col, pattern: defaultPattern })
        return
      }
      const next = setPatternDecimal(pattern, getDecimalFromPattern(pattern) + step)
      // Mirrors upstream: Increase skips a no-op pattern, Decrease writes it.
      if (step < 0 || next !== pattern) values.push({ row, col, pattern: next })
    })
  }
  if (values.length === 0) return
  void runtime.univerAPI.executeCommand('sheet.command.numfmt.set.numfmt', {
    unitId,
    subUnitId,
    values,
  })
}
