import { CellValueType, type ICellData } from '@univerjs/core'

import {
  FILL_SERIES_OPEN_ENDED_LIMIT,
  type FillSeriesDirection,
  type FillSeriesOptions,
  planFillSeries,
} from './fill-series'
import type { TFunc } from './i18n/locale'
import type { UniverRuntime } from './univer-state'

/// Writes the series into the active selection; each write is a normal
/// set-range-values edit, so it journals and undoes like typing.
export function applyFillSeries(
  runtime: UniverRuntime,
  options: FillSeriesOptions,
  t: TFunc,
): string | null {
  // Open-ended runs (single cell + stop) may need rows/columns past the
  // current grid; the sheet grows to fit, like Excel's unbounded grid.
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  const active = workbook?.getActiveRange()
  if (!sheet || !active) return t('appSelectRangeFirst')
  const { startRow, startColumn } = active.getRange()
  const cells = active.getCellDatas()
  const computed = active.getValues()
  // Seeds: raw numbers, numeric text typed as NUMBER, or a formula's result.
  const grid = cells.map((row, r) =>
    row.map((cell, c) => {
      const v = cell?.v
      if (typeof v === 'number') return v
      if (typeof v === 'string' && cell?.t === CellValueType.NUMBER && v.trim() !== '') {
        const numeric = Number(v)
        if (Number.isFinite(numeric)) return numeric
      }
      if (cell?.f !== undefined || cell?.si !== undefined) {
        const result = computed[r]?.[c]
        if (typeof result === 'number') return result
      }
      return v ?? null
    }),
  )
  const writes = planFillSeries(grid, options, FILL_SERIES_OPEN_ENDED_LIMIT)
  if (writes.length === 0) return t('appFillSeriesNothing')
  const needRows = Math.max(...writes.map((write) => startRow + write.row + write.values.length))
  const needColumns = Math.max(
    ...writes.map((write) => startColumn + write.column + (write.values[0]?.length ?? 0)),
  )
  if (needRows > sheet.getMaxRows())
    sheet.insertRows(sheet.getMaxRows(), needRows - sheet.getMaxRows())
  if (needColumns > sheet.getMaxColumns()) {
    sheet.insertColumns(sheet.getMaxColumns(), needColumns - sheet.getMaxColumns())
  }
  for (const write of writes) {
    // Filled cells take their line's seed format so a date series keeps
    // showing dates.
    const seedStyle =
      options.direction === 'rows' ? cells[write.row]?.[0]?.s : cells[0]?.[write.column]?.s
    // Excel's Series replaces the target cells outright: setValues merges,
    // so formulas and rich text must be cleared explicitly.
    const matrix = write.values.map((row) =>
      row.map((v): ICellData => ({
        v,
        t: CellValueType.NUMBER,
        f: null,
        si: null,
        p: null,
        ...(seedStyle ? { s: seedStyle } : {}),
      })),
    )
    sheet
      .getRange(
        startRow + write.row,
        startColumn + write.column,
        matrix.length,
        matrix[0]?.length ?? 1,
      )
      .setValues(matrix)
  }
  return null
}

export interface FillSeriesContext {
  readonly direction: FillSeriesDirection
  readonly date1904: boolean
}

/// Dialog defaults: Series in follows the selection shape (taller than wide
/// = Columns); the date system decides how calendar stop text is read.
export function fillSeriesContext(
  runtime: UniverRuntime | null,
  date1904: boolean,
): FillSeriesContext {
  const active = runtime?.univerAPI.getActiveWorkbook()?.getActiveRange()
  return {
    direction: active && active.getHeight() > active.getWidth() ? 'columns' : 'rows',
    date1904,
  }
}
