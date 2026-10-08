import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'

import type { WorkbookFile, WorkbookTableAdd, WorkbookTableEdit } from '../shared/desktop-api'
import { expandContiguousRegion, type TableRegion } from './table-actions'

/// Pure Table Design rules shared by the ribbon echo, the actions and tests.

type FileTable = WorkbookFile['sheets'][number]['tables'][number]

export interface TableStyleOptionFlags {
  readonly headerRow: boolean
  readonly totalsRow: boolean
  readonly bandedRows: boolean
  readonly bandedColumns: boolean
  readonly firstColumn: boolean
  readonly lastColumn: boolean
  readonly filterButton: boolean
}

export type TableOptionKey = keyof TableStyleOptionFlags

export const TABLE_OPTION_KEYS: readonly TableOptionKey[] = [
  'headerRow',
  'totalsRow',
  'bandedRows',
  'firstColumn',
  'lastColumn',
  'bandedColumns',
  'filterButton',
]

export function optionsOfTableAdd(add: WorkbookTableAdd): TableStyleOptionFlags {
  return {
    headerRow: add.headerRow ?? true,
    totalsRow: add.totalsRow ?? false,
    bandedRows: add.bandedRows,
    bandedColumns: add.bandedColumns ?? false,
    firstColumn: add.firstColumn ?? false,
    lastColumn: add.lastColumn ?? false,
    filterButton: add.filterButton ?? true,
  }
}

/// The file part only carries the band counts and stripe flags; the
/// column emphasis and filter-button flags start at Excel's defaults and
/// follow the session's pending edit afterwards.
export function optionsOfFileTable(
  table: FileTable,
  edit: WorkbookTableEdit | undefined,
): TableStyleOptionFlags {
  return {
    headerRow: table.headerRowCount > 0,
    totalsRow: (table.totalsRowCount ?? 0) > 0,
    bandedRows: table.showRowStripes,
    bandedColumns: table.showColumnStripes,
    firstColumn: edit?.firstColumn ?? false,
    lastColumn: edit?.lastColumn ?? false,
    filterButton: edit?.filterButton ?? true,
  }
}

/// Excel table names: start with a letter, underscore or backslash; letters,
/// digits, periods and underscores after; never a cell reference; 255 max.
export function isValidTableName(name: string): boolean {
  if (name.length === 0 || name.length > 255) return false
  if (!/^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(name)) return false
  return !/^(\$?[A-Za-z]{1,3}\$?[0-9]+|[rR]\d*[cC]\d*)$/.test(name)
}

export function areaToA1(area: TableRegion): string {
  return (
    `${columnLabel(area.startColumn)}${area.startRow + 1}` +
    `:${columnLabel(area.endColumn)}${area.endRow + 1}`
  )
}

export function areasIntersect(a: TableRegion, b: TableRegion): boolean {
  return (
    a.startRow <= b.endRow &&
    b.startRow <= a.endRow &&
    a.startColumn <= b.endColumn &&
    b.startColumn <= a.endColumn
  )
}

export function containsCell(area: TableRegion, row: number, column: number): boolean {
  return (
    row >= area.startRow &&
    row <= area.endRow &&
    column >= area.startColumn &&
    column <= area.endColumn
  )
}

/// Ctrl+T on a single cell takes the contiguous data block around it (Excel's
/// CurrentRegion); a multi-cell selection is used as drawn.
export function resolveCreateTableRange(
  selection: TableRegion,
  hasValue: (row: number, column: number) => boolean,
  last: { row: number; column: number },
): TableRegion {
  const single =
    selection.startRow === selection.endRow && selection.startColumn === selection.endColumn
  if (!single) return selection
  return (
    expandContiguousRegion(
      hasValue,
      { row: selection.startRow, column: selection.startColumn },
      last,
    ) ?? selection
  )
}

/// Header texts → unique column names: blanks become ColumnN, repeats take a
/// numeric suffix (case-insensitive, like Excel).
export function uniqueColumnNames(headers: readonly unknown[]): string[] {
  const names: string[] = []
  const used = new Set<string>()
  headers.forEach((raw, index) => {
    const text = String(raw ?? '')
      .trim()
      .slice(0, 255)
    const base = text.length === 0 ? `Column${index + 1}` : text
    let candidate = base
    for (let suffix = 2; used.has(candidate.toLowerCase()); suffix += 1) {
      candidate = `${base}${suffix}`
    }
    used.add(candidate.toLowerCase())
    names.push(candidate)
  })
  return names
}

export type ResizeProblem = 'header-row-moved' | 'no-data-rows' | 'no-overlap'

/// Excel's Resize Table rules: the header row stays put, the new range must
/// overlap the old one and keep at least one data row beside the bands.
export function resizeProblem(
  current: TableRegion,
  next: TableRegion,
  options: Pick<TableStyleOptionFlags, 'headerRow' | 'totalsRow'>,
): ResizeProblem | null {
  if (options.headerRow && next.startRow !== current.startRow) return 'header-row-moved'
  if (!areasIntersect(current, next)) return 'no-overlap'
  const bandRows = (options.headerRow ? 1 : 0) + (options.totalsRow ? 1 : 0)
  if (next.endRow - next.startRow + 1 <= bandRows) return 'no-data-rows'
  return null
}

/// SUBTOTAL function Excel picks for a new totals row: 109 (SUM) over a
/// numeric column, 103 (COUNTA) otherwise.
export function totalsFunctionFor(values: readonly unknown[]): 109 | 103 {
  const filled = values.filter((value) => value !== null && value !== undefined && value !== '')
  if (filled.length === 0) return 103
  return filled.every((value) => typeof value === 'number') ? 109 : 103
}
