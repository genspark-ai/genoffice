import type { WorkbookFile } from '../shared/desktop-api'

type FileSheet = WorkbookFile['sheets'][number]
type FileTable = FileSheet['tables'][number]

export interface FileTableRegistration {
  readonly sheetId: string
  readonly tableIndex: number
  readonly tableId: string
  /// The structured-reference token the formula engine resolves.
  readonly tableName: string
  /// Table range without the totals band, so `Table[Column]` sums data only.
  readonly range: FileTable['range']
  readonly columns: readonly string[] | undefined
}

const INVALID_NAME_CHAR = /[^\p{L}\p{N}_.]/gu
const CELL_LIKE = /^(\$?[A-Za-z]{1,3}\$?[0-9]+|[rR]\d+[cC]\d+|\d+)$/

/// Excel table-name rules (letter/underscore start, no spaces, not a cell
/// address, 255 chars) applied to a file displayName so Univer accepts it.
export function sanitizeTableName(raw: string): string {
  let name = raw.trim().replace(INVALID_NAME_CHAR, '_').slice(0, 255)
  if (name.length === 0) return ''
  if (!/^[\p{L}_]/u.test(name) || CELL_LIKE.test(name)) name = `_${name}`
  return name.slice(0, 255)
}

export function fileTableId(sheetId: string, tableIndex: number): string {
  return `file-table-${sheetId}-${tableIndex}`
}

/// Names every header-bearing file table for the formula engine. Names are
/// unique case-insensitively across the workbook (Excel's rule; Univer only
/// checks exact matches) and never collide with a sheet name or a name in
/// `reserved` (session tables).
export function planFileTableRegistrations(
  sheets: readonly FileSheet[],
  reserved: Iterable<string> = [],
): FileTableRegistration[] {
  const taken = new Set<string>()
  for (const sheet of sheets) taken.add(sheet.name.toLowerCase())
  for (const name of reserved) taken.add(name.toLowerCase())
  const plan: FileTableRegistration[] = []
  let fallback = 0
  for (const sheet of sheets) {
    sheet.tables.forEach((table, tableIndex) => {
      // Univer's table header is not optional yet: registering a headerless
      // table injects synthesized "Column N" labels over the first data row.
      if (table.headerRowCount === 0) return
      let base = sanitizeTableName(table.name ?? '')
      if (base.length === 0) base = `Table${++fallback}`
      let candidate = base
      for (let suffix = 2; taken.has(candidate.toLowerCase()); suffix += 1) {
        candidate = `${base.slice(0, 255 - `_${suffix}`.length)}_${suffix}`
      }
      taken.add(candidate.toLowerCase())
      const totals = table.totalsRowCount ?? 0
      const endRow = Math.max(table.range.startRow, table.range.endRow - totals)
      plan.push({
        sheetId: sheet.id,
        tableIndex,
        tableId: fileTableId(sheet.id, tableIndex),
        tableName: candidate,
        range: { ...table.range, endRow },
        columns: table.columns?.length ? table.columns : undefined,
      })
    })
  }
  return plan
}
