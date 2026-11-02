/// Where a fetched query table lands, and how it becomes the values the grid
/// takes.
///
/// A workbook's web query fills a range on a sheet. Finding the query and its
/// destination is the listing side's job: this module builds on
/// `listWorkbookQueries` (genoffice#2345), which walks both shapes Excel
/// writes — a ListObject-backed query (the range lives in
/// `xl/tables/tableN.xml`) and a legacy range query (the range lives in a
/// sheet-scoped `ExternalData_N` defined name). This side answers the two
/// questions a refresh needs: "which URL does this connection read?" and
/// "given the fetched table and the destination range, what goes in each
/// cell?".
///
/// It deliberately does not touch the package: the caller writes the values
/// through the app's normal grid edits, and the save path stores them.

import { listWorkbookQueries, parseConnections, type PartReader } from './xlsx-connections'

export type { PartReader }

/// A connection the Data menu can refresh, with the range its rows fill.
export interface RefreshableQuery {
  readonly id: string
  readonly name: string
  /// Always a web query for now: a text query reads a local path, which the
  /// SSRF gate rejects, so text connections are left out of the listing.
  readonly kind: 'web'
  /// The URL the connection reads.
  readonly source: string
  readonly sheetName: string
  /// A1 range of the loaded rows, header included.
  readonly range: string
}

export interface RefreshTarget {
  readonly sheetName: string
  /// 0-based, as the cell-edit API uses.
  readonly startRow: number
  /// 0-based, as the cell-edit API uses.
  readonly startColumn: number
  readonly width: number
  readonly height: number
}

export interface RefreshedTable {
  readonly headers: readonly string[]
  readonly rows: readonly (readonly string[])[]
}

/// One cell of a refresh plan: a typed value to write, or null to clear.
export type RefreshCell = string | number | null

export interface TableRefreshPlan {
  /// Row-major over the destination range; null cells clear what was there.
  readonly values: readonly (readonly RefreshCell[])[]
  /// Rows the fetched table had beyond the destination range (the range is
  /// not grown yet, so they are dropped and reported to the user).
  readonly droppedRows: number
  /// Columns the fetched table had beyond the destination range.
  readonly droppedColumns: number
}

/**
 * The connections this build can refresh: web queries with a destination.
 *
 * A text query is not listed: it reads a local file, which the SSRF gate
 * rightly rejects, so it would only fail with "could not refresh". A
 * connection-only query has no rows of its own, and a Power Query needs an M
 * interpreter for its DataMashup definition.
 *
 * `listWorkbookQueries` returns one entry per `<connection>` in document
 * order, and `parseConnections` returns the same connections in the same
 * order, so the connection id is zipped in from the latter.
 */
export async function refreshableQueriesOf(read: PartReader): Promise<RefreshableQuery[]> {
  const queries = await listWorkbookQueries(read)
  if (queries.length === 0) return []
  const connectionsXml = await read('xl/connections.xml')
  const ids = connectionsXml === undefined ? [] : parseConnections(connectionsXml).map((c) => c.id)
  const out: RefreshableQuery[] = []
  for (const [index, query] of queries.entries()) {
    const id = ids[index]
    if (id === undefined || query.kind !== 'web' || !query.source || !query.destination) continue
    out.push({
      id,
      name: query.name,
      kind: 'web',
      source: query.source,
      sheetName: query.destination.sheet,
      range: query.destination.range,
    })
  }
  return out
}

/**
 * The URL a refreshable connection reads, resolved from the workbook part
 * itself. The renderer names only the connection id, so this IPC stays a
 * "refresh this workbook's query" call instead of a general fetch proxy.
 * Null when the workbook has no such web connection.
 */
export async function connectionSourceOf(
  read: PartReader,
  connectionId: string,
): Promise<string | null> {
  const connectionsXml = await read('xl/connections.xml')
  if (connectionsXml === undefined) return null
  const connection = parseConnections(connectionsXml).find((c) => c.id === connectionId)
  const query = connection?.query
  // A text, Power Query or database connection has no URL this path can fetch.
  return query?.kind === 'web' && query.source ? query.source : null
}

/** `A1` → column 0, `B14` → row 13. */
function columnOf(letters: string): number {
  let index = 0
  for (const ch of letters) index = index * 26 + (ch.charCodeAt(0) - 64)
  return index - 1
}

const A1_CELL = /^([A-Za-z]{1,3})([1-9][0-9]{0,6})$/
/// Excel's grid: 1,048,576 rows by 16,384 columns.
const MAX_GRID_ROWS = 1_048_576
const MAX_GRID_COLUMNS = 16_384
/// A refresh rewrites the destination cell by cell, so the plan allocates one
/// value per cell: a range this large is a corrupt or hostile package, not a
/// real query table.
const MAX_TARGET_CELLS = 1_000_000

/**
 * A range like `A1:D21` as a target. A single cell (`A1`) is a one-by-one
 * range, which is how a query table with one column is written.
 */
export function a1RangeToTarget(range: string, sheetName: string): RefreshTarget | null {
  const parts = range.split(':')
  const first = A1_CELL.exec(parts[0]?.trim() ?? '')
  if (!first) return null
  const last = parts.length > 1 ? A1_CELL.exec(parts[1]?.trim() ?? '') : null
  const startRow = Number(first[2]) - 1
  const startColumn = columnOf(first[1]!.toUpperCase())
  const endRow = last ? Number(last[2]) - 1 : startRow
  const endColumn = last ? columnOf(last[1]!.toUpperCase()) : startColumn
  if (startRow < 0 || startColumn < 0) return null
  if (endRow < startRow || endColumn < startColumn) return null
  if (endRow >= MAX_GRID_ROWS || endColumn >= MAX_GRID_COLUMNS) return null
  const width = endColumn - startColumn + 1
  const height = endRow - startRow + 1
  if (width * height > MAX_TARGET_CELLS) return null
  return { sheetName, startRow, startColumn, width, height }
}

/// A plain numeric token: signs, digits, an optional decimal part, an optional
/// exponent — but no grouping, currency or percent sign. A "1,09" from one
/// locale and a "1.09" from another cannot be told apart without the query's
/// `textPr` separators, so both stay text rather than become a wrong number.
const PLAIN_NUMBER = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/

/**
 * One fetched cell as a grid value. Excel's legacy web queries recognise
 * numbers and leave only what they cannot parse as text — that is what keeps
 * `=SUM(B2:B6)` over a refreshed rate table working instead of adding zeros.
 * Dates are not converted here.
 */
function refreshCellValue(text: string | undefined): RefreshCell {
  if (text === undefined || text === '') return null
  const trimmed = text.trim()
  if (!PLAIN_NUMBER.test(trimmed)) return text
  const value = Number(trimmed)
  return Number.isFinite(value) ? value : text
}

/**
 * The fetched table as the values to write into the destination range.
 *
 * The whole destination is covered, not just the rows the answer happened to
 * return: Excel's default refresh clears the cells a shrinking answer no
 * longer fills, and a stale row left behind reads as current data. Growing the
 * range is out of scope, so rows and columns that no longer fit are dropped
 * and counted for the status message.
 */
export function planTableRefresh(target: RefreshTarget, table: RefreshedTable): TableRefreshPlan {
  const sourceRows: readonly (readonly string[])[] = [table.headers, ...table.rows]
  let sourceWidth = 0
  for (const row of sourceRows) sourceWidth = Math.max(sourceWidth, row.length)
  const values: RefreshCell[][] = []
  for (let row = 0; row < target.height; row++) {
    const source = sourceRows[row]
    const cells: RefreshCell[] = []
    for (let column = 0; column < target.width; column++) {
      cells.push(refreshCellValue(source?.[column]))
    }
    values.push(cells)
  }
  return {
    values,
    droppedRows: Math.max(0, sourceRows.length - target.height),
    droppedColumns: Math.max(0, sourceWidth - target.width),
  }
}
