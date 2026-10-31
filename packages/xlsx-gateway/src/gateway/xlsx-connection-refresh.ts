/// Where a refreshed table lands, and how it becomes cell edits.
///
/// A workbook's web or text query fills a range: the query table part carries
/// the `ref` it occupies, and the sheet that owns it carries the name. This
/// module turns that range into a target and turns a fetched table into the
/// cell edits that write it there.
///
/// It deliberately does not read `xl/connections.xml` — finding the queries a
/// workbook has is the listing side's job (genoffice#2051 stage 1). This side
/// only answers "given this range and this table, what do I write?".

export interface RefreshTarget {
  readonly sheetName: string
  /// 0-based, as the cell-edit API uses.
  readonly startRow: number
  /// 0-based, as the cell-edit API uses.
  readonly startColumn: number
  readonly width: number
  readonly height: number
}

import { resolveRelTarget } from './xlsx-drawing-add'

export interface RefreshedTable {
  readonly headers: readonly string[]
  readonly rows: readonly (readonly string[])[]
}

/// A connection that can be refreshed without a query engine: a web query
/// names a URL, a text query names a file. A Power Query connection does not
/// qualify — its M code lives in the DataMashup part and needs an interpreter.
export interface RefreshableConnection {
  readonly id: string
  readonly name: string
  readonly kind: 'web' | 'text'
  /// The URL (web) or the file path (text) the connection reads.
  readonly source: string
}

/// A query table part: which connection it belongs to, and where it lands.
export interface QueryTableBinding {
  readonly connectionId: string
  readonly ref: string
}

/**
 * The connections a refresh can act on, read from `xl/connections.xml`.
 *
 * The file uses one `<connection>` element per source, with the kind in a
 * child element: `<webPr url>` for a web query, `<textPr sourceFile>` for a
 * text query, `<dbPr connection>` for a database. Only the first two name a
 * place to read from, and a database connection's string can carry
 * credentials, so it is not listed.
 */
export function refreshableConnections(connectionsXml: string): RefreshableConnection[] {
  const out: RefreshableConnection[] = []
  for (const match of connectionsXml.matchAll(
    /<connection\b[^>]*?(?:\/>|>[\s\S]*?<\/connection>)/g,
  )) {
    const xml = match[0]
    const open = /<connection\b[^>]*>/.exec(xml)?.[0] ?? xml
    const id = /\bid="([^"]+)"/.exec(open)?.[1]
    if (!id) continue
    const name = /\bname="([^"]*)"/.exec(open)?.[1] ?? `Connection ${id}`
    const webUrl = /<webPr\b[^>]*>/.exec(xml)?.[0]
    const webSource = webUrl ? /\burl="([^"]*)"/.exec(webUrl)?.[1] : undefined
    if (webSource) {
      out.push({ id, name, kind: 'web', source: webSource })
      continue
    }
    const textPr = /<textPr\b[^>]*>/.exec(xml)?.[0]
    const textSource = textPr ? /\bsourceFile="([^"]*)"/.exec(textPr)?.[1] : undefined
    if (textSource) {
      out.push({ id, name, kind: 'text', source: textSource })
    }
  }
  return out
}

/**
 * The query table part's binding: the connection it refreshes and the range
 * it fills. The sheet that owns it says which sheet that range is on.
 */
export function queryTableBinding(queryTableXml: string): QueryTableBinding | null {
  const open = /<queryTable\b[^>]*>/.exec(queryTableXml)?.[0]
  if (!open) return null
  const connectionId = /\bconnectionId="([^"]+)"/.exec(open)?.[1]
  const ref = /\bref="([^"]+)"/.exec(open)?.[1]
  if (!connectionId || !ref) return null
  return { connectionId, ref }
}

/// A query table bound to a sheet: where its rows land.
export interface BoundConnection {
  readonly connection: RefreshableConnection
  readonly sheetName: string
  readonly range: string
}

/// The package parts the connection walk needs.
export interface ConnectionPackageParts {
  readonly connectionsXml: string
  readonly workbookXml: string
  readonly workbookRelsXml: string
  /// The `xl/worksheets/_rels/sheetN.xml.rels` parts, with the sheet part they
  /// belong to (`xl/worksheets/sheetN.xml`) — a relationship target is relative
  /// to that part, not to the rels file.
  readonly sheetRels: readonly { readonly partPath: string; readonly xml: string }[]
  /// The `xl/queryTables/queryTableN.xml` parts, with their part path.
  readonly queryTables: readonly { readonly path: string; readonly xml: string }[]
}

const QUERY_TABLE_REL = 'relationships/queryTable'

/** The sheet a `xl/worksheets/sheetN.xml` belongs to, by its name. */
function sheetNamesByPart(workbookXml: string, workbookRelsXml: string): Map<string, string> {
  const relTargets = new Map<string, string>()
  for (const match of workbookRelsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const open = match[0]
    const id = /\bId="([^"]+)"/.exec(open)?.[1]
    const target = /\bTarget="([^"]+)"/.exec(open)?.[1]
    if (!id || !target) continue
    try {
      relTargets.set(id, resolveRelTarget('xl/workbook.xml', target))
    } catch {
      /* an unresolvable target names no sheet */
    }
  }
  const names = new Map<string, string>()
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*\/>/g)) {
    const open = match[0]
    const name = /\bname="([^"]*)"/.exec(open)?.[1]
    const relId = /\br:id="([^"]+)"/.exec(open)?.[1]
    if (!name || !relId) continue
    const target = relTargets.get(relId)
    if (!target) continue
    names.set(target, name.replace(/&amp;/g, '&'))
  }
  return names
}

/** The query table part a sheet's relationships point at. */
function queryTableTargetsOf(sheetPartPath: string, sheetRelsXml: string): string[] {
  const targets: string[] = []
  for (const match of sheetRelsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const open = match[0]
    if (!open.includes(QUERY_TABLE_REL)) continue
    const target = /\bTarget="([^"]+)"/.exec(open)?.[1]
    if (!target) continue
    try {
      targets.push(resolveRelTarget(sheetPartPath, target))
    } catch {
      /* an unresolvable target names no query table */
    }
  }
  return targets
}

/**
 * The queries this workbook can refresh, with the sheet and range each one
 * fills. A connection with no query table is a connection-only query: it has
 * no rows of its own, so there is nowhere to write the answer and it is left
 * out.
 */
export function boundConnections(parts: ConnectionPackageParts): BoundConnection[] {
  const connections = refreshableConnections(parts.connectionsXml)
  if (connections.length === 0) return []
  const sheetNames = sheetNamesByPart(parts.workbookXml, parts.workbookRelsXml)
  const queryTables = new Map(parts.queryTables.map((q) => [q.path, q.xml]))
  const out: BoundConnection[] = []
  for (const sheetRels of parts.sheetRels) {
    const sheetName = sheetNames.get(sheetRels.partPath)
    if (!sheetName) continue
    for (const target of queryTableTargetsOf(sheetRels.partPath, sheetRels.xml)) {
      const xml = queryTables.get(target)
      if (!xml) continue
      const binding = queryTableBinding(xml)
      if (!binding) continue
      const connection = connections.find((c) => c.id === binding.connectionId)
      if (!connection) continue
      out.push({ connection, sheetName, range: binding.ref })
    }
  }
  return out
}

/// A cell edit the save path understands.
export interface TableCellEdit {
  readonly sheetName: string
  readonly row: number
  readonly column: number
  readonly writeValue: true
  readonly cell: { readonly value: string | null }
}

/** `A1` → column 0, `B14` → row 13. */
function columnOf(letters: string): number {
  let index = 0
  for (const ch of letters) index = index * 26 + (ch.charCodeAt(0) - 64)
  return index - 1
}

const A1_CELL = /^([A-Za-z]{1,3})([1-9][0-9]{0,6})$/

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
  return {
    sheetName,
    startRow,
    startColumn,
    width: endColumn - startColumn + 1,
    height: endRow - startRow + 1,
  }
}

/**
 * The table as cell edits, written from the target's top-left corner.
 *
 * Values go in as text, the way Excel's own web query imports them: a page
 * that means "1.09" and a page that means "1,09" cannot be told apart
 * without the page's locale, and a number guessed wrong is worse than a
 * string the user can convert. The header row is written too — the query's
 * range includes it.
 */
export function tableToCellEdits(target: RefreshTarget, table: RefreshedTable): TableCellEdit[] {
  const edits: TableCellEdit[] = []
  const push = (row: number, column: number, value: string): void => {
    if (column - target.startColumn >= target.width) return
    if (row - target.startRow >= target.height) return
    edits.push({
      sheetName: target.sheetName,
      row,
      column,
      writeValue: true,
      cell: { value },
    })
  }
  table.headers.forEach((header, column) => {
    push(target.startRow, target.startColumn + column, header)
  })
  table.rows.forEach((row, rowIndex) => {
    row.forEach((value, column) => {
      push(target.startRow + 1 + rowIndex, target.startColumn + column, value)
    })
  })
  return edits
}
