/// Read-only list of a workbook's data connections and Power Query queries
/// for Data › Queries & Connections (genoffice#2051). Nothing here refreshes
/// or rewrites a query: the parts stay untouched in the package, so Excel can
/// still refresh them, and this only says what the workbook depends on and
/// where each query's rows land.

import { relsPathFor, resolveRelTarget } from './xlsx-drawing-add'
import {
  decodeAttribute,
  parseRelationships,
  parseSheetElements,
  readAttribute,
  tableDisplayName,
} from './xlsx-sheets'

export type WorkbookQueryKind = 'powerQuery' | 'web' | 'text' | 'database' | 'other'

export interface WorkbookQuery {
  /// The Power Query query name, or the connection name for other kinds.
  readonly name: string
  readonly kind: WorkbookQueryKind
  /// URL (web) or file (text) the connection reads; absent for Power Query
  /// (its sources live inside the query definition) and databases (the
  /// connection string can carry credentials).
  readonly source?: string
  readonly description?: string
  /// Where the query's rows are loaded; absent for a connection-only query.
  readonly destination?: {
    readonly sheet: string
    readonly table?: string
    /// A1 range of the loaded rows, header included.
    readonly range: string
    /// Data rows (the header row excluded).
    readonly rows: number
  }
}

/// Reads a package part as text; undefined when the part does not exist.
export type PartReader = (path: string) => Promise<string | undefined>

const CONNECTIONS_PART = 'xl/connections.xml'
const WORKBOOK_PART = 'xl/workbook.xml'
/// A workbook with more query tables than this is not a real one; the walk stops.
const MAX_QUERY_TABLES = 500
/// The limits of `workbookQueriesResultSchema`: one hostile entry must not fail the whole list.
const MAX_CONNECTIONS = 1_000
const MAX_NAME_LENGTH = 1_000
const MAX_TEXT_LENGTH = 4_000
const A1_RANGE = /^[A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$/i

interface QueryTableHost {
  readonly name?: string
  readonly ref: string
  readonly header: number
  readonly totals: number
}

interface ParsedConnection {
  readonly id: string
  readonly query: Omit<WorkbookQuery, 'destination'>
}

function attribute(xml: string, name: string): string | undefined {
  const value = readAttribute(xml, name)
  return value === undefined ? undefined : decodeAttribute(value)
}

function clip(text: string, length: number): string {
  return text.length > length ? `${text.slice(0, length - 1)}…` : text
}

/// `Location=Rates` or `Location="Sales 2024"` in a Mashup OLE DB connection string.
function powerQueryName(connectionString: string): string | undefined {
  const match = /(?:^|;)\s*Location=(?:"([^"]*)"|([^;]*))/i.exec(connectionString)
  const name = (match?.[1] ?? match?.[2] ?? '').trim()
  return name || undefined
}

export function parseConnections(connectionsXml: string): ParsedConnection[] {
  const out: ParsedConnection[] = []
  for (const match of connectionsXml.matchAll(
    /<connection\b[^>]*?(?:\/>|>[\s\S]*?<\/connection>)/g,
  )) {
    const xml = match[0]
    const open = /<connection\b[^>]*>/.exec(xml)?.[0] ?? xml
    const id = readAttribute(open, 'id')
    if (id === undefined) continue
    const connectionName = clip(attribute(open, 'name') ?? `Connection ${id}`, MAX_NAME_LENGTH)
    const description = clip(attribute(open, 'description') ?? '', MAX_TEXT_LENGTH) || undefined
    const dbPr = /<dbPr\b[^>]*>/.exec(xml)?.[0]
    const connectionString = dbPr ? (attribute(dbPr, 'connection') ?? '') : ''
    const webUrl = /<webPr\b[^>]*>/.exec(xml)?.[0]
    const textPr = /<textPr\b[^>]*>/.exec(xml)?.[0]
    let query: ParsedConnection['query']
    if (/Provider=Microsoft\.Mashup\.OleDb/i.test(connectionString)) {
      query = {
        name: clip(powerQueryName(connectionString) ?? connectionName, MAX_NAME_LENGTH),
        kind: 'powerQuery',
        ...(description ? { description } : {}),
      }
    } else if (webUrl) {
      const url = attribute(webUrl, 'url')
      const source = url && clip(url, MAX_TEXT_LENGTH)
      query = { name: connectionName, kind: 'web', ...(source ? { source } : {}) }
    } else if (textPr) {
      const file = attribute(textPr, 'sourceFile')
      const source = file && clip(file, MAX_TEXT_LENGTH)
      query = { name: connectionName, kind: 'text', ...(source ? { source } : {}) }
    } else {
      query = { name: connectionName, kind: dbPr ? 'database' : 'other' }
    }
    if (description && query.kind !== 'powerQuery') query = { ...query, description }
    out.push({ id, query })
    if (out.length >= MAX_CONNECTIONS) break
  }
  return out
}

/// Data rows of an A1 range such as `A1:D21`, without its header and totals rows.
function dataRows(ref: string, headerRows: number, totalsRows: number): number {
  const rows = [...ref.matchAll(/\d+/g)].map((m) => Number(m[0]))
  const first = rows[0] ?? 0
  const height = Math.abs((rows[1] ?? first) - first) + 1
  return Math.max(0, height - headerRows - totalsRows)
}

function countAttribute(open: string, name: string, fallback: number): number {
  const value = Number(attribute(open, name) ?? fallback)
  return Number.isSafeInteger(value) && value >= 0 ? value : fallback
}

/// A sheet-scoped defined name's range: a query table loaded straight onto a
/// sheet has no table part, only a name such as `ExternalData_1`.
function definedRange(workbookXml: string, name: string, sheetIndex: number): string | undefined {
  for (const match of workbookXml.matchAll(/<definedName\b([^>]*)>([^<]*)<\/definedName>/g)) {
    const open = match[1] ?? ''
    if (attribute(open, 'name') !== name) continue
    const scope = attribute(open, 'localSheetId')
    if (scope !== undefined && Number(scope) !== sheetIndex) continue
    const range = decodeAttribute(match[2] ?? '')
      .split('!')
      .pop()
      ?.replaceAll('$', '')
    if (range && A1_RANGE.test(range)) return range
  }
  return undefined
}

async function relationshipsOf(read: PartReader, part: string) {
  const xml = await read(relsPathFor(part))
  return xml ? parseRelationships(xml) : []
}

function resolve(fromPart: string, target: string): string | undefined {
  try {
    return resolveRelTarget(fromPart, target)
  } catch {
    return undefined
  }
}

/// Every connection in the workbook with the sheet and table its rows load
/// into. An empty list means the workbook has no connections part.
export async function listWorkbookQueries(read: PartReader): Promise<WorkbookQuery[]> {
  const connectionsXml = await read(CONNECTIONS_PART)
  if (!connectionsXml) return []
  const connections = parseConnections(connectionsXml)
  const destinations = new Map<string, WorkbookQuery['destination']>()

  const workbookXml = (await read(WORKBOOK_PART)) ?? ''
  const workbookRels = await relationshipsOf(read, WORKBOOK_PART)
  let seen = 0
  /// False when the walk should stop.
  const addQueryTable = async (
    fromPart: string,
    sheet: string,
    table: QueryTableHost | undefined,
    sheetIndex: number,
    target: string,
  ): Promise<boolean> => {
    if (++seen > MAX_QUERY_TABLES) return false
    const queryPart = resolve(fromPart, target)
    const queryXml = queryPart && (await read(queryPart))
    const open = queryXml ? /<queryTable\b[^>]*>/.exec(queryXml)?.[0] : undefined
    const connectionId = open && readAttribute(open, 'connectionId')
    if (!open || !connectionId || destinations.has(connectionId)) return true
    const name = attribute(open, 'name')
    const ref = table?.ref ?? (name ? definedRange(workbookXml, name, sheetIndex) : undefined)
    if (!ref || !A1_RANGE.test(ref)) return true
    destinations.set(connectionId, {
      sheet,
      ...(table?.name ? { table: table.name } : {}),
      range: ref,
      rows: dataRows(
        ref,
        table ? table.header : countAttribute(open, 'headers', 1) === 0 ? 0 : 1,
        table?.totals ?? 0,
      ),
    })
    return true
  }
  for (const [sheetIndex, sheet] of parseSheetElements(workbookXml).entries()) {
    const rel = workbookRels.find((r) => r.id === sheet.relationshipId && !r.external)
    const sheetPart = rel && resolve(WORKBOOK_PART, rel.target)
    if (!sheetPart) continue
    for (const sheetRel of await relationshipsOf(read, sheetPart)) {
      if (sheetRel.external) continue
      if (/\/queryTable$/.test(sheetRel.type)) {
        const more = await addQueryTable(
          sheetPart,
          sheet.name,
          undefined,
          sheetIndex,
          sheetRel.target,
        )
        if (!more) return withDestinations(connections, destinations)
        continue
      }
      if (!/\/table$/.test(sheetRel.type)) continue
      const tablePart = resolve(sheetPart, sheetRel.target)
      const tableXml = tablePart && (await read(tablePart))
      if (!tablePart || !tableXml) continue
      const open = /<table\b[^>]*>/.exec(tableXml)?.[0] ?? ''
      const ref = attribute(open, 'ref')
      if (!ref) continue
      const name = tableDisplayName(tableXml)
      const table = {
        ...(name ? { name } : {}),
        ref,
        header: countAttribute(open, 'headerRowCount', 1),
        totals: countAttribute(open, 'totalsRowCount', 0),
      }
      for (const tableRel of await relationshipsOf(read, tablePart)) {
        if (tableRel.external || !/\/queryTable$/.test(tableRel.type)) continue
        const more = await addQueryTable(tablePart, sheet.name, table, sheetIndex, tableRel.target)
        if (!more) return withDestinations(connections, destinations)
      }
    }
  }
  return withDestinations(connections, destinations)
}

function withDestinations(
  connections: readonly ParsedConnection[],
  destinations: ReadonlyMap<string, WorkbookQuery['destination']>,
): WorkbookQuery[] {
  return connections.map(({ id, query }) => {
    const destination = destinations.get(id)
    return destination ? { ...query, destination } : query
  })
}
