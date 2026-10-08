import { relsPathFor, resolveRelTarget, type MutablePackage } from './xlsx-drawing-add'
import { escapeRegExp, removeRelationshipById } from './xlsx-sheets'
import {
  TABLE_REL_TYPE,
  areaToRef,
  areasOverlap,
  assertNameNotDefined,
  autoFilterRef,
  collectExistingTableNames,
  escapeAttribute,
  parseRef,
  type TableArea,
  type TableStyleOptions,
} from './xlsx-table-add'

/// Table Design changes to tables already in the file: the table part is
/// rewritten in place (name, ref, columns, header/totals bands, style info)
/// or unregistered for Convert to Range. Fail-closed like table additions.

export class TableEditError extends Error {}

export interface TableEdit extends TableStyleOptions {
  readonly worksheetPath: string
  /// displayName (or name) of the table part at open time.
  readonly tableName: string
  readonly name?: string | undefined
  readonly area?: TableArea | undefined
  readonly columnNames?: readonly string[] | undefined
  readonly style?: string | undefined
  readonly bandedRows?: boolean | undefined
  readonly remove?: true | undefined
}

const CONTENT_TYPES_PATH = '[Content_Types].xml'
const TABLE_NAME_RULE = /^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u
const CELL_LIKE = /^(\$?[A-Za-z]{1,3}\$?[0-9]+|[rR]\d*[cC]\d*)$/

interface LocatedTable {
  readonly path: string
  readonly relId: string
  readonly xml: string
}

export async function applyTableEdits(
  pkg: MutablePackage,
  edits: readonly TableEdit[],
  touchedEntries: Set<string>,
): Promise<void> {
  if (edits.length === 0) return
  for (const edit of edits) {
    const located = await locateTable(pkg, edit)
    if (edit.remove) {
      await removeTable(pkg, edit.worksheetPath, located, touchedEntries)
      continue
    }
    const rewritten = await rewriteTable(pkg, edit, located)
    pkg.write(located.path, rewritten)
    touchedEntries.add(located.path)
  }
}

async function sheetTableParts(
  pkg: MutablePackage,
  worksheetPath: string,
): Promise<LocatedTable[]> {
  const relsPath = relsPathFor(worksheetPath)
  if (!(await pkg.has(relsPath))) return []
  const relsXml = await pkg.readText(relsPath)
  const parts: LocatedTable[] = []
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
    const tag = match[0]
    if (!tag.includes(`Type="${TABLE_REL_TYPE}"`)) continue
    const target = /\bTarget="([^"]+)"/.exec(tag)?.[1]
    const relId = /\bId="([^"]+)"/.exec(tag)?.[1]
    if (!target || !relId) continue
    const path = resolveRelTarget(worksheetPath, target)
    if (!(await pkg.has(path))) continue
    parts.push({ path, relId, xml: await pkg.readText(path) })
  }
  return parts
}

function tableNames(xml: string): string[] {
  const open = /<table\b[^>]*>/.exec(xml)?.[0] ?? ''
  const names: string[] = []
  for (const attribute of ['displayName', 'name']) {
    const value = new RegExp(`\\b${attribute}="([^"]*)"`).exec(open)?.[1]
    if (value) names.push(unescapeAttribute(value))
  }
  return names
}

async function locateTable(pkg: MutablePackage, edit: TableEdit): Promise<LocatedTable> {
  const wanted = edit.tableName.toLowerCase()
  for (const part of await sheetTableParts(pkg, edit.worksheetPath)) {
    if (tableNames(part.xml).some((name) => name.toLowerCase() === wanted)) return part
  }
  throw new TableEditError(`Table "${edit.tableName}" was not found on its worksheet.`)
}

async function removeTable(
  pkg: MutablePackage,
  worksheetPath: string,
  table: LocatedTable,
  touchedEntries: Set<string>,
): Promise<void> {
  const relsPath = relsPathFor(worksheetPath)
  pkg.write(relsPath, removeRelationshipById(await pkg.readText(relsPath), table.relId))
  touchedEntries.add(relsPath)
  pkg.write(worksheetPath, removeTablePart(await pkg.readText(worksheetPath), table.relId))
  touchedEntries.add(worksheetPath)
  pkg.remove(table.path)
  touchedEntries.add(table.path)
  const contentTypes = await pkg.readText(CONTENT_TYPES_PATH)
  const stripped = contentTypes.replace(
    new RegExp(`<Override\\b[^>]*PartName="/${escapeRegExp(table.path)}"[^>]*/>`),
    '',
  )
  if (stripped !== contentTypes) {
    pkg.write(CONTENT_TYPES_PATH, stripped)
    touchedEntries.add(CONTENT_TYPES_PATH)
  }
}

/// Drops one <tablePart>; an emptied <tableParts> container goes with it.
export function removeTablePart(worksheetXml: string, relId: string): string {
  const without = worksheetXml.replace(
    new RegExp(`<tablePart\\b[^>]*\\br:id="${escapeRegExp(relId)}"[^>]*/>\\s*`),
    '',
  )
  const container = /<tableParts\b([^>]*?)(\/?)>/.exec(without)
  if (!container) return without
  const remaining = (without.match(/<tablePart\b/g) ?? []).length
  if (remaining === 0) {
    return without.replace(/<tableParts\b[^>]*?(?:\/>|>\s*<\/tableParts>)/, '')
  }
  if (/count="\d+"/.test(container[1] ?? '')) {
    return without.replace(
      container[0],
      container[0].replace(/count="\d+"/, `count="${remaining}"`),
    )
  }
  return without
}

async function rewriteTable(
  pkg: MutablePackage,
  edit: TableEdit,
  table: LocatedTable,
): Promise<string> {
  let xml = table.xml
  const openMatch = /<table\b[^>]*>/.exec(xml)
  if (!openMatch) throw new TableEditError(`Table part ${table.path} is malformed.`)
  const attributes = parseAttributes(openMatch[0])
  const currentRef = attributes.get('ref')
  if (!currentRef) throw new TableEditError(`Table part ${table.path} has no ref.`)

  if (edit.name !== undefined) {
    await assertRenameAllowed(pkg, edit, table)
    attributes.set('name', escapeAttribute(edit.name))
    attributes.set('displayName', escapeAttribute(edit.name))
  }

  const headerRows =
    edit.headerRow === undefined
      ? Number(attributes.get('headerRowCount') ?? '1')
      : edit.headerRow
        ? 1
        : 0
  const totalsRows =
    edit.totalsRow === undefined
      ? Number(attributes.get('totalsRowCount') ?? '0')
      : edit.totalsRow
        ? 1
        : 0
  const area = edit.area ?? parseRef(currentRef)
  if (area.endRow - area.startRow + 1 <= headerRows + totalsRows) {
    throw new TableEditError(
      `Table "${edit.tableName}" needs at least one data row beside its header and totals rows.`,
    )
  }
  if (edit.area) {
    await assertAreaFree(pkg, edit, table, area)
    attributes.set('ref', areaToRef(area))
  }
  if (headerRows === 0) attributes.set('headerRowCount', '0')
  else attributes.delete('headerRowCount')
  if (totalsRows > 0) {
    attributes.set('totalsRowCount', String(totalsRows))
    attributes.set('totalsRowShown', '1')
  } else {
    attributes.delete('totalsRowCount')
    if (edit.totalsRow === false) attributes.set('totalsRowShown', '0')
  }
  xml = xml.replace(openMatch[0], `<table ${serializeAttributes(attributes)}>`)

  const previousArea = parseRef(currentRef)
  xml = rewriteColumns(xml, edit, area, previousArea)

  const filterPattern = /<autoFilter\b[^>]*\/>|<autoFilter\b[^>]*>[\s\S]*?<\/autoFilter>/
  const existingFilter = filterPattern.exec(xml)?.[0]
  const wantsFilter =
    headerRows > 0 &&
    (edit.filterButton ?? (edit.headerRow === true ? true : existingFilter !== undefined)) === true
  const columnsMoved =
    area.startColumn !== previousArea.startColumn || area.endColumn !== previousArea.endColumn
  const filterRef = autoFilterRef(area, totalsRows)
  if (!wantsFilter) {
    xml = xml.replace(filterPattern, '')
  } else if (existingFilter === undefined) {
    xml = xml.replace(/<table\b[^>]*>/, (open) => `${open}<autoFilter ref="${filterRef}"/>`)
  } else if (columnsMoved || edit.headerRow !== undefined) {
    // Criteria index columns from the filter's left edge: a column shift or
    // a re-headed table would point them at the wrong cells, so they drop.
    xml = xml.replace(filterPattern, `<autoFilter ref="${filterRef}"/>`)
  } else {
    xml = xml.replace(
      existingFilter,
      existingFilter.replace(/(<autoFilter\b[^>]*?\bref=")[^"]*(")/, `$1${filterRef}$2`),
    )
  }

  return rewriteStyleInfo(xml, edit)
}

async function assertRenameAllowed(
  pkg: MutablePackage,
  edit: TableEdit,
  table: LocatedTable,
): Promise<void> {
  const name = edit.name ?? ''
  if (
    name.length === 0 ||
    name.length > 255 ||
    !TABLE_NAME_RULE.test(name) ||
    CELL_LIKE.test(name)
  ) {
    throw new TableEditError(`"${name}" is not a valid table name.`)
  }
  const own = new Set(tableNames(table.xml).map((value) => value.toLowerCase()))
  const reserved = await collectExistingTableNames(pkg)
  for (const value of own) reserved.delete(value)
  if (reserved.has(name.toLowerCase())) {
    throw new TableEditError(`Table name "${name}" is already taken in this workbook.`)
  }
  await assertNameNotDefined(pkg, name)
}

async function assertAreaFree(
  pkg: MutablePackage,
  edit: TableEdit,
  table: LocatedTable,
  area: TableArea,
): Promise<void> {
  for (const part of await sheetTableParts(pkg, edit.worksheetPath)) {
    if (part.path === table.path) continue
    const ref = /<table\b[^>]*\bref="([^"]+)"/.exec(part.xml)?.[1]
    if (ref && areasOverlap(area, parseRef(ref))) {
      throw new TableEditError(`Table "${edit.tableName}" would overlap table ${ref}.`)
    }
  }
  const worksheetXml = await pkg.readText(edit.worksheetPath)
  for (const match of worksheetXml.matchAll(/<mergeCell\b[^>]*\bref="([^"]+)"/g)) {
    const ref = match[1]
    if (ref && areasOverlap(area, parseRef(ref))) {
      throw new TableEditError(
        `Table "${edit.tableName}" would overlap merged cells (${ref}) — unmerge them first.`,
      )
    }
  }
}

/// Existing <tableColumn> elements follow their sheet column, so a left-edge
/// resize keeps ids, formulas and totals metadata with the right cells.
function rewriteColumns(
  xml: string,
  edit: TableEdit,
  area: TableArea,
  previousArea: TableArea,
): string {
  const width = area.endColumn - area.startColumn + 1
  const block = /<tableColumns\b[^>]*>([\s\S]*?)<\/tableColumns>/.exec(xml)
  if (!block) throw new TableEditError(`Table "${edit.tableName}" has no tableColumns.`)
  const existing = [
    ...(block[1] ?? '').matchAll(/<tableColumn\b[^>]*?(?:\/>|>[\s\S]*?<\/tableColumn>)/g),
  ].map((match) => match[0])
  if (edit.columnNames === undefined) {
    if (existing.length !== width) {
      throw new TableEditError(
        `Table "${edit.tableName}" spans ${width} columns but lists ${existing.length} column names.`,
      )
    }
    return xml
  }
  if (edit.columnNames.length !== width) {
    throw new TableEditError(
      `Table "${edit.tableName}" spans ${width} columns but has ${edit.columnNames.length} column names.`,
    )
  }
  const seen = new Set<string>()
  for (const name of edit.columnNames) {
    const key = name.trim().toLowerCase()
    if (key.length === 0)
      throw new TableEditError(`Table "${edit.tableName}" has a blank column name.`)
    if (seen.has(key)) {
      throw new TableEditError(`Table "${edit.tableName}" has duplicate column name "${name}".`)
    }
    seen.add(key)
  }
  let nextId = existing.reduce(
    (max, element) => Math.max(max, Number(/\bid="(\d+)"/.exec(element)?.[1] ?? 0)),
    0,
  )
  const columns = edit.columnNames.map((name, index) => {
    const previousIndex = area.startColumn + index - previousArea.startColumn
    const element = previousIndex >= 0 ? existing[previousIndex] : undefined
    const escaped = escapeAttribute(name)
    if (element === undefined) {
      nextId += 1
      return `<tableColumn id="${nextId}" name="${escaped}"/>`
    }
    return /\bname="/.test(element)
      ? element.replace(/\bname="[^"]*"/, `name="${escaped}"`)
      : element.replace(/<tableColumn\b/, `<tableColumn name="${escaped}"`)
  })
  return xml.replace(
    block[0],
    `<tableColumns count="${columns.length}">${columns.join('')}</tableColumns>`,
  )
}

function rewriteStyleInfo(xml: string, edit: TableEdit): string {
  const existing =
    /<tableStyleInfo\b[^>]*\/>|<tableStyleInfo\b[^>]*>[\s\S]*?<\/tableStyleInfo>/.exec(xml)
  const attributes = existing
    ? parseAttributes(existing[0])
    : new Map<string, string>([
        ['showFirstColumn', '0'],
        ['showLastColumn', '0'],
        ['showRowStripes', '1'],
        ['showColumnStripes', '0'],
      ])
  if (edit.style !== undefined) attributes.set('name', escapeAttribute(edit.style))
  const flags: [keyof TableEdit, string][] = [
    ['firstColumn', 'showFirstColumn'],
    ['lastColumn', 'showLastColumn'],
    ['bandedRows', 'showRowStripes'],
    ['bandedColumns', 'showColumnStripes'],
  ]
  for (const [key, attribute] of flags) {
    const value = edit[key]
    if (typeof value === 'boolean') attributes.set(attribute, value ? '1' : '0')
  }
  const element = `<tableStyleInfo ${serializeAttributes(attributes)}/>`
  if (existing) return xml.replace(existing[0], element)
  const extLstAt = xml.search(/<extLst\b/)
  if (extLstAt >= 0) return xml.slice(0, extLstAt) + element + xml.slice(extLstAt)
  return xml.replace('</table>', `${element}</table>`)
}

function parseAttributes(openTag: string): Map<string, string> {
  const attributes = new Map<string, string>()
  const body = openTag.replace(/^<[^\s/>]+/, '').replace(/\/?>$/, '')
  for (const match of body.matchAll(/([^\s=]+)="([^"]*)"/g)) {
    attributes.set(match[1]!, match[2]!)
  }
  return attributes
}

function serializeAttributes(attributes: Map<string, string>): string {
  return [...attributes].map(([key, value]) => `${key}="${value}"`).join(' ')
}

function unescapeAttribute(input: string): string {
  return input
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, '>')
    .replace(/&lt;/g, '<')
    .replace(/&amp;/g, '&')
}
