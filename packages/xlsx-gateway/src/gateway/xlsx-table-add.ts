import { columnLabel } from '../domain/cell-address'
import {
  allocatePartPath,
  appendRelationship,
  registerContentTypeOverride,
  relativeTarget,
  relsPathFor,
  type MutablePackage,
} from './xlsx-drawing-add'
import { ensureRelationshipNamespace } from './xlsx-namespace'

/// Persists tables (ListObjects) created in the editor as brand-new OOXML
/// parts: xl/tables/tableN.xml, the worksheet's <tableParts> element, the
/// worksheet→table relationship, and the [Content_Types].xml override.
/// Fail-closed: name collisions, overlaps with existing tables / the sheet
/// auto-filter / merged cells, and bad column names all abort the save.

export class TableAddError extends Error {}

export interface TableArea {
  readonly startRow: number
  readonly startColumn: number
  readonly endRow: number
  readonly endColumn: number
}

export interface TableAddition {
  readonly worksheetPath: string
  /// 0-based, header row included.
  readonly area: TableArea
  readonly name: string
  /// One per column, in order; must be non-blank and unique (the renderer
  /// sanitizes screen values before recording the add).
  readonly columnNames: readonly string[]
  /// Built-in style name; undefined = TableStyleMedium2.
  readonly style?: string | undefined
  readonly bandedRows: boolean
}

const TABLE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
const TABLE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
const DEFAULT_TABLE_STYLE = 'TableStyleMedium2'

export async function applyTableAdditions(
  pkg: MutablePackage,
  additions: readonly TableAddition[],
  touchedEntries: Set<string>,
): Promise<void> {
  if (additions.length === 0) return
  const reserved = await collectExistingTableNames(pkg)
  let nextId = await maxExistingTableId(pkg)
  for (const addition of additions) {
    validateAddition(addition)
    const lowerName = addition.name.toLowerCase()
    if (reserved.has(lowerName)) {
      throw new TableAddError(`Table name "${addition.name}" is already taken in this workbook.`)
    }
    reserved.add(lowerName)
    await assertNameNotDefined(pkg, addition.name)
    await assertNoTableOverlap(pkg, addition)
    assertNoSheetConflicts(addition, await pkg.readText(addition.worksheetPath))

    nextId += 1
    const tablePath = await allocatePartPath(pkg, 'xl/tables/table', '.xml')
    pkg.add(tablePath, buildTableXml(nextId, addition))
    touchedEntries.add(tablePath)
    await registerContentTypeOverride(pkg, tablePath, TABLE_CONTENT_TYPE, touchedEntries)
    const relId = await appendRelationship(
      pkg,
      relsPathFor(addition.worksheetPath),
      TABLE_REL_TYPE,
      relativeTarget(addition.worksheetPath, tablePath),
    )
    touchedEntries.add(relsPathFor(addition.worksheetPath))
    pkg.write(
      addition.worksheetPath,
      appendTablePart(await pkg.readText(addition.worksheetPath), relId),
    )
    touchedEntries.add(addition.worksheetPath)
  }
}

function validateAddition(addition: TableAddition): void {
  const { area, columnNames } = addition
  if (area.endRow <= area.startRow) {
    throw new TableAddError(
      `Table "${addition.name}" needs a header row plus at least one data row.`,
    )
  }
  const width = area.endColumn - area.startColumn + 1
  if (width !== columnNames.length) {
    throw new TableAddError(
      `Table "${addition.name}" spans ${width} columns but has ${columnNames.length} column names.`,
    )
  }
  const seen = new Set<string>()
  for (const name of columnNames) {
    const normalized = name.trim().toLowerCase()
    if (normalized.length === 0) {
      throw new TableAddError(`Table "${addition.name}" has a blank column name.`)
    }
    if (seen.has(normalized)) {
      throw new TableAddError(`Table "${addition.name}" has duplicate column name "${name}".`)
    }
    seen.add(normalized)
  }
}

/// name= and displayName= of every existing table part, lowercased.
async function collectExistingTableNames(pkg: MutablePackage): Promise<Set<string>> {
  const names = new Set<string>()
  for (const path of await tablePartPaths(pkg)) {
    const xml = await pkg.readText(path)
    for (const attribute of ['name', 'displayName']) {
      const value = new RegExp(`<table\\b[^>]*\\b${attribute}="([^"]+)"`).exec(xml)?.[1]
      if (value) names.add(value.toLowerCase())
    }
  }
  return names
}

async function maxExistingTableId(pkg: MutablePackage): Promise<number> {
  let max = 0
  for (const path of await tablePartPaths(pkg)) {
    const xml = await pkg.readText(path)
    const id = Number(/<table\b[^>]*\bid="(\d+)"/.exec(xml)?.[1] ?? 0)
    if (Number.isFinite(id)) max = Math.max(max, id)
  }
  return max
}

async function tablePartPaths(pkg: MutablePackage): Promise<string[]> {
  return (await pkg.paths()).filter((path) => /^xl\/tables\/[^/]+\.xml$/.test(path))
}

/// Table names share Excel's defined-name namespace.
async function assertNameNotDefined(pkg: MutablePackage, name: string): Promise<void> {
  const workbookXml = await pkg.readText('xl/workbook.xml')
  for (const match of workbookXml.matchAll(/<definedName\b[^>]*\bname="([^"]+)"/g)) {
    if (match[1]?.toLowerCase() === name.toLowerCase()) {
      throw new TableAddError(`Table name "${name}" collides with a defined name in the workbook.`)
    }
  }
}

/// The new table must not overlap any table already on the same worksheet.
/// Additions earlier in this save already appended their relationships, so
/// the rels scan covers unsaved tables too.
async function assertNoTableOverlap(pkg: MutablePackage, addition: TableAddition): Promise<void> {
  const relsPath = relsPathFor(addition.worksheetPath)
  if (!(await pkg.has(relsPath))) return
  const relsXml = await pkg.readText(relsPath)
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
    const tag = match[0]
    if (!tag.includes(`Type="${TABLE_REL_TYPE}"`)) continue
    const target = /\bTarget="([^"]+)"/.exec(tag)?.[1]
    if (!target) continue
    const tablePath = resolveTarget(addition.worksheetPath, target)
    if (!(await pkg.has(tablePath))) continue
    const ref = /<table\b[^>]*\bref="([^"]+)"/.exec(await pkg.readText(tablePath))?.[1]
    if (ref && areasOverlap(addition.area, parseRef(ref))) {
      throw new TableAddError(`Table "${addition.name}" overlaps an existing table (${ref}).`)
    }
  }
}

/// Excel forbids a table overlapping the sheet auto-filter or merged cells.
function assertNoSheetConflicts(addition: TableAddition, worksheetXml: string): void {
  const autoFilterRef = /<autoFilter\b[^>]*\bref="([^"]+)"/.exec(worksheetXml)?.[1]
  if (autoFilterRef && areasOverlap(addition.area, parseRef(autoFilterRef))) {
    throw new TableAddError(
      `Table "${addition.name}" overlaps the sheet auto-filter (${autoFilterRef}) — clear it first.`,
    )
  }
  for (const match of worksheetXml.matchAll(/<mergeCell\b[^>]*\bref="([^"]+)"/g)) {
    const ref = match[1]
    if (ref && areasOverlap(addition.area, parseRef(ref))) {
      throw new TableAddError(
        `Table "${addition.name}" overlaps merged cells (${ref}) — unmerge them first.`,
      )
    }
  }
}

/// Resolves a relationship target relative to its source part.
function resolveTarget(fromPart: string, target: string): string {
  const base = fromPart.split('/').slice(0, -1)
  for (const segment of target.split('/')) {
    if (segment === '..') base.pop()
    else if (segment !== '.' && segment !== '') base.push(segment)
  }
  return base.join('/')
}

function buildTableXml(id: number, addition: TableAddition): string {
  const ref = areaToRef(addition.area)
  const name = escapeAttribute(addition.name)
  const columns = addition.columnNames
    .map(
      (columnName, index) =>
        `<tableColumn id="${index + 1}" name="${escapeAttribute(columnName)}"/>`,
    )
    .join('')
  const style = escapeAttribute(addition.style ?? DEFAULT_TABLE_STYLE)
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    `id="${id}" name="${name}" displayName="${name}" ref="${ref}" totalsRowShown="0">` +
    `<autoFilter ref="${ref}"/>` +
    `<tableColumns count="${addition.columnNames.length}">${columns}</tableColumns>` +
    `<tableStyleInfo name="${style}" showFirstColumn="0" showLastColumn="0" ` +
    `showRowStripes="${addition.bandedRows ? 1 : 0}" showColumnStripes="0"/>` +
    '</table>'
  )
}

/// Appends (or extends) the worksheet's <tableParts> element. Schema order
/// puts tableParts after every print/drawing element, before extLst.
function appendTablePart(worksheetXml: string, relId: string): string {
  const xml = ensureRelationshipNamespace(worksheetXml)
  const part = `<tablePart r:id="${relId}"/>`
  const existing = /<tableParts\b[^>]*count="(\d+)"[^>]*>/.exec(xml)
  if (existing) {
    const opener = existing[0]
    const count = Number(existing[1] ?? 0) + 1
    const reopened = opener.replace(/count="\d+"/, `count="${count}"`)
    return xml.replace(opener, reopened).replace('</tableParts>', `${part}</tableParts>`)
  }
  const element = `<tableParts count="1">${part}</tableParts>`
  const extLstAt = xml.search(/<extLst\b/)
  if (extLstAt >= 0) return xml.slice(0, extLstAt) + element + xml.slice(extLstAt)
  const closeAt = xml.lastIndexOf('</worksheet>')
  if (closeAt < 0) throw new TableAddError('The worksheet part is malformed.')
  return xml.slice(0, closeAt) + element + xml.slice(closeAt)
}

function areaToRef(area: TableArea): string {
  return (
    `${columnLabel(area.startColumn)}${area.startRow + 1}` +
    `:${columnLabel(area.endColumn)}${area.endRow + 1}`
  )
}

function parseRef(ref: string): TableArea {
  const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref.replace(/\$/g, ''))
  if (!match || !match[1] || !match[2]) {
    throw new TableAddError(`Unsupported range reference "${ref}" in the workbook.`)
  }
  const startColumn = labelToIndex(match[1])
  const startRow = Number(match[2]) - 1
  const endColumn = match[3] ? labelToIndex(match[3]) : startColumn
  const endRow = match[4] ? Number(match[4]) - 1 : startRow
  return { startRow, startColumn, endRow, endColumn }
}

function labelToIndex(label: string): number {
  let index = 0
  for (const char of label) index = index * 26 + (char.charCodeAt(0) - 64)
  return index - 1
}

function areasOverlap(a: TableArea, b: TableArea): boolean {
  return (
    a.startRow <= b.endRow &&
    b.startRow <= a.endRow &&
    a.startColumn <= b.endColumn &&
    b.startColumn <= a.endColumn
  )
}

function escapeAttribute(input: string): string {
  return input
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}
