import { columnLabel } from '../domain/cell-address'
import { isValidPivotFilter, type PivotFilterDef } from '../domain/pivot-filters'
import { formatPivotFormula, parsePivotFormula } from '../domain/pivot-formula'
import { groupLabel, isValidGrouping, type PivotFieldGrouping } from '../domain/pivot-grouping'
import {
  allocatePartPath,
  appendRelationship,
  registerContentTypeOverride,
  relativeTarget,
  relsPathFor,
  type MutablePackage,
} from './xlsx-drawing-add'
import { parseSheetElements } from './xlsx-sheets'

/// Persists pivots created in the editor as native OOXML parts: a cache
/// definition with real sharedItems + records derived from the source range,
/// and a pivotTableDefinition whose location covers the baked output grid.
/// Excel / LibreOffice open the file and see a fully populated pivot cache
/// without having to perform a refresh first.
///
/// The linkage is rels-only: workbook.xml gains a <pivotCaches> entry (via
/// the returned workbook XML — the caller owns that string), the workbook
/// rels point at the cache, the target worksheet rels point at the pivot
/// table part, and the pivot table rels point back at the cache.

export class PivotAddError extends Error {}

export interface PivotArea {
  readonly startRow: number
  readonly startColumn: number
  readonly endRow: number
  readonly endColumn: number
}

/// Grouping rule for one dimension field (fieldIndex points into fieldNames).
export type PivotAddGrouping = { readonly fieldIndex: number } & PivotFieldGrouping

export interface PivotValueSpec {
  /// Source field index; -1 for calculated fields (formula present).
  readonly fieldIndex: number
  readonly agg: 'sum' | 'count' | 'average' | 'max' | 'min'
  /// Optional Excel number format string for this data field (e.g. "#,##0.00")
  readonly numFmt?: string | undefined
  /// "Show values as" mode (ECMA-376 dataField@showDataAs): percent of grand
  /// total / row / column; undefined = normal.
  readonly showDataAs?: 'percentOfTotal' | 'percentOfRow' | 'percentOfCol' | undefined
  /// Calculated field: formula (references source field names, basic arithmetic)
  /// and the new field name. Written as a <cacheField formula=… databaseField="0">
  /// appended at the end of cacheFields, taking no cache records; agg is fixed to
  /// sum.
  readonly formula?: string | undefined
  readonly calcName?: string | undefined
}

/// One line of a multi-level row layout: a data row (t='data', members cover all
/// levels) or a subtotal row (t='default', members only hold the fixed prefix
/// levels). Each element of members is the member's index within that level's
/// rowLevelItems.
export interface PivotAddRowLine {
  readonly t: 'data' | 'default'
  readonly members: readonly number[]
}

export interface PivotAddition {
  /// Worksheet part that receives the pivot output.
  readonly worksheetPath: string
  /// File name of the source sheet (worksheetSource@sheet).
  readonly sourceSheetName: string
  readonly sourceArea: PivotArea
  /// Baked output area, headers and grand totals included.
  readonly location: PivotArea
  readonly name: string
  /// All source headers in column order — the cacheFields order.
  readonly fieldNames: readonly string[]
  /// Indices into fieldNames for the row dimension levels (outer → inner).
  readonly rowFieldIndices: readonly number[]
  readonly columnFieldIndex?: number | undefined
  /// Indices into fieldNames for report-filter (page) fields.
  readonly pageFieldIndices?: readonly number[] | undefined
  /// Deduplicated members of the level-0 (outermost) row field; with multiple
  /// levels this equals rowLevelItems[0].
  readonly rowItems: readonly string[]
  /// Deduplicated member lists per row level (i.e. each level's sharedItems
  /// order). When omitted, treated as single-level [rowItems]; required when
  /// there are ≥2 row levels.
  readonly rowLevelItems?: readonly (readonly string[])[] | undefined
  /// Row-by-row layout of the baked output data rows (excluding the header row
  /// and the trailing grand-total row). When omitted, treated as single-level
  /// rowItems expanded one per row; required with ≥2 row levels and must match
  /// the baked grid one-to-one.
  readonly rowLines?: readonly PivotAddRowLine[] | undefined
  readonly columnItems?: readonly string[] | undefined
  /// Multi-level columns: column dimension field indices (outer first). Takes
  /// precedence over columnFieldIndex when provided.
  readonly columnFieldIndices?: readonly number[] | undefined
  /// Deduplicated member lists per column level (required with ≥2 column levels,
  /// symmetric with rowLevelItems).
  readonly colLevelItems?: readonly (readonly string[])[] | undefined
  /// Column-by-column layout of the baked output data columns (excluding the
  /// trailing grand-total column; required with ≥2 column levels, symmetric with
  /// rowLines: data columns cover all levels, default columns are subtotals).
  readonly colLines?: readonly PivotAddRowLine[] | undefined
  /// Grouping rules for dimension fields (date/numeric ranges): the field's
  /// rowLevelItems / colLevelItems are already group labels; cache records are
  /// grouped by the same rule when written.
  readonly groupings?: readonly PivotAddGrouping[] | undefined
  /// Value/label filters: applied to row/column dimension fields; filtered-out
  /// members are marked hidden via rowHiddenItems/colHiddenItems, and layout
  /// lines exclude them.
  readonly filters?: readonly PivotFilterDef[] | undefined
  /// Filter-hidden member indices per row level (rowLevelItems order).
  readonly rowHiddenItems?: readonly (readonly number[])[] | undefined
  /// Filter-hidden member indices per column level (colLevelItems order).
  readonly colHiddenItems?: readonly (readonly number[])[] | undefined
  readonly values: readonly PivotValueSpec[]
}

/// Per-level row member lists: compatible with the legacy single-level rowItems form.
function rowLevelItemsOf(addition: PivotAddition): readonly (readonly string[])[] {
  return addition.rowLevelItems ?? [addition.rowItems]
}

/// Output-area row layout: compatible with the legacy single-level form (one data
/// row per rowItem).
function rowLinesOf(addition: PivotAddition): readonly PivotAddRowLine[] {
  return (
    addition.rowLines ??
    addition.rowItems.map((_, index) => ({ t: 'data' as const, members: [index] }))
  )
}

/// Column dimension field indices (outer first): compatible with the legacy
/// single-field columnFieldIndex form.
function colFieldIndicesOf(addition: PivotAddition): readonly number[] {
  if (addition.columnFieldIndices !== undefined) return addition.columnFieldIndices
  return addition.columnFieldIndex === undefined ? [] : [addition.columnFieldIndex]
}

/// Per-level column member lists: compatible with the legacy single-level
/// columnItems form.
function colLevelItemsOf(addition: PivotAddition): readonly (readonly string[])[] {
  if (addition.colLevelItems !== undefined) return addition.colLevelItems
  return addition.columnItems === undefined ? [] : [addition.columnItems]
}

/// Output-area column layout: compatible with the legacy single-level form (one
/// data column per columnItem).
function colLinesOf(addition: PivotAddition): readonly PivotAddRowLine[] {
  return (
    addition.colLines ??
    (addition.columnItems ?? []).map((_, index) => ({ t: 'data' as const, members: [index] }))
  )
}

/// Value specs in calculated-field form (formula present; fieldIndex is -1).
function calcValueSpecsOf(addition: PivotAddition): readonly PivotValueSpec[] {
  return addition.values.filter((value) => value.formula !== undefined)
}

const PIVOT_TABLE_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable'
const PIVOT_CACHE_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition'
const PIVOT_RECORDS_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords'
const PIVOT_TABLE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml'
const PIVOT_CACHE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml'
const PIVOT_RECORDS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml'
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/// The Excel data-field caption prefixes; must match the renderer's baked
/// header row so the pre-refresh grid and the definition agree.
const AGG_CAPTIONS: Record<PivotValueSpec['agg'], string> = {
  sum: 'Sum',
  count: 'Count',
  average: 'Average',
  max: 'Max',
  min: 'Min',
}
/// OOXML dataField@subtotal values; sum is the default and stays implicit.
const AGG_SUBTOTALS: Record<PivotValueSpec['agg'], string | null> = {
  sum: null,
  count: 'count',
  average: 'average',
  max: 'max',
  min: 'min',
}

/// Applies every pivot addition to the package and returns the workbook XML
/// with the new <pivotCaches> entries merged in.
export async function applyPivotAdditions(
  pkg: MutablePackage,
  additions: readonly PivotAddition[],
  workbookXml: string,
  touchedEntries: Set<string>,
): Promise<string> {
  if (additions.length === 0) return workbookXml
  const reserved = await collectExistingPivotNames(pkg)
  let nextCacheId = maxCacheId(workbookXml)
  let patchedWorkbookXml = workbookXml
  for (const addition of additions) {
    validateAddition(addition)
    const lowerName = addition.name.toLowerCase()
    if (reserved.has(lowerName)) {
      throw new PivotAddError(`Pivot name "${addition.name}" is already taken in this workbook.`)
    }
    reserved.add(lowerName)
    nextCacheId += 1

    // Read actual source data so the cache is pre-populated and LibreOffice /
    // WPS can display the correct values without a manual refresh.
    const sourceRows = await readSourceRows(pkg, addition)

    const recordsXml = buildCacheRecordsXml(sourceRows, addition)
    const recordCount = sourceRows.length

    const recordsPath = await allocatePartPath(pkg, 'xl/pivotCache/pivotCacheRecords', '.xml')
    pkg.add(recordsPath, recordsXml)
    touchedEntries.add(recordsPath)
    await registerContentTypeOverride(pkg, recordsPath, PIVOT_RECORDS_CONTENT_TYPE, touchedEntries)

    const cachePath = await allocatePartPath(pkg, 'xl/pivotCache/pivotCacheDefinition', '.xml')
    const recordsRelId = await appendRelationship(
      pkg,
      relsPathFor(cachePath),
      PIVOT_RECORDS_REL_TYPE,
      relativeTarget(cachePath, recordsPath),
    )
    touchedEntries.add(relsPathFor(cachePath))
    pkg.add(cachePath, buildCacheDefinitionXml(recordsRelId, addition, recordCount))
    touchedEntries.add(cachePath)
    await registerContentTypeOverride(pkg, cachePath, PIVOT_CACHE_CONTENT_TYPE, touchedEntries)

    const tablePath = await allocatePartPath(pkg, 'xl/pivotTables/pivotTable', '.xml')
    // The pivot table part's rels carry the cache linkage; the definition
    // itself references the cache only by cacheId.
    await appendRelationship(
      pkg,
      relsPathFor(tablePath),
      PIVOT_CACHE_REL_TYPE,
      relativeTarget(tablePath, cachePath),
    )
    touchedEntries.add(relsPathFor(tablePath))
    pkg.add(tablePath, buildPivotTableXml(nextCacheId, addition))
    touchedEntries.add(tablePath)
    await registerContentTypeOverride(pkg, tablePath, PIVOT_TABLE_CONTENT_TYPE, touchedEntries)

    await appendRelationship(
      pkg,
      relsPathFor(addition.worksheetPath),
      PIVOT_TABLE_REL_TYPE,
      relativeTarget(addition.worksheetPath, tablePath),
    )
    touchedEntries.add(relsPathFor(addition.worksheetPath))

    const workbookRelId = await appendRelationship(
      pkg,
      'xl/_rels/workbook.xml.rels',
      PIVOT_CACHE_REL_TYPE,
      relativeTarget('xl/workbook.xml', cachePath),
    )
    touchedEntries.add('xl/_rels/workbook.xml.rels')
    patchedWorkbookXml = addWorkbookPivotCache(patchedWorkbookXml, nextCacheId, workbookRelId)
  }
  return patchedWorkbookXml
}

// ─── Source data extraction ───────────────────────────────────────────────

/// One raw source row: element per cacheField, in fieldNames order.
/// Strings are the decoded text; numbers are JS numbers; null = blank.
type SourceValue = string | number | null

/// Read the shared-string table from the package (returns [] when absent).
async function readPivotSharedStrings(pkg: MutablePackage): Promise<readonly string[]> {
  if (!(await pkg.has('xl/sharedStrings.xml'))) return []
  const xml = await pkg.readText('xl/sharedStrings.xml')
  return [...xml.matchAll(/<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g)].map((m) =>
    [...(m[1] ?? '').matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)]
      .map((tm) => decodePivotXmlEntities(tm[1] ?? ''))
      .join(''),
  )
}

function decodePivotXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/// Resolve the worksheet path for a source sheet name. Attribute order and
/// entity encoding vary by producer, so the sheet is found by decoded name
/// rather than by pattern-matching serialized XML.
async function resolveSourcePath(pkg: MutablePackage, sheetName: string): Promise<string> {
  const workbookXml = await pkg.readText('xl/workbook.xml')
  const relationshipId = parseSheetElements(workbookXml).find(
    (element) => element.name === sheetName,
  )?.relationshipId
  if (relationshipId === undefined)
    throw new PivotAddError(`Source sheet "${sheetName}" not found in workbook.xml.`)

  const relXml = await pkg.readText('xl/_rels/workbook.xml.rels')
  // Two-step lookup: attribute order varies by producer (openpyxl puts Target before Id)
  const relationshipXml = new RegExp(
    `<Relationship\\b[^>]*\\bId="${relationshipId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/?>`,
  ).exec(relXml)?.[0]
  const relTarget =
    relationshipXml === undefined ? undefined : /\bTarget="([^"]+)"/.exec(relationshipXml)?.[1]
  if (!relTarget) throw new PivotAddError(`Relationship ${relationshipId} not found.`)
  const target = relTarget.replace(/^\/?xl\//, '')
  return `xl/${target.replace(/^\.\//, '')}`
}

/// Extract data rows (skipping the header row) from the source area.
/// Returns an array of SourceValue arrays, one per data row, in column order.
async function readSourceRows(
  pkg: MutablePackage,
  addition: PivotAddition,
): Promise<SourceValue[][]> {
  let worksheetXml: string
  try {
    const sourcePath = await resolveSourcePath(pkg, addition.sourceSheetName)
    worksheetXml = await pkg.readText(sourcePath)
  } catch {
    // Source sheet not accessible (e.g. test fixture without the sheet) →
    // fall back to empty records and let refreshOnLoad handle it at open time.
    return []
  }
  const sharedStrings = await readPivotSharedStrings(pkg)

  const { startRow, startColumn, endRow, endColumn } = addition.sourceArea
  // Row 0-indexed in PivotArea; row 0 is the header → data starts at row 1
  const rows: SourceValue[][] = []
  for (let r = startRow + 1; r <= endRow; r++) {
    const row: SourceValue[] = []
    for (let c = startColumn; c <= endColumn; c++) {
      const address = `${columnLabel(c)}${r + 1}`
      row.push(extractCellValue(worksheetXml, address, sharedStrings))
    }
    rows.push(row)
  }
  return rows
}

/// Extract the value from a single cell in worksheet XML.
function extractCellValue(
  worksheetXml: string,
  address: string,
  sharedStrings: readonly string[],
): SourceValue {
  const cellPattern = new RegExp(
    `<c\\b([^>]*)\\br="${address.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"([^>]*)(?:/>|>([\\s\\S]*?)</c>)`,
  )
  const match = cellPattern.exec(worksheetXml)
  if (!match) return null

  const attributes = `${match[1] ?? ''}${match[2] ?? ''}`
  const body = match[3] ?? ''
  const type = /\bt="([^"]+)"/.exec(attributes)?.[1]

  if (type === 'inlineStr') {
    return decodePivotXmlEntities(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/.exec(body)?.[1] ?? '')
  }
  const rawValue = /<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/.exec(body)?.[1]
  if (rawValue === undefined) return null
  if (type === 's') {
    const idx = Number(rawValue)
    return sharedStrings[idx] ?? null
  }
  if (type === 'str') return decodePivotXmlEntities(rawValue)
  if (type === 'b') return rawValue === '1' ? 1 : 0
  const n = Number(rawValue)
  return Number.isFinite(n) ? n : decodePivotXmlEntities(rawValue)
}

// ─── Cache records XML builder ───────────────────────────────────────────────

/// Build the pivotCacheRecords XML from raw source rows.
/// Each <r> element has one child per cacheField:
///   - dimension fields (rowFieldIndices, columnFieldIndex): <x v="N"/> (sharedItems index)
///   - numeric value fields: <n v="V"/>
///   - blank cells: <m/>
///   - string-type value fields (unusual): <s v="TEXT"/>
function buildCacheRecordsXml(
  sourceRows: readonly SourceValue[][],
  addition: PivotAddition,
): string {
  if (sourceRows.length === 0) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      `<pivotCacheRecords xmlns="${MAIN_NS}" xmlns:r="${REL_NS}" count="0"/>`
    )
  }

  // Build sharedItems index maps for dimension fields (row levels + column levels)
  const colFieldIndices = colFieldIndicesOf(addition)
  const dimensionFields = new Set<number>([...addition.rowFieldIndices, ...colFieldIndices])
  // pageFields are dimension fields too but sharedItems are empty; treat as non-index fields
  // (their records will use <s v="..."/> or <m/> like value fields)

  // Each row/column-level field has its own member table (independent per level
  // when multi-level); records write <x v="idx"/>.
  const sharedItemsByField = new Map<number, Map<string, number>>()
  const registerLevelItems = (
    fieldIndices: readonly number[],
    levelItems: readonly (readonly string[])[],
  ): void => {
    levelItems.forEach((items, level) => {
      const fieldIdx = fieldIndices[level]
      if (fieldIdx === undefined) return
      const lookup = new Map<string, number>()
      items.forEach((item, i) => lookup.set(String(item), i))
      sharedItemsByField.set(fieldIdx, lookup)
    })
  }
  registerLevelItems(addition.rowFieldIndices, rowLevelItemsOf(addition))
  registerLevelItems(colFieldIndices, colLevelItemsOf(addition))

  // Grouped fields' member tables hold group labels: raw values are grouped by
  // the same rule before records are written.
  const groupingByField = new Map<number, PivotFieldGrouping>()
  for (const { fieldIndex, ...grouping } of addition.groupings ?? []) {
    groupingByField.set(fieldIndex, grouping)
  }

  const records = sourceRows
    .map((row) => {
      const cells = row.map((value, colOffset) => {
        if (dimensionFields.has(colOffset)) {
          const lookup = sharedItemsByField.get(colOffset)!
          const grouping = groupingByField.get(colOffset)
          const key = grouping ? groupLabel(grouping, value) : value === null ? '' : String(value)
          const idx = lookup.get(key)
          if (idx === undefined) {
            return '<m/>'
          }
          return `<x v="${idx}"/>`
        }
        // Value field (or pageField — treat as raw value)
        if (value === null || value === '') return '<m/>'
        if (typeof value === 'number') return `<n v="${value}"/>`
        const n = Number(value)
        if (Number.isFinite(n)) return `<n v="${n}"/>`
        return `<s v="${escapeAttribute(String(value))}"/>`
      })
      return `<r>${cells.join('')}</r>`
    })
    .join('')

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<pivotCacheRecords xmlns="${MAIN_NS}" xmlns:r="${REL_NS}" count="${sourceRows.length}">` +
    records +
    '</pivotCacheRecords>'
  )
}

function validateAddition(addition: PivotAddition): void {
  const fieldCount = addition.fieldNames.length
  const colFieldIndices = colFieldIndicesOf(addition)
  const inBounds = (index: number): boolean => index >= 0 && index < fieldCount
  if (
    addition.rowFieldIndices.length === 0 ||
    addition.rowFieldIndices.some((i) => !inBounds(i)) ||
    colFieldIndices.some((i) => !inBounds(i)) ||
    addition.values.some((value) => value.formula === undefined && !inBounds(value.fieldIndex)) ||
    (addition.pageFieldIndices ?? []).some((i) => !inBounds(i))
  ) {
    throw new PivotAddError(`Pivot "${addition.name}" references a field outside its source.`)
  }
  // Calculated fields: name is required and must not collide with source fields
  // or other calculated fields (case-insensitive); the formula must parse against
  // source field names; aggregation is fixed to sum.
  const calcNames = new Set(addition.fieldNames.map((name) => name.trim().toLowerCase()))
  for (const value of calcValueSpecsOf(addition)) {
    const calcName = (value.calcName ?? '').trim()
    if (calcName === '' || calcNames.has(calcName.toLowerCase()) || value.agg !== 'sum') {
      throw new PivotAddError(`Pivot "${addition.name}" has an invalid calculated field.`)
    }
    calcNames.add(calcName.toLowerCase())
    try {
      parsePivotFormula(value.formula ?? '', addition.fieldNames)
    } catch (error) {
      throw new PivotAddError(
        `Pivot "${addition.name}" has a bad calculated-field formula: ` +
          `${error instanceof Error ? error.message : String(error)}`,
      )
    }
  }
  if (colFieldIndices.length > 0 && addition.values.length !== 1) {
    throw new PivotAddError(
      `Pivot "${addition.name}" with a column field supports exactly one values entry.`,
    )
  }
  // Multi-level axes must carry per-level member tables and a row-by-row
  // (column-by-column) layout, otherwise valid rowItems/colItems cannot be
  // generated; rows and columns share the same consistency checks.
  const validateAxisLines = (
    axis: 'row' | 'column',
    levels: number,
    levelItems: readonly (readonly string[])[] | undefined,
    lines: readonly PivotAddRowLine[] | undefined,
  ): void => {
    if (levels < 2) return
    if (!levelItems || levelItems.length !== levels || !lines) {
      throw new PivotAddError(
        `Pivot "${addition.name}" has ${levels} ${axis} levels but no per-level items/lines.`,
      )
    }
    for (const line of lines) {
      if (
        line.members.length === 0 ||
        line.members.length > levels ||
        (line.t === 'data' && line.members.length !== levels) ||
        line.members.some(
          (member, level) => member < 0 || member >= (levelItems[level]?.length ?? 0),
        )
      ) {
        throw new PivotAddError(`Pivot "${addition.name}" has an inconsistent ${axis} line.`)
      }
    }
  }
  const levels = addition.rowFieldIndices.length
  validateAxisLines('row', levels, addition.rowLevelItems, addition.rowLines)
  validateAxisLines('column', colFieldIndices.length, addition.colLevelItems, addition.colLines)
  // Grouping is only allowed on row/column dimension fields, and the rule itself
  // must be valid.
  const dimensionFields = new Set([...addition.rowFieldIndices, ...colFieldIndices])
  for (const { fieldIndex, ...grouping } of addition.groupings ?? []) {
    if (!dimensionFields.has(fieldIndex) || !isValidGrouping(grouping)) {
      throw new PivotAddError(`Pivot "${addition.name}" has an invalid field grouping.`)
    }
  }
  // Filters are only allowed on row/column dimension fields, at most one per
  // field; the data field a value filter references must exist; hidden member
  // indices must fall within the member table of the matching level.
  const filteredFields = new Set<number>()
  for (const filter of addition.filters ?? []) {
    if (
      !isValidPivotFilter(filter) ||
      !dimensionFields.has(filter.field) ||
      filteredFields.has(filter.field) ||
      (filter.kind === 'value' &&
        (filter.dataField < 0 || filter.dataField >= addition.values.length))
    ) {
      throw new PivotAddError(`Pivot "${addition.name}" has an invalid pivot filter.`)
    }
    filteredFields.add(filter.field)
  }
  const validateHidden = (
    hiddenPerLevel: readonly (readonly number[])[] | undefined,
    levelItems: readonly (readonly string[])[],
  ): boolean =>
    (hiddenPerLevel ?? []).every((hidden, level) =>
      hidden.every((member) => member >= 0 && member < (levelItems[level]?.length ?? 0)),
    )
  if (
    !validateHidden(addition.rowHiddenItems, rowLevelItemsOf(addition)) ||
    !validateHidden(addition.colHiddenItems, colLevelItemsOf(addition))
  ) {
    throw new PivotAddError(`Pivot "${addition.name}" has an out-of-range hidden item.`)
  }
  const sourceWidth = addition.sourceArea.endColumn - addition.sourceArea.startColumn + 1
  if (sourceWidth !== addition.fieldNames.length) {
    throw new PivotAddError(
      `Pivot "${addition.name}" has ${addition.fieldNames.length} field names for a ${sourceWidth}-column source.`,
    )
  }
}

async function collectExistingPivotNames(pkg: MutablePackage): Promise<Set<string>> {
  const names = new Set<string>()
  for (const path of await pkg.paths()) {
    if (!/^xl\/pivotTables\/[^/]+\.xml$/.test(path)) continue
    const name = /<pivotTableDefinition\b[^>]*\bname="([^"]+)"/.exec(await pkg.readText(path))?.[1]
    if (name) names.add(name.toLowerCase())
  }
  return names
}

function maxCacheId(workbookXml: string): number {
  let max = 0
  for (const match of workbookXml.matchAll(/<pivotCache\b[^>]*\bcacheId="(\d+)"/g)) {
    max = Math.max(max, Number(match[1]))
  }
  return max
}

/// Adds a <pivotCache> entry; workbook.xml schema order puts pivotCaches
/// after calcPr, before extLst.
function addWorkbookPivotCache(workbookXml: string, cacheId: number, relId: string): string {
  const entry = `<pivotCache cacheId="${cacheId}" r:id="${relId}"/>`
  if (workbookXml.includes('</pivotCaches>')) {
    return workbookXml.replace('</pivotCaches>', `${entry}</pivotCaches>`)
  }
  const element = `<pivotCaches>${entry}</pivotCaches>`
  const extLstAt = workbookXml.search(/<extLst\b/)
  if (extLstAt >= 0) {
    return workbookXml.slice(0, extLstAt) + element + workbookXml.slice(extLstAt)
  }
  const closeAt = workbookXml.lastIndexOf('</workbook>')
  if (closeAt < 0) throw new PivotAddError('xl/workbook.xml is malformed.')
  return workbookXml.slice(0, closeAt) + element + workbookXml.slice(closeAt)
}

/// Exported for pivot layout edits (A3): the cacheDefinition is regenerated
/// with the edited dimension fields' sharedItems, keeping the records rel id.
export function buildCacheDefinitionXml(
  recordsRelId: string,
  addition: PivotAddition,
  recordCount: number,
): string {
  // All dimension fields (row levels + column levels + page) get sharedItems lists.
  const colFieldIndices = colFieldIndicesOf(addition)
  const dimensionFieldSet = new Set<number>([
    ...addition.rowFieldIndices,
    ...colFieldIndices,
    ...(addition.pageFieldIndices ?? []),
  ])
  // Row/column fields have one member table per level (cacheField sharedItems
  // order = index within the level).
  const levelItemsByField = new Map<number, readonly string[]>()
  addition.rowFieldIndices.forEach((fieldIdx, level) => {
    levelItemsByField.set(fieldIdx, rowLevelItemsOf(addition)[level] ?? [])
  })
  colFieldIndices.forEach((fieldIdx, level) => {
    levelItemsByField.set(fieldIdx, colLevelItemsOf(addition)[level] ?? [])
  })

  const fields = addition.fieldNames
    .map((name, index) => {
      if (!dimensionFieldSet.has(index)) {
        return `<cacheField name="${escapeAttribute(name)}" numFmtId="0"><sharedItems/></cacheField>`
      }
      // Determine sharedItems for this dimension field;
      // pageField-only fields need no items (shown as filter, no sharedItems required)
      const items: readonly string[] = levelItemsByField.get(index) ?? []
      const sharedItems =
        items.length === 0
          ? '<sharedItems/>'
          : `<sharedItems count="${items.length}">` +
            items.map((item) => `<s v="${escapeAttribute(item)}"/>`).join('') +
            '</sharedItems>'
      return `<cacheField name="${escapeAttribute(name)}" numFmtId="0">${sharedItems}</cacheField>`
    })
    .join('')
  // Calculated fields are appended after all source fields: databaseField="0",
  // no cache records; formulas are normalized to the quoted-reference form (see
  // pivot-formula).
  const calcFields = calcValueSpecsOf(addition)
    .map(
      (spec) =>
        `<cacheField name="${escapeAttribute(spec.calcName ?? '')}" ` +
        `formula="${escapeAttribute(formatPivotFormula(parsePivotFormula(spec.formula ?? '', addition.fieldNames)))}" ` +
        'databaseField="0" numFmtId="0"/>',
    )
    .join('')
  const cacheFieldCount = addition.fieldNames.length + calcValueSpecsOf(addition).length
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<pivotCacheDefinition xmlns="${MAIN_NS}" xmlns:r="${REL_NS}" r:id="${recordsRelId}" ` +
    `refreshedBy="genoffice" createdVersion="8" refreshedVersion="8" ` +
    `minRefreshableVersion="3" recordCount="${recordCount}">` +
    '<cacheSource type="worksheet">' +
    `<worksheetSource ref="${areaToRef(addition.sourceArea)}" ` +
    `sheet="${escapeAttribute(addition.sourceSheetName)}"/>` +
    '</cacheSource>' +
    `<cacheFields count="${cacheFieldCount}">${fields}${calcFields}</cacheFields>` +
    '</pivotCacheDefinition>'
  )
}

/// Exported for pivot layout edits (A3): the pivotTableDefinition is
/// regenerated from the edited layout, keeping the original name and cacheId.
export function buildPivotTableXml(cacheId: number, addition: PivotAddition): string {
  const rowLevelItems = rowLevelItemsOf(addition)
  const levels = addition.rowFieldIndices.length
  const colFieldIndices = colFieldIndicesOf(addition)
  const colLevelItems = colLevelItemsOf(addition)
  const colLevels = colFieldIndices.length
  // Multi-level rows use tabular layout (compact/outline off) so each level takes
  // one column, matching the baked grid's N label columns; single level keeps the
  // existing compact output.
  const tabular = levels >= 2
  const levelByField = new Map<number, number>()
  addition.rowFieldIndices.forEach((fieldIdx, level) => levelByField.set(fieldIdx, level))
  const colLevelByField = new Map<number, number>()
  colFieldIndices.forEach((fieldIdx, level) => colLevelByField.set(fieldIdx, level))
  const pageFieldSet = new Set(addition.pageFieldIndices ?? [])

  const pivotFields = addition.fieldNames
    .map((_, index) => {
      const level = levelByField.get(index)
      if (level !== undefined) {
        return axisPivotField(
          'axisRow',
          (rowLevelItems[level] ?? []).length,
          tabular,
          new Set(addition.rowHiddenItems?.[level] ?? []),
        )
      }
      const colLevel = colLevelByField.get(index)
      if (colLevel !== undefined) {
        return axisPivotField(
          'axisCol',
          (colLevelItems[colLevel] ?? []).length,
          tabular,
          new Set(addition.colHiddenItems?.[colLevel] ?? []),
        )
      }
      if (addition.values.some((value) => value.fieldIndex === index)) {
        return `<pivotField dataField="1" showAll="0"${tabular ? ' compact="0" outline="0"' : ''}/>`
      }
      if (pageFieldSet.has(index)) {
        return '<pivotField axis="axisPage" showAll="0"><items count="1"><item t="default"/></items></pivotField>'
      }
      return `<pivotField showAll="0"${tabular ? ' compact="0" outline="0"' : ''}/>`
    })
    .join('')
  // Calculated fields follow the same append order as cacheFields, each taking a
  // value-only pivotField.
  const calcSpecs = calcValueSpecsOf(addition)
  const calcPivotFields = calcSpecs
    .map(
      () => `<pivotField dataField="1" showAll="0"${tabular ? ' compact="0" outline="0"' : ''}/>`,
    )
    .join('')
  const pivotFieldCount = addition.fieldNames.length + calcSpecs.length

  // rowFields: all row dimension indices in order
  const rowFieldsXml =
    `<rowFields count="${addition.rowFieldIndices.length}">` +
    addition.rowFieldIndices.map((idx) => `<field x="${idx}"/>`).join('') +
    '</rowFields>'

  // rowItems: generated from the baked grid's row-by-row layout (with per-level
  // members and subtotal rows when multi-level), appending the grand-total row.
  const rowItems = buildAxisItemsXml(rowLinesOf(addition), levels, 'rowItems')

  // pageFields section (report filters)
  const pageFieldsXml =
    (addition.pageFieldIndices ?? []).length > 0
      ? `<pageFields count="${(addition.pageFieldIndices ?? []).length}">` +
        (addition.pageFieldIndices ?? [])
          .map((idx) => `<pageField fld="${idx}" hier="-1"/>`)
          .join('') +
        '</pageFields>'
      : ''

  // Column axis: symmetric with the row axis — colLines (data/subtotal columns)
  // are encoded with r (repeat) into <colItems>, appending the grand-total column;
  // single level stays compatible with the legacy columnItems form.
  let columnPart: string
  if (colFieldIndices.length > 0) {
    columnPart =
      `<colFields count="${colFieldIndices.length}">` +
      colFieldIndices.map((index) => `<field x="${index}"/>`).join('') +
      '</colFields>' +
      buildAxisItemsXml(colLinesOf(addition), colLevels, 'colItems')
  } else if (addition.values.length > 1) {
    columnPart =
      '<colFields count="1"><field x="-2"/></colFields>' +
      `<colItems count="${addition.values.length}">` +
      addition.values
        .map((_, index) => (index === 0 ? '<i><x/></i>' : `<i i="${index}"><x v="${index}"/></i>`))
        .join('') +
      '</colItems>'
  } else {
    columnPart = '<colItems count="1"><i/></colItems>'
  }

  const dataFields = addition.values
    .map((value) => {
      // A calculated field's fld points at the cacheField appended after the source
      // fields; aggregation is fixed to sum.
      const isCalc = value.formula !== undefined
      const field = isCalc ? (value.calcName ?? '') : (addition.fieldNames[value.fieldIndex] ?? '')
      const fld = isCalc ? addition.fieldNames.length + calcSpecs.indexOf(value) : value.fieldIndex
      const caption = `${AGG_CAPTIONS[value.agg]} of ${field}`
      const subtotal = isCalc ? null : AGG_SUBTOTALS[value.agg]
      // numFmtId: 0 = General; map common patterns to standard Excel IDs.
      // Percent display modes default to 0.00% (id 10) when no format is given.
      const numFmtId = value.numFmt
        ? resolveNumFmtId(value.numFmt)
        : value.showDataAs !== undefined
          ? 10
          : 0
      // These three percent modes reference no base field/item; baseField/baseItem
      // keep their defaults.
      return (
        `<dataField name="${escapeAttribute(caption)}" fld="${fld}"` +
        `${subtotal === null ? '' : ` subtotal="${subtotal}"`}` +
        `${value.showDataAs === undefined ? '' : ` showDataAs="${value.showDataAs}"`}` +
        ' baseField="0" baseItem="0"' +
        `${numFmtId !== 0 ? ` numFmtId="${numFmtId}"` : ''}/>`
      )
    })
    .join('')

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<pivotTableDefinition xmlns="${MAIN_NS}" name="${escapeAttribute(addition.name)}" ` +
    `cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" ` +
    'applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" ' +
    'dataCaption="Values" updatedVersion="8" createdVersion="8" minRefreshableVersion="3" ' +
    'useAutoFormatting="1" itemPrintTitles="1" ' +
    (tabular
      ? 'compact="0" compactData="0" outline="0" outlineData="0" '
      : 'outline="1" outlineData="1" ') +
    'multipleFieldFilters="0">' +
    // With multi-level columns each column level takes one header row and the
    // data area starts at row colLevels; under tabular layout each row level takes
    // one column and the data area starts at column levels.
    `<location ref="${areaToRef(addition.location)}" firstHeaderRow="1" ` +
    `firstDataRow="${Math.max(1, colLevels)}" firstDataCol="${levels}"/>` +
    `<pivotFields count="${pivotFieldCount}">${pivotFields}${calcPivotFields}</pivotFields>` +
    rowFieldsXml +
    rowItems +
    pageFieldsXml +
    columnPart +
    `<dataFields count="${addition.values.length}">${dataFields}</dataFields>` +
    '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" ' +
    'showRowStripes="0" showColStripes="0" showLastColumn="1"/>' +
    buildFiltersXml(addition) +
    buildGroupingExtLst(addition) +
    '</pivotTableDefinition>'
  )
}

/// <filters> block for value/label filters (ECMA-376 CT_PivotFilters): label
/// filters write captionEqual/captionContains/captionBeginsWith + stringValue1;
/// value filters write count (top10)/valueGreaterThan/valueBetween + iMeasureFld;
/// each filter embeds the autoFilter (top10/customFilters) required by the spec.
/// Filtered-out members are already marked h="1" in pivotField items, so Excel
/// shows the filtered result on open.
function buildFiltersXml(addition: PivotAddition): string {
  const filters = addition.filters ?? []
  if (filters.length === 0) return ''
  const entries = filters.map((filter, index) => {
    const id = index + 1
    if (filter.kind === 'label') {
      const type =
        filter.op === 'equal'
          ? 'captionEqual'
          : filter.op === 'contains'
            ? 'captionContains'
            : 'captionBeginsWith'
      // The autoFilter match value expresses contains/begins-with via wildcards
      // (same notation Excel uses).
      const wildcard =
        filter.op === 'equal'
          ? filter.value
          : filter.op === 'contains'
            ? `*${filter.value}*`
            : `${filter.value}*`
      return (
        `<filter fld="${filter.field}" type="${type}" evalOrder="-1" id="${id}" ` +
        `stringValue1="${escapeAttribute(filter.value)}">` +
        '<autoFilter ref="A1"><filterColumn colId="0"><customFilters>' +
        `<customFilter val="${escapeAttribute(wildcard)}"/>` +
        '</customFilters></filterColumn></autoFilter></filter>'
      )
    }
    if (filter.op === 'top') {
      return (
        `<filter fld="${filter.field}" type="count" evalOrder="-1" id="${id}" ` +
        `iMeasureFld="${filter.dataField}">` +
        `<autoFilter ref="A1"><filterColumn colId="0"><top10 val="${filter.count ?? 10}"/>` +
        '</filterColumn></autoFilter></filter>'
      )
    }
    if (filter.op === 'greaterThan') {
      return (
        `<filter fld="${filter.field}" type="valueGreaterThan" evalOrder="-1" id="${id}" ` +
        `iMeasureFld="${filter.dataField}">` +
        '<autoFilter ref="A1"><filterColumn colId="0"><customFilters>' +
        `<customFilter operator="greaterThan" val="${filter.from ?? 0}"/>` +
        '</customFilters></filterColumn></autoFilter></filter>'
      )
    }
    return (
      `<filter fld="${filter.field}" type="valueBetween" evalOrder="-1" id="${id}" ` +
      `iMeasureFld="${filter.dataField}">` +
      '<autoFilter ref="A1"><filterColumn colId="0"><customFilters and="1">' +
      `<customFilter operator="greaterThanOrEqual" val="${filter.from ?? 0}"/>` +
      `<customFilter operator="lessThanOrEqual" val="${filter.to ?? 0}"/>` +
      '</customFilters></filterColumn></autoFilter></filter>'
    )
  })
  return `<filters count="${entries.length}">${entries.join('')}</filters>`
}

/// Private extLst extension for grouped-field metadata: Excel ignores unknown
/// ext elements (the file stays valid), and our parser reads it back so refresh
/// still groups accordingly.
/// TODO(OOXML persistence): emit real <fieldGroup> + <rangePr> + <groupItems> so
/// refreshing inside Excel also keeps the grouping (currently a manual Excel
/// refresh degrades to grouping by raw values).
function buildGroupingExtLst(addition: PivotAddition): string {
  const groupings = addition.groupings ?? []
  if (groupings.length === 0) return ''
  return (
    '<extLst><ext uri="{AIO-PIVOT-GROUPINGS}" xmlns:aio="urn:aioffice:pivot">' +
    `<aio:aioPivotGroupings v="${escapeAttribute(JSON.stringify(groupings))}"/>` +
    '</ext></extLst>'
  )
}

/// Map common Excel number format strings to built-in numFmtId values.
/// Returns 0 (General) for unrecognised patterns.
function resolveNumFmtId(numFmt: string): number {
  const fmt = numFmt.trim()
  const knownFormats: Record<string, number> = {
    '0': 1,
    '0.00': 2,
    '#,##0': 3,
    '#,##0.00': 4,
    '0%': 9,
    '0.00%': 10,
    '0.00E+00': 11,
    '# ?/?': 12,
    '# ??/??': 13,
    'mm-dd-yy': 14,
    'm/d/yyyy': 14,
    'd-mmm-yy': 15,
    'h:mm AM/PM': 18,
    'h:mm:ss AM/PM': 19,
    'h:mm': 20,
    'h:mm:ss': 21,
    '#,##0.00;(#,##0.00)': 4,
    '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)': 164,
  }
  return knownFormats[fmt] ?? 0
}

function axisPivotField(
  axis: 'axisRow' | 'axisCol',
  itemCount: number,
  tabular = false,
  hidden?: ReadonlySet<number>,
): string {
  // Filtered-out members are written as hidden entries (h="1"), which is how
  // an applied filter is persisted in the file.
  const items = Array.from(
    { length: itemCount },
    (_, index) => `<item x="${index}"${hidden?.has(index) ? ' h="1"' : ''}/>`,
  ).join('')
  return (
    `<pivotField axis="${axis}" showAll="0"${tabular ? ' compact="0" outline="0"' : ''}>` +
    `<items count="${itemCount + 1}">${items}<item t="default"/></items>` +
    '</pivotField>'
  )
}

/// Generates <rowItems>/<colItems> per OOXML r (repeat) semantics: members within
/// the common prefix shared with the previous row (column) are omitted, and r
/// points at the level of this line's first explicit <x>; subtotal lines are
/// written t="default". Mirrors the carry rules of parseLayoutLines on the
/// parsing side.
function buildAxisItemsXml(
  lines: readonly PivotAddRowLine[],
  levels: number,
  tag: 'rowItems' | 'colItems',
): string {
  const parts: string[] = []
  // Simulate the parser's carry state: data lines fill all levels; default lines
  // clear the levels past their depth.
  let carried: (number | null)[] = new Array(levels).fill(null)
  for (const line of lines) {
    const fixed = line.members.length
    let repeat = 0
    // Each line must emit at least one <x> as an anchor, so the cap is fixed-1.
    while (
      repeat < fixed - 1 &&
      carried[repeat] !== null &&
      carried[repeat] === line.members[repeat]
    ) {
      repeat += 1
    }
    const xs = line.members
      .slice(repeat)
      .map((member) => (member === 0 ? '<x/>' : `<x v="${member}"/>`))
      .join('')
    parts.push(
      `<i${repeat === 0 ? '' : ` r="${repeat}"`}${line.t === 'data' ? '' : ` t="${line.t}"`}>${xs}</i>`,
    )
    // Data lines have fixed === levels; default lines reset levels past their
    // depth to the unfixed state.
    carried = [...line.members]
    while (carried.length < levels) carried.push(null)
  }
  parts.push('<i t="grand"><x/></i>')
  return `<${tag} count="${lines.length + 1}">${parts.join('')}</${tag}>`
}

function areaToRef(area: PivotArea): string {
  return (
    `${columnLabel(area.startColumn)}${area.startRow + 1}` +
    `:${columnLabel(area.endColumn)}${area.endRow + 1}`
  )
}

function escapeAttribute(input: string): string {
  return input
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}
