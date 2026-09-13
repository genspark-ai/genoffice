/// PivotTable definition reader: parses a pivotTable part plus its
/// pivotCacheDefinition into the layout/fields model the recompute engine
/// consumes. Anything outside the supported envelope lands in `unsupported`
/// (with a reason) instead of producing a silently wrong refresh.

import type { PivotFilterDef } from '../domain/pivot-filters'
import { parsePivotFormula } from '../domain/pivot-formula'
import { isValidGrouping, type PivotFieldGrouping } from '../domain/pivot-grouping'

export class PivotParseError extends Error {}

export type { PivotFieldGrouping } from '../domain/pivot-grouping'
export type { PivotFilterDef, PivotLabelFilter, PivotValueFilter } from '../domain/pivot-filters'

export type PivotSharedItem = string | number | boolean | null

export interface PivotFieldItem {
  /// sharedItems index; null for subtotal/function items.
  readonly x: number | null
  readonly hidden: boolean
}

export interface PivotCacheField {
  readonly name: string
  readonly sharedItems: readonly PivotSharedItem[]
  /// Grouped field (date/numeric ranges): when present, sharedItems are group
  /// labels, and refresh groups raw source values before matching. Stored in our
  /// private pivotTable extLst extension (real <fieldGroup> OOXML persistence: see
  /// the TODO in pivot-grouping).
  readonly grouping?: PivotFieldGrouping | undefined
  /// Calculated field (cacheField@formula): not a source data column, takes no
  /// cache records; value is computed by the formula over the grouped aggregates
  /// of other fields (see pivot-formula).
  readonly formula?: string | undefined
}

export interface PivotLayoutLine {
  /// 'data' rows/cols carry members; 'default' = subtotal, 'grand' = total.
  readonly t: string
  /// members fixed by this line: pivotField item index per axis level,
  /// already propagated across the r-attribute repeats.
  readonly members: readonly (number | null)[]
  /// how many leading axis levels the line fixes (subtotal depth).
  readonly depth: number
  /// data-field index this line renders (multi-value pivots).
  readonly dataField: number
}

/// "Show values as" modes that support refresh (subset of the ECMA-376
/// dataField@showDataAs enum): percent of grand total / row total / column total.
export type PivotShowDataAs = 'percentOfTotal' | 'percentOfRow' | 'percentOfCol'

export interface PivotDataField {
  readonly name: string
  readonly field: number
  readonly subtotal: string
  /// "Show values as" mode; undefined = normal (show the aggregate directly).
  readonly showDataAs?: PivotShowDataAs | undefined
  /// Formula of a calculated field (present when field points at a cacheField
  /// with a formula): each group first SUMs the referenced source fields, then
  /// evaluates the formula.
  readonly formula?: string | undefined
}

export interface PivotDefinition {
  readonly outputRef: string
  readonly firstDataRow: number
  readonly firstDataCol: number
  readonly fields: readonly PivotCacheField[]
  /// per cache field: axis items (empty for pure data fields).
  readonly fieldItems: readonly (readonly PivotFieldItem[])[]
  /// cache-field indexes; -2 marks the data-field position.
  readonly rowFields: readonly number[]
  readonly colFields: readonly number[]
  readonly rowLines: readonly PivotLayoutLine[]
  readonly colLines: readonly PivotLayoutLine[]
  readonly dataFields: readonly PivotDataField[]
  /// page (report-filter) selections: sharedItems index or null for "all".
  readonly pageFields: readonly { readonly field: number; readonly item: number | null }[]
  /// Value/label filters (pivotFilters): members that fail the filter are hidden
  /// entries in fieldItems; refresh re-applies them against the source data
  /// (value filters aggregate first, then filter).
  readonly filters: readonly PivotFilterDef[]
  readonly sourceSheet: string
  readonly sourceRef: string
  /// reasons this pivot cannot be recomputed; empty = refresh supported.
  readonly unsupported: readonly string[]
}

const SUPPORTED_SUBTOTALS = new Set([
  'sum',
  'count',
  'countNums',
  'average',
  'max',
  'min',
  'product',
])

const SUPPORTED_SHOW_DATA_AS = new Set<string>(['percentOfTotal', 'percentOfRow', 'percentOfCol'])

function attr(element: string, name: string): string | null {
  const found = new RegExp(` ${name}="([^"]*)"`).exec(element)
  return found?.[1] ?? null
}

function decodeEntities(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/// All <tag …>…</tag> or <tag …/> occurrences at any depth (regex-scoped —
/// fine for pivot parts, which never nest a tag inside itself). The self-closing
/// branch is matched first on its own: otherwise, when <tag …/> is followed by a
/// paired tag of the same name, the `>[\s\S]*?` branch would swallow both elements
/// as one (mixing both forms inside filters triggers this).
function elements(xml: string, tag: string): string[] {
  const pattern = new RegExp(`<${tag}(?:\\s[^>]*)?/>|<${tag}(?:\\s[^>]*)?>[\\s\\S]*?</${tag}>`, 'g')
  return [...xml.matchAll(pattern)].map((match) => match[0])
}

function firstElement(xml: string, tag: string): string | null {
  return elements(xml, tag)[0] ?? null
}

function parseSharedItems(cacheFieldXml: string): PivotSharedItem[] {
  const container = firstElement(cacheFieldXml, 'sharedItems')
  if (!container) return []
  const items: PivotSharedItem[] = []
  const pattern = /<(s|n|b|d|m|e)(\s[^>]*)?\/?>(?:<\/\1>)?/g
  for (const match of container.matchAll(pattern)) {
    const kind = match[1]
    const value = attr(match[0], 'v')
    if (kind === 's') items.push(decodeEntities(value ?? ''))
    else if (kind === 'n') items.push(Number(value ?? '0'))
    else if (kind === 'b') items.push(value === '1' || value === 'true')
    else if (kind === 'd') items.push(decodeEntities(value ?? ''))
    else if (kind === 'm') items.push(null)
    else items.push(decodeEntities(value ?? '#ERROR'))
  }
  return items
}

function parseLayoutLines(containerXml: string | null, axisFieldCount: number): PivotLayoutLine[] {
  if (!containerXml) {
    // No axis fields: the implicit single line covers everything.
    return [{ t: 'data', members: [], depth: 0, dataField: 0 }]
  }
  const lines: PivotLayoutLine[] = []
  const current: (number | null)[] = new Array(axisFieldCount).fill(null)
  let currentDataField = 0
  for (const line of elements(containerXml, 'i')) {
    const t = attr(line, 't') ?? 'data'
    const repeat = Number(attr(line, 'r') ?? '0')
    const values = elements(line, 'x').map((x) => Number(attr(x, 'v') ?? '0'))
    values.forEach((value, offset) => {
      current[repeat + offset] = value
    })
    const depth = repeat + values.length
    // Levels past the line's depth are stale carries; a data line fixes all.
    for (let level = depth; level < axisFieldCount; level += 1) {
      if (t !== 'data') current[level] = null
    }
    const explicitDataField = attr(line, 'i')
    if (explicitDataField !== null) currentDataField = Number(explicitDataField)
    lines.push({
      t,
      members: [...current],
      depth: t === 'grand' ? 0 : depth,
      dataField: currentDataField,
    })
  }
  return lines
}

/// Marks a pivotCacheDefinition so Excel rebuilds the cache on open — the
/// recomputed output cells and the stale cacheRecords then re-converge.
export function setPivotRefreshOnLoad(cacheDefinitionXml: string): string {
  const root = /<pivotCacheDefinition\b[^>]*>/.exec(cacheDefinitionXml)
  if (!root) throw new PivotParseError('pivotCacheDefinition part has no root element.')
  const updated = root[0].includes(' refreshOnLoad="')
    ? root[0].replace(/ refreshOnLoad="[^"]*"/, ' refreshOnLoad="1"')
    : root[0].replace('<pivotCacheDefinition', '<pivotCacheDefinition refreshOnLoad="1"')
  return (
    cacheDefinitionXml.slice(0, root.index) +
    updated +
    cacheDefinitionXml.slice(root.index + root[0].length)
  )
}

export function parsePivotDefinition(
  pivotTableXml: string,
  cacheDefinitionXml: string,
): PivotDefinition {
  const unsupported: string[] = []

  const location = firstElement(pivotTableXml, 'location')
  if (!location) throw new PivotParseError('pivotTable part has no location element.')
  const outputRef = attr(location, 'ref')
  if (!outputRef) throw new PivotParseError('pivotTable location has no ref.')
  const firstDataRow = Number(attr(location, 'firstDataRow') ?? '1')
  const firstDataCol = Number(attr(location, 'firstDataCol') ?? '1')

  const worksheetSource = firstElement(cacheDefinitionXml, 'worksheetSource')
  if (!worksheetSource) {
    unsupported.push('the data source is not a worksheet range (external/consolidated source)')
  }
  const sourceSheet = decodeEntities(attr(worksheetSource ?? '', 'sheet') ?? '')
  const sourceRef = attr(worksheetSource ?? '', 'ref') ?? ''
  if (worksheetSource && (!sourceSheet || !sourceRef)) {
    unsupported.push('the data source uses a defined name, recompute not supported yet')
  }

  // Our own grouped-field metadata: stored in a private pivotTable extLst
  // extension (fieldIndex → grouping rule). Corrupt/invalid metadata fails closed.
  const groupingByField = new Map<number, PivotFieldGrouping>()
  const groupingExt = /<(?:\w+:)?aioPivotGroupings\b[^>]*\bv="([^"]*)"/.exec(pivotTableXml)
  if (groupingExt) {
    try {
      const entries = JSON.parse(decodeEntities(groupingExt[1] ?? '')) as unknown
      if (!Array.isArray(entries)) throw new Error('not an array')
      for (const entry of entries) {
        const { fieldIndex, ...grouping } = entry as { fieldIndex?: unknown }
        if (typeof fieldIndex !== 'number' || !isValidGrouping(grouping)) {
          throw new Error('bad grouping entry')
        }
        groupingByField.set(fieldIndex, grouping)
      }
    } catch {
      unsupported.push('grouped-field extension metadata cannot be parsed')
    }
  }

  const cacheFieldsBlock = firstElement(cacheDefinitionXml, 'cacheFields')
  const cacheFields = cacheFieldsBlock === null ? [] : elements(cacheFieldsBlock, 'cacheField')
  // Calculated fields (cacheField@formula): strip the optional leading = and
  // validate against source field names.
  const formulaByField = new Map<number, string>()
  cacheFields.forEach((fieldXml, index) => {
    const formula = attr(fieldXml, 'formula')
    if (formula !== null) formulaByField.set(index, decodeEntities(formula).replace(/^\s*=/, ''))
  })
  const fields: PivotCacheField[] = cacheFields.map((fieldXml, index) => ({
    name: decodeEntities(attr(fieldXml, 'name') ?? ''),
    sharedItems: parseSharedItems(fieldXml),
    ...(groupingByField.has(index) ? { grouping: groupingByField.get(index)! } : {}),
    ...(formulaByField.has(index) ? { formula: formulaByField.get(index)! } : {}),
  }))
  // Calculated-field formulas must parse and may only reference non-calculated
  // source fields (chained references are unsupported).
  const sourceFieldNames = fields
    .filter((_, index) => !formulaByField.has(index))
    .map((field) => field.name)
  for (const [index, formula] of formulaByField) {
    // The refresh engine assumes cache-field index == source column index:
    // calculated fields must come after all source fields (Excel also appends
    // them at the end); otherwise fail closed.
    if (index < sourceFieldNames.length) {
      unsupported.push(
        `calculated field "${fields[index]?.name}" does not come after the source fields, recompute not supported`,
      )
      continue
    }
    try {
      parsePivotFormula(formula, sourceFieldNames)
    } catch (error) {
      unsupported.push(
        `the formula of calculated field "${fields[index]?.name}" does not support recompute` +
          ` (${error instanceof Error ? error.message : String(error)})`,
      )
    }
  }
  cacheFields.forEach((fieldXml, index) => {
    // <fieldGroup> grouping from external (Excel-created) files is still
    // unsupported: we cannot reliably rebuild the semantics of its groupItems and
    // par/base chain; our own grouping goes through extLst metadata.
    if (firstElement(fieldXml, 'fieldGroup') !== null) {
      unsupported.push(`grouped field "${fields[index]?.name}" does not support recompute`)
    }
  })

  const pivotFieldsBlock = firstElement(pivotTableXml, 'pivotFields')
  const pivotFields = pivotFieldsBlock === null ? [] : elements(pivotFieldsBlock, 'pivotField')
  const fieldItems: PivotFieldItem[][] = pivotFields.map((fieldXml) => {
    const itemsBlock = firstElement(fieldXml, 'items')
    if (!itemsBlock) return []
    return elements(itemsBlock, 'item')
      .filter((item) => attr(item, 't') === null)
      .map((item) => ({
        x: attr(item, 'x') === null ? null : Number(attr(item, 'x')),
        hidden: attr(item, 'h') === '1',
      }))
  })

  const readAxis = (tag: string): number[] => {
    const block = firstElement(pivotTableXml, tag)
    if (!block) return []
    return elements(block, 'field').map((field) => Number(attr(field, 'x') ?? '0'))
  }
  const rowFields = readAxis('rowFields')
  const colFields = readAxis('colFields')

  const dataFieldsBlock = firstElement(pivotTableXml, 'dataFields')
  const dataFields: PivotDataField[] = (
    dataFieldsBlock === null ? [] : elements(dataFieldsBlock, 'dataField')
  ).map((fieldXml) => {
    // showDataAs="normal" (or absent) = normal display, not recorded; unsupported
    // display modes only record an unsupported reason, and the mode itself is not
    // carried over (the IPC schema only accepts supported enum values).
    const showDataAs = attr(fieldXml, 'showDataAs')
    if (showDataAs !== null && showDataAs !== 'normal' && !SUPPORTED_SHOW_DATA_AS.has(showDataAs)) {
      unsupported.push(`"show values as" mode "${showDataAs}" does not support recompute`)
    }
    const field = Number(attr(fieldXml, 'fld') ?? '0')
    return {
      name: decodeEntities(attr(fieldXml, 'name') ?? ''),
      field,
      subtotal: attr(fieldXml, 'subtotal') ?? 'sum',
      ...(showDataAs !== null && SUPPORTED_SHOW_DATA_AS.has(showDataAs)
        ? { showDataAs: showDataAs as PivotShowDataAs }
        : {}),
      // Data fields pointing at a calculated field carry the formula; refresh
      // evaluates it.
      ...(formulaByField.has(field) ? { formula: formulaByField.get(field)! } : {}),
    }
  })
  for (const dataField of dataFields) {
    if (!SUPPORTED_SUBTOTALS.has(dataField.subtotal)) {
      unsupported.push(`aggregation "${dataField.subtotal}" does not support recompute`)
    }
  }
  if (dataFields.length === 0) unsupported.push('this pivot table has no data fields')

  const pageFieldsBlock = firstElement(pivotTableXml, 'pageFields')
  const pageFields = (pageFieldsBlock === null ? [] : elements(pageFieldsBlock, 'pageField')).map(
    (fieldXml) => {
      const item = attr(fieldXml, 'item')
      return {
        field: Number(attr(fieldXml, 'fld') ?? '0'),
        item: item === null ? null : Number(item),
      }
    },
  )

  // Calculated fields may only be data fields: used as a row/column/report-filter
  // dimension they cannot be grouped by formula.
  for (const [index] of formulaByField) {
    if (
      rowFields.includes(index) ||
      colFields.includes(index) ||
      pageFields.some((page) => page.field === index)
    ) {
      unsupported.push(
        `calculated field "${fields[index]?.name}" is used as a dimension, recompute not supported`,
      )
    }
  }

  // Value/label filters (pivotFilters): supports label equals/contains/begins-with
  // plus top N (top10)/greater-than/between; other types (date filters, bottom N,
  // percent…) still fail closed. Filtered-out members are also hidden entries, and
  // refresh drops them via the hidden semantics.
  const LABEL_FILTER_OPS: Record<string, 'equal' | 'contains' | 'beginsWith'> = {
    captionEqual: 'equal',
    captionContains: 'contains',
    captionBeginsWith: 'beginsWith',
  }
  const filters: PivotFilterDef[] = []
  const filtersBlock = firstElement(pivotTableXml, 'filters')
  for (const filterXml of filtersBlock === null ? [] : elements(filtersBlock, 'filter')) {
    const type = attr(filterXml, 'type') ?? ''
    const field = Number(attr(filterXml, 'fld') ?? '0')
    const dataField = Number(attr(filterXml, 'iMeasureFld') ?? '0')
    const customValues = elements(filterXml, 'customFilter').map((custom) =>
      Number(attr(custom, 'val') ?? 'NaN'),
    )
    const labelOp = LABEL_FILTER_OPS[type]
    if (labelOp !== undefined) {
      filters.push({
        kind: 'label',
        field,
        op: labelOp,
        value: decodeEntities(attr(filterXml, 'stringValue1') ?? ''),
      })
    } else if (type === 'count') {
      const top10 = firstElement(filterXml, 'top10')
      const count = Number(top10 === null ? 'NaN' : (attr(top10, 'val') ?? 'NaN'))
      // top="0" means bottom N, which is unsupported.
      if (!Number.isFinite(count) || count < 1 || (top10 !== null && attr(top10, 'top') === '0')) {
        unsupported.push(
          `value/label filter (pivotFilters) type "${type}" does not support recompute`,
        )
        continue
      }
      filters.push({ kind: 'value', field, dataField, op: 'top', count })
    } else if (type === 'valueGreaterThan' && Number.isFinite(customValues[0])) {
      filters.push({ kind: 'value', field, dataField, op: 'greaterThan', from: customValues[0]! })
    } else if (
      type === 'valueBetween' &&
      Number.isFinite(customValues[0]) &&
      Number.isFinite(customValues[1])
    ) {
      filters.push({
        kind: 'value',
        field,
        dataField,
        op: 'between',
        from: Math.min(customValues[0]!, customValues[1]!),
        to: Math.max(customValues[0]!, customValues[1]!),
      })
    } else {
      unsupported.push(
        `value/label filter (pivotFilters) type "${type}" does not support recompute`,
      )
    }
  }

  return {
    outputRef,
    firstDataRow,
    firstDataCol,
    fields,
    fieldItems,
    rowFields,
    colFields,
    rowLines: parseLayoutLines(firstElement(pivotTableXml, 'rowItems'), rowFields.length),
    colLines: parseLayoutLines(firstElement(pivotTableXml, 'colItems'), colFields.length),
    dataFields,
    pageFields,
    filters,
    sourceSheet,
    sourceRef,
    unsupported,
  }
}
