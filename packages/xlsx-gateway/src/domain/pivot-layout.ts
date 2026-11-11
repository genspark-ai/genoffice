import { allowedByValueFilter, matchesLabelFilter, type PivotFilterDef } from './pivot-filters'
import { evaluatePivotFormula, parsePivotFormula } from './pivot-formula'
import { groupValue, type PivotFieldGrouping } from './pivot-grouping'
import {
  defaultDataFieldCaption,
  normalizePivotPageFields,
  showDataAsIsPercent,
  showDataAsNeedsBaseField,
  showDataAsNeedsBaseItem,
  type PivotAggregation,
  type PivotShowDataAs,
} from './pivot-value-modes'
import type { AddPivotOperation } from './workbook-dsl'

export type PivotScalar = string | number | boolean | null

/// Mirror pivot-engine's numeric coercion so the grid baked here and a later
/// refresh total the same thing. A cell holding the text "42" is a number to
/// both; blank, empty and unparseable cells are not.
function numericValue(value: PivotScalar): number | null {
  if (value === null || value === '') return null
  const numeric = typeof value === 'number' ? value : Number(String(value).trim())
  return Number.isFinite(numeric) ? numeric : null
}

export type PivotLayoutErrorCode =
  | 'sourceNeedsRows'
  | 'sourceRowLimit'
  | 'sourceColLimit'
  | 'headerBlank'
  | 'headerDuplicate'
  | 'fieldNotHeader'
  | 'calcFieldNameClash'
  | 'calcFieldNameDuplicate'
  | 'valueFilterFieldMissing'
  | 'tooManyRowItems'
  | 'tooManyColItems'
  | 'tooManyColLines'
  | 'tooManyRowLines'
  | 'needsValues'
  | 'pageItemMissing'
  | 'baseFieldMissing'
  | 'baseFieldNotAxis'
  | 'baseItemMissing'

const MESSAGES: Record<PivotLayoutErrorCode, string> = {
  sourceNeedsRows: 'The pivot source needs a header row plus data rows.',
  sourceRowLimit: 'The pivot source is limited to 10,000 data rows.',
  sourceColLimit: 'The pivot source is limited to 200 columns.',
  headerBlank: 'Every source header cell must be non-blank.',
  headerDuplicate: 'Source header names must be unique.',
  fieldNotHeader: 'Field "{label}" is not a source header.',
  calcFieldNameClash: 'The calculated field "{name}" cannot share a name with a source header.',
  calcFieldNameDuplicate: 'Calculated field names must be unique.',
  valueFilterFieldMissing: 'The value filter references value field {index}, which does not exist.',
  tooManyRowItems: 'The pivot has more than 10,000 row items.',
  tooManyColItems: 'The pivot has more than 1,000 column items.',
  tooManyColLines: 'The pivot has more than 1,000 column lines.',
  tooManyRowLines: 'The pivot has more than 20,000 row lines.',
  needsValues: 'The pivot needs one values entry.',
  pageItemMissing: 'Report filter item "{item}" is not a member of field "{label}".',
  baseFieldMissing: 'The "show values as" mode of a value field needs a base field.',
  baseFieldNotAxis: 'Base field "{label}" must sit on the pivot rows or columns.',
  baseItemMissing: 'Base item "{item}" is not a member of field "{label}".',
}

export class PivotLayoutError extends Error {
  constructor(
    readonly code: PivotLayoutErrorCode,
    readonly params: Record<string, string | number> = {},
  ) {
    super(
      MESSAGES[code].replace(/\{(\w+)\}/g, (_, key: string) => String(params[key] ?? `{${key}}`)),
    )
    this.name = 'PivotLayoutError'
  }
}

export interface PivotLayoutLabels {
  readonly subtotal: string
  readonly grandTotal: string
}

export interface PivotLayoutValue {
  fieldIndex: number
  agg: PivotAggregation
  numFmt?: string | undefined
  showDataAs?: PivotShowDataAs | undefined
  /** "show values as" base field (a row/column dimension field index) */
  baseFieldIndex?: number | undefined
  /** base item's index within the base field's member list */
  baseItemIndex?: number | undefined
  /** custom caption; absent = "Sum of <field>" */
  name?: string | undefined
  formula?: string | undefined
  calcName?: string | undefined
}

export interface PivotLayoutLine {
  t: 'data' | 'default'
  members: number[]
}

/** The sheet-independent part of a pivot addition: what the OOXML definition needs. Mutable so it feeds the app's journal as-is. */
export interface PivotLayoutDefinition {
  fieldNames: string[]
  rowFieldIndices: number[]
  columnFieldIndex?: number | undefined
  pageFieldIndices?: number[] | undefined
  /** member lists per page field (sharedItems order), parallel to pageFieldIndices */
  pageLevelItems?: string[][] | undefined
  /** selected member per page field (index into pageLevelItems[p]); null = (All) */
  pageItems?: (number | null)[] | undefined
  rowItems: string[]
  rowLevelItems: string[][]
  rowLines: PivotLayoutLine[]
  columnItems?: string[] | undefined
  columnFieldIndices?: number[] | undefined
  colLevelItems?: string[][] | undefined
  colLines?: PivotLayoutLine[] | undefined
  groupings?: ({ fieldIndex: number } & PivotFieldGrouping)[] | undefined
  filters?: PivotFilterDef[] | undefined
  rowHiddenItems?: number[][] | undefined
  colHiddenItems?: number[][] | undefined
  values: PivotLayoutValue[]
}

export interface PivotLayout {
  readonly definition: PivotLayoutDefinition
  /** baked output grid, header rows first, grand total last */
  readonly matrix: readonly (readonly (string | number | null)[])[]
  readonly width: number
  readonly height: number
  /** number formats for value columns (offset from the anchor column); only without column dimensions */
  readonly numberFormats: readonly { readonly columnOffset: number; readonly format: string }[]
}

export type PivotLayoutSpec = Pick<
  AddPivotOperation,
  'rowFields' | 'columnField' | 'pageFields' | 'values' | 'groupings' | 'filters'
>

/// Resolved "show values as" base field: which axis level it lives on, that
/// level's members, the visible member indexes (display order, filter-hidden
/// members excluded), and the base item for the difference/percent modes.
export interface ResolvedShowAsBase {
  readonly axis: 'row' | 'column'
  readonly level: number
  readonly items: readonly string[]
  readonly visible: readonly number[]
  readonly baseItemIndex: number | undefined
}

export { AGG_CAPTIONS } from './pivot-value-modes'

export const PIVOT_SOURCE_ROW_LIMIT = 10_001
export const PIVOT_SOURCE_COL_LIMIT = 200

/**
 * Computes the pivot output the app bakes into cells and the definition the
 * gateway writes as OOXML, from the source grid (header row first). Pure: the
 * app and the CLI both call it, then place the result on a sheet.
 */
export function buildPivotLayout(
  grid: readonly (readonly PivotScalar[])[],
  op: PivotLayoutSpec,
  labels: PivotLayoutLabels = { subtotal: 'Subtotal', grandTotal: 'Grand Total' },
): PivotLayout {
  if (grid.length < 2) throw new PivotLayoutError('sourceNeedsRows')
  if (grid.length > PIVOT_SOURCE_ROW_LIMIT) throw new PivotLayoutError('sourceRowLimit')
  if ((grid[0]?.length ?? 0) > PIVOT_SOURCE_COL_LIMIT) throw new PivotLayoutError('sourceColLimit')

  const fieldNames = (grid[0] ?? []).map((value) => String(value ?? '').trim())
  if (fieldNames.some((name) => name.length === 0)) throw new PivotLayoutError('headerBlank')
  if (new Set(fieldNames.map((name) => name.toLowerCase())).size !== fieldNames.length) {
    throw new PivotLayoutError('headerDuplicate')
  }
  const fieldIndex = (label: string): number => {
    const index = fieldNames.indexOf(label)
    if (index < 0) throw new PivotLayoutError('fieldNotHeader', { label })
    return index
  }
  const rowFieldsArray = Array.isArray(op.rowFields) ? op.rowFields : [op.rowFields]
  const rowFieldIndices = rowFieldsArray.map(fieldIndex)
  const levels = rowFieldIndices.length
  const columnFieldsArray =
    op.columnField === undefined
      ? []
      : Array.isArray(op.columnField)
        ? op.columnField
        : [op.columnField]
  const columnFieldIndices = columnFieldsArray.map(fieldIndex)
  const colLevels = columnFieldIndices.length
  const pageSpecs = normalizePivotPageFields(op.pageFields)
  const pageFieldIndices = pageSpecs.map((page) => fieldIndex(page.field))
  const groupings = (op.groupings ?? []).map(({ field, ...rule }) => ({
    fieldIndex: fieldIndex(field),
    ...rule,
  }))
  const groupingByField = new Map<number, PivotFieldGrouping>()
  for (const { fieldIndex: groupedField, ...rule } of groupings) {
    groupingByField.set(groupedField, rule)
  }
  const dimGroupValue = (
    fieldIdx: number,
    row: readonly PivotScalar[],
  ): { label: string; sort: number | null } => {
    const grouping = groupingByField.get(fieldIdx)
    if (!grouping) return { label: String(row[fieldIdx] ?? ''), sort: null }
    return groupValue(grouping, row[fieldIdx])
  }
  const dimValue = (fieldIdx: number, row: readonly PivotScalar[]): string =>
    dimGroupValue(fieldIdx, row).label

  const valueSpecs = op.values.map((value) => {
    const isCalc = value.formula !== undefined
    if (
      isCalc &&
      fieldNames.some((name) => name.toLowerCase() === value.field.trim().toLowerCase())
    ) {
      throw new PivotLayoutError('calcFieldNameClash', { name: value.field })
    }
    return {
      fieldIndex: isCalc ? -1 : fieldIndex(value.field),
      agg: value.agg,
      // percent modes default to 0.00%, matching dataField numFmtId=10
      numFmt: value.numFmt ?? (showDataAsIsPercent(value.showDataAs) ? '0.00%' : undefined),
      showDataAs: value.showDataAs,
      baseFieldLabel: value.baseField,
      baseItemLabel: value.baseItem,
      base: undefined as ResolvedShowAsBase | undefined,
      name: value.name?.trim() || undefined,
      ...(isCalc ? { formula: value.formula, calcName: value.field } : {}),
      caption: value.name?.trim() || defaultDataFieldCaption(value.agg, value.field),
      ast: isCalc ? parsePivotFormula(value.formula!, fieldNames) : null,
    }
  })
  // Data-field captions must be unique; repeats get Excel's numeric suffix,
  // carried as the explicit name so the OOXML writer agrees with the grid.
  const seenCaptions = new Map<string, number>()
  for (const spec of valueSpecs) {
    const key = spec.caption.toLowerCase()
    const seen = seenCaptions.get(key) ?? 0
    seenCaptions.set(key, seen + 1)
    if (seen > 0) {
      spec.caption = `${spec.caption}${seen + 1}`
      spec.name = spec.caption
    }
  }
  const calcNameKeys = valueSpecs
    .filter((spec) => spec.calcName !== undefined)
    .map((spec) => spec.calcName!.trim().toLowerCase())
  if (new Set(calcNameKeys).size !== calcNameKeys.length) {
    throw new PivotLayoutError('calcFieldNameDuplicate')
  }

  // one globally deduplicated member list per level (first-seen = sharedItems order)
  const levelItems: string[][] = rowFieldIndices.map(() => [])
  const comboKeys = new Set<string>()
  const colLevelItems: string[][] = columnFieldIndices.map(() => [])
  const colComboKeys = new Set<string>()
  const levelSortKeys: Map<string, number | null>[] = rowFieldIndices.map(() => new Map())
  const colLevelSortKeys: Map<string, number | null>[] = columnFieldIndices.map(() => new Map())
  const joinPath = (path: readonly string[]): string => path.join('\u0000')
  const allDataRows = grid.slice(1)
  // Page (report filter) members come from every source row; a selected item
  // then narrows the rows the rest of the report sees, as in Excel.
  const pageLevelItems: string[][] = pageFieldIndices.map((fieldIdx) => {
    const items: string[] = []
    for (const row of allDataRows) {
      const label = String(row[fieldIdx] ?? '')
      if (!items.includes(label)) items.push(label)
    }
    return items
  })
  const pageItems: (number | null)[] = pageSpecs.map((page, index) => {
    if (page.item === undefined) return null
    const at = pageLevelItems[index]!.indexOf(page.item)
    if (at < 0)
      throw new PivotLayoutError('pageItemMissing', { item: page.item, label: page.field })
    return at
  })
  const dataRows = allDataRows.filter((row) =>
    pageItems.every(
      (item, index) =>
        item === null ||
        String(row[pageFieldIndices[index]!] ?? '') === pageLevelItems[index]![item],
    ),
  )
  for (const row of allDataRows) {
    rowFieldIndices.forEach((fieldIdx, level) => {
      const { label: key, sort } = dimGroupValue(fieldIdx, row)
      if (!levelItems[level]!.includes(key)) {
        levelItems[level]!.push(key)
        levelSortKeys[level]!.set(key, sort)
      }
    })
    columnFieldIndices.forEach((fieldIdx, level) => {
      const { label: key, sort } = dimGroupValue(fieldIdx, row)
      if (!colLevelItems[level]!.includes(key)) {
        colLevelItems[level]!.push(key)
        colLevelSortKeys[level]!.set(key, sort)
      }
    })
  }
  // grouped levels sort by bucket; pass-through values keep first-seen order after them
  const sortGroupedLevel = (
    items: string[],
    fieldIdx: number,
    sortKeys: Map<string, number | null>,
  ): void => {
    if (!groupingByField.has(fieldIdx)) return
    items.sort((a, b) => {
      const sortA = sortKeys.get(a) ?? null
      const sortB = sortKeys.get(b) ?? null
      if (sortA !== null && sortB !== null) return sortA - sortB
      if (sortA !== null) return -1
      if (sortB !== null) return 1
      return 0
    })
  }
  rowFieldIndices.forEach((fieldIdx, level) =>
    sortGroupedLevel(levelItems[level]!, fieldIdx, levelSortKeys[level]!),
  )
  columnFieldIndices.forEach((fieldIdx, level) =>
    sortGroupedLevel(colLevelItems[level]!, fieldIdx, colLevelSortKeys[level]!),
  )
  const rowItems = levelItems[0] ?? []
  if (levelItems.some((items) => items.length > 10_000)) {
    throw new PivotLayoutError('tooManyRowItems')
  }
  if (colLevelItems.some((items) => items.length > 1_000)) {
    throw new PivotLayoutError('tooManyColItems')
  }

  const aggregate = (
    rows: readonly (readonly PivotScalar[])[],
    spec: (typeof valueSpecs)[number],
  ): number | null => {
    if (spec.ast) {
      return evaluatePivotFormula(spec.ast, (name) => {
        const refIndex = fieldIndex(name)
        return rows
          .map((row) => numericValue(row[refIndex] ?? null))
          .filter((value): value is number => value !== null)
          .reduce((total, value) => total + value, 0)
      })
    }
    if (spec.agg === 'count') {
      return rows.filter((row) => row[spec.fieldIndex] != null && row[spec.fieldIndex] !== '')
        .length
    }
    const numbers = rows
      .map((row) => numericValue(row[spec.fieldIndex] ?? null))
      .filter((value): value is number => value !== null)
    if (spec.agg === 'countNums') return numbers.length
    if (numbers.length === 0) return null
    switch (spec.agg) {
      case 'sum':
        return numbers.reduce((total, value) => total + value, 0)
      case 'product':
        return numbers.reduce((total, value) => total * value, 1)
      case 'average':
        return numbers.reduce((total, value) => total + value, 0) / numbers.length
      case 'max':
        return numbers.reduce((max, value) => (value > max ? value : max))
      case 'min':
        return numbers.reduce((min, value) => (value < min ? value : min))
      case 'var':
      case 'stdDev': {
        // sample statistics (n-1), matching the refresh engine; a one-value
        // bucket is blank here where Excel shows #DIV/0!
        if (numbers.length < 2) return null
        const mean = numbers.reduce((t, v) => t + v, 0) / numbers.length
        const variance =
          numbers.reduce((t, v) => t + (v - mean) * (v - mean), 0) / (numbers.length - 1)
        return spec.agg === 'var' ? variance : Math.sqrt(variance)
      }
      case 'varp':
      case 'stdDevp': {
        if (numbers.length === 0) return null
        const mean = numbers.reduce((t, v) => t + v, 0) / numbers.length
        const variance = numbers.reduce((t, v) => t + (v - mean) * (v - mean), 0) / numbers.length
        return spec.agg === 'varp' ? variance : Math.sqrt(variance)
      }
    }
  }

  const filterEntries: PivotFilterDef[] = (op.filters ?? []).map((filter) =>
    filter.kind === 'label'
      ? { kind: 'label', field: fieldIndex(filter.field), op: filter.op, value: filter.value }
      : {
          kind: 'value',
          field: fieldIndex(filter.field),
          dataField: filter.valueIndex,
          op: filter.op,
          ...(filter.count !== undefined ? { count: filter.count } : {}),
          ...(filter.from !== undefined ? { from: filter.from } : {}),
          ...(filter.to !== undefined ? { to: filter.to } : {}),
        },
  )
  // label filters pick members; value filters aggregate each candidate over the label-filtered rows
  const membersOfField = (fieldIdx: number): string[] => {
    const rowLevel = rowFieldIndices.indexOf(fieldIdx)
    if (rowLevel >= 0) return levelItems[rowLevel]!
    const colLevel = columnFieldIndices.indexOf(fieldIdx)
    return colLevel >= 0 ? colLevelItems[colLevel]! : []
  }
  const hiddenByField = new Map<number, Set<string>>()
  for (const filter of filterEntries) {
    if (filter.kind !== 'label') continue
    const hidden = hiddenByField.get(filter.field) ?? new Set<string>()
    for (const member of membersOfField(filter.field)) {
      if (!matchesLabelFilter(filter, member)) hidden.add(member)
    }
    hiddenByField.set(filter.field, hidden)
  }
  const rowVisible = (row: readonly PivotScalar[]): boolean => {
    for (const [fieldIdx, hidden] of hiddenByField) {
      if (hidden.has(dimValue(fieldIdx, row))) return false
    }
    return true
  }
  for (const filter of filterEntries) {
    if (filter.kind !== 'value') continue
    const spec = valueSpecs[filter.dataField]
    if (!spec)
      throw new PivotLayoutError('valueFilterFieldMissing', { index: filter.dataField + 1 })
    const labelRows = dataRows.filter(rowVisible)
    const hidden = hiddenByField.get(filter.field) ?? new Set<string>()
    const candidates = membersOfField(filter.field).filter((member) => !hidden.has(member))
    const totals = candidates.map((member) => ({
      key: member,
      total: aggregate(
        labelRows.filter((row) => dimValue(filter.field, row) === member),
        spec,
      ),
    }))
    const kept = allowedByValueFilter(filter, totals)
    for (const member of candidates) {
      if (!kept.has(member)) hidden.add(member)
    }
    hiddenByField.set(filter.field, hidden)
  }
  const visibleRows = dataRows.filter(rowVisible)
  const hiddenIndexesOf = (
    fieldIndices: readonly number[],
    items: readonly (readonly string[])[],
  ): number[][] =>
    fieldIndices.map((fieldIdx, level) => {
      const hidden = hiddenByField.get(fieldIdx)
      if (!hidden) return []
      return (items[level] ?? []).flatMap((member, index) => (hidden.has(member) ? [index] : []))
    })
  const rowHiddenItems = hiddenIndexesOf(rowFieldIndices, levelItems)
  const colHiddenItems = hiddenIndexesOf(columnFieldIndices, colLevelItems)

  // "Show values as" base fields: resolved against the axis levels once member
  // lists and filter-hidden items are known. A base field outside the axes, or
  // a base item outside the base field's members, fails closed.
  for (const spec of valueSpecs) {
    const mode = spec.showDataAs
    if (mode === undefined || !showDataAsNeedsBaseField(mode)) continue
    const label = spec.baseFieldLabel
    if (label === undefined) throw new PivotLayoutError('baseFieldMissing')
    const fieldIdx = fieldIndex(label)
    const rowLevel = rowFieldIndices.indexOf(fieldIdx)
    const columnLevel = columnFieldIndices.indexOf(fieldIdx)
    if (rowLevel < 0 && columnLevel < 0) {
      throw new PivotLayoutError('baseFieldNotAxis', { label })
    }
    const axis: 'row' | 'column' = rowLevel >= 0 ? 'row' : 'column'
    const level = rowLevel >= 0 ? rowLevel : columnLevel
    const items = axis === 'row' ? levelItems[level]! : colLevelItems[level]!
    const hidden = new Set(axis === 'row' ? rowHiddenItems[level]! : colHiddenItems[level]!)
    let baseItemIndex: number | undefined
    if (showDataAsNeedsBaseItem(mode)) {
      const itemLabel = spec.baseItemLabel ?? ''
      baseItemIndex = items.indexOf(itemLabel)
      if (baseItemIndex < 0) {
        throw new PivotLayoutError('baseItemMissing', { item: itemLabel, label })
      }
    }
    spec.base = {
      axis,
      level,
      items,
      visible: items.flatMap((_, index) => (hidden.has(index) ? [] : [index])),
      baseItemIndex,
    }
  }

  for (const row of visibleRows) {
    const path: string[] = []
    rowFieldIndices.forEach((fieldIdx) => {
      path.push(dimValue(fieldIdx, row))
      comboKeys.add(joinPath(path))
    })
    const colPath: string[] = []
    columnFieldIndices.forEach((fieldIdx) => {
      colPath.push(dimValue(fieldIdx, row))
      colComboKeys.add(joinPath(colPath))
    })
  }

  // column lines (without the grand-total column): data columns on every level, a subtotal column after each non-leaf member
  const colLines: PivotLayoutLine[] = []
  const colLinePaths: string[][] = []
  const emitColLevel = (path: string[], indices: number[], depth: number): void => {
    colLevelItems[depth]!.forEach((member, memberIndex) => {
      if (!colComboKeys.has(joinPath([...path, member]))) return
      if (depth === colLevels - 1) {
        colLines.push({ t: 'data', members: [...indices, memberIndex] })
        colLinePaths.push([...path, member])
      } else {
        emitColLevel([...path, member], [...indices, memberIndex], depth + 1)
        colLines.push({ t: 'default', members: [...indices, memberIndex] })
        colLinePaths.push([...path, member])
      }
    })
  }
  if (colLevels > 0) emitColLevel([], [], 0)
  if (colLines.length > 1_000) throw new PivotLayoutError('tooManyColLines')

  const bucketRows = (prefix: readonly string[], colPrefix: readonly string[]): PivotScalar[][] =>
    visibleRows.filter((row) => {
      for (let level = 0; level < prefix.length; level += 1) {
        if (dimValue(rowFieldIndices[level]!, row) !== prefix[level]) return false
      }
      for (let level = 0; level < colPrefix.length; level += 1) {
        if (dimValue(columnFieldIndices[level]!, row) !== colPrefix[level]) return false
      }
      return true
    }) as PivotScalar[][]

  // Running-total / rank accumulations per (value spec × group): the group is
  // every coordinate except the base-field level, so the cumulative restarts
  // for each new outer member of the base axis and each other-axis group —
  // Excel's reset rule.
  const showAsAccumulations = new Map<string, { cum: number[]; values: (number | null)[] }>()
  const axisAggregate = (
    spec: (typeof valueSpecs)[number],
    ownPrefix: readonly string[],
    otherPrefix: readonly string[],
  ): number | null =>
    spec.base?.axis === 'row'
      ? aggregate(bucketRows(ownPrefix, otherPrefix), spec)
      : aggregate(bucketRows(otherPrefix, ownPrefix), spec)

  const applyShowDataAs = (
    raw: number | null,
    spec: (typeof valueSpecs)[number],
    prefix: readonly string[],
    colPrefix: readonly string[],
  ): number | null => {
    const mode = spec.showDataAs
    if (mode === undefined || raw === null) return raw
    const ratio = (base: number | null): number | null =>
      base === null || base === 0 ? null : raw / base
    const base = spec.base
    // Per-axis prefixes of the base-field modes: `own` is the axis holding the
    // base field; a cell whose own prefix does not fix the base level (a
    // subtotal or grand total spanning it) shows blank, as in Excel.
    const ownPrefix = base?.axis === 'column' ? colPrefix : prefix
    const otherPrefix = base?.axis === 'column' ? prefix : colPrefix
    const ownFixed = base !== undefined && ownPrefix.length > base.level
    const replaceItem = (item: string): readonly string[] => {
      const at = base!.level
      return [...ownPrefix.slice(0, at), item, ...ownPrefix.slice(at + 1)]
    }
    const parentScope = (): readonly string[] => ownPrefix.slice(0, base!.level)
    switch (mode) {
      case 'percentOfTotal':
        return ratio(aggregate(bucketRows([], []), spec))
      case 'percentOfRow':
        return ratio(aggregate(bucketRows(prefix, []), spec))
      case 'percentOfCol':
        return ratio(aggregate(bucketRows([], colPrefix), spec))
      case 'percentOfParentRow':
        return ratio(aggregate(bucketRows(prefix.slice(0, -1), colPrefix), spec))
      case 'percentOfParentCol':
        return ratio(aggregate(bucketRows(prefix, colPrefix.slice(0, -1)), spec))
      case 'index': {
        const grand = aggregate(bucketRows([], []), spec)
        const rowTotal = aggregate(bucketRows(prefix, []), spec)
        const colTotal = aggregate(bucketRows([], colPrefix), spec)
        if (grand === null || rowTotal === null || colTotal === null) return null
        const denominator = rowTotal * colTotal
        return denominator === 0 ? null : (raw * grand) / denominator
      }
      case 'difference':
      case 'percentDiff':
      case 'percentOf': {
        if (!base || !ownFixed || base.baseItemIndex === undefined) return null
        const baseItem = base.items[base.baseItemIndex]!
        // Difference and % difference are blank for the base item itself;
        // "% Of" shows the natural raw/raw = 100%.
        if (mode !== 'percentOf' && ownPrefix[base.level] === baseItem) return null
        const baseValue = axisAggregate(spec, replaceItem(baseItem), otherPrefix)
        if (baseValue === null) return null
        if (mode === 'difference') return raw - baseValue
        if (baseValue === 0) return null
        return mode === 'percentOf' ? raw / baseValue : (raw - baseValue) / baseValue
      }
      case 'percentOfParent': {
        if (!base) return null
        return ratio(axisAggregate(spec, parentScope(), otherPrefix))
      }
      case 'runTotal':
      case 'percentOfRunningTotal':
      case 'rankAscending':
      case 'rankDescending': {
        if (!base || !ownFixed) return null
        const ownIndex = base.items.indexOf(ownPrefix[base.level]!)
        if (ownIndex < 0) return null
        const groupKey =
          `${spec.caption}\u0000${otherPrefix.join('\u0000')}\u0000` +
          [...ownPrefix.slice(0, base.level), ...ownPrefix.slice(base.level + 1)].join('\u0000')
        let acc = showAsAccumulations.get(groupKey)
        if (acc === undefined) {
          // Values per visible base item (item display order), then cumulative.
          const values = base.visible.map((index) =>
            axisAggregate(spec, replaceItem(base.items[index]!), otherPrefix),
          )
          let sum = 0
          const cum = values.map((value) => (sum += value ?? 0))
          acc = { cum, values }
          showAsAccumulations.set(groupKey, acc)
        }
        const at = base.visible.indexOf(ownIndex)
        if (at < 0) return null
        if (mode === 'runTotal') return acc.cum[at] ?? null
        if (mode === 'percentOfRunningTotal') {
          const denominator = axisAggregate(spec, parentScope(), otherPrefix)
          return denominator === null || denominator === 0 ? null : acc.cum[at]! / denominator
        }
        // RANK.EQ-style ties: equal values share the smaller rank.
        const better = acc.values.filter(
          (value, index) =>
            index !== at &&
            value !== null &&
            (mode === 'rankAscending' ? value < raw : value > raw),
        ).length
        return better + 1
      }
    }
  }

  const valueCells = (prefix: readonly string[]): (number | null)[] => {
    if (colLevels === 0) {
      return valueSpecs.map((spec) =>
        applyShowDataAs(aggregate(bucketRows(prefix, []), spec), spec, prefix, []),
      )
    }
    const spec = valueSpecs[0]
    if (!spec) throw new PivotLayoutError('needsValues')
    return [
      ...colLinePaths.map((colPrefix) =>
        applyShowDataAs(aggregate(bucketRows(prefix, colPrefix), spec), spec, prefix, colPrefix),
      ),
      applyShowDataAs(aggregate(bucketRows(prefix, []), spec), spec, prefix, []),
    ]
  }

  const matrix: (string | number | null)[][] = []
  if (colLevels === 0) {
    matrix.push([...rowFieldsArray, ...valueSpecs.map((spec) => spec.caption)])
  } else {
    const totalWidth = levels + colLines.length + 1
    const headers: (string | number | null)[][] = Array.from({ length: colLevels }, () =>
      new Array<string | number | null>(totalWidth).fill(null),
    )
    rowFieldsArray.forEach((name, level) => {
      headers[colLevels - 1]![level] = name
    })
    let previousColPath: readonly string[] | null = null
    colLines.forEach((line, index) => {
      const colOffset = levels + index
      const path = colLinePaths[index]!
      if (line.t === 'default') {
        path.forEach((member, level) => {
          headers[level]![colOffset] = member
        })
        if (path.length < colLevels) headers[path.length]![colOffset] = labels.subtotal
        previousColPath = null
      } else {
        path.forEach((member, level) => {
          if (previousColPath !== null && level < path.length - 1) {
            let samePrefix = true
            for (let k = 0; k <= level; k += 1) {
              if (previousColPath[k] !== path[k]) {
                samePrefix = false
                break
              }
            }
            if (samePrefix) return
          }
          headers[level]![colOffset] = member
        })
        previousColPath = path
      }
    })
    headers[0]![levels + colLines.length] = labels.grandTotal
    matrix.push(...headers)
  }
  const rowLines: PivotLayoutLine[] = []
  const pendingLabels: (string | null)[] = new Array(Math.max(0, levels - 1)).fill(null)
  const emitLevel = (path: string[], indices: number[], depth: number): void => {
    levelItems[depth]!.forEach((member, memberIndex) => {
      if (!comboKeys.has(joinPath([...path, member]))) return
      if (depth === levels - 1) {
        const rowLabels: (string | null)[] = [...pendingLabels, member]
        pendingLabels.fill(null)
        matrix.push([...rowLabels, ...valueCells([...path, member])])
        rowLines.push({ t: 'data', members: [...indices, memberIndex] })
      } else {
        pendingLabels[depth] = member
        emitLevel([...path, member], [...indices, memberIndex], depth + 1)
        const subtotalLabels: (string | null)[] = new Array(levels).fill(null)
        for (let level = 0; level < depth; level += 1) subtotalLabels[level] = path[level]!
        subtotalLabels[depth] = member
        subtotalLabels[depth + 1] = labels.subtotal
        matrix.push([...subtotalLabels, ...valueCells([...path, member])])
        rowLines.push({ t: 'default', members: [...indices, memberIndex] })
      }
    })
  }
  emitLevel([], [], 0)
  if (rowLines.length > 20_000) throw new PivotLayoutError('tooManyRowLines')
  matrix.push([
    ...rowFieldsArray.map((_, i) => (i === 0 ? labels.grandTotal : '')),
    ...valueCells([]),
  ])

  const height = matrix.length
  const width = matrix[0]?.length ?? 0
  const numberFormats =
    colLevels === 0
      ? valueSpecs.flatMap((spec, vi) =>
          spec.numFmt ? [{ columnOffset: levels + vi, format: spec.numFmt }] : [],
        )
      : []

  return {
    matrix,
    width,
    height,
    numberFormats,
    definition: {
      fieldNames,
      rowFieldIndices,
      ...(colLevels === 1 ? { columnFieldIndex: columnFieldIndices[0]! } : {}),
      ...(pageFieldIndices.length > 0 ? { pageFieldIndices, pageLevelItems, pageItems } : {}),
      rowItems,
      rowLevelItems: levelItems,
      rowLines,
      ...(colLevels === 1 ? { columnItems: colLevelItems[0]! } : {}),
      ...(colLevels >= 2 ? { columnFieldIndices, colLevelItems, colLines } : {}),
      ...(groupings.length > 0 ? { groupings } : {}),
      ...(filterEntries.length > 0
        ? { filters: filterEntries, rowHiddenItems, colHiddenItems }
        : {}),
      values: valueSpecs.map((spec) => {
        const baseFieldIndex =
          spec.base === undefined
            ? undefined
            : spec.base.axis === 'row'
              ? rowFieldIndices[spec.base.level]!
              : columnFieldIndices[spec.base.level]!
        return {
          fieldIndex: spec.fieldIndex,
          agg: spec.agg,
          ...(spec.numFmt ? { numFmt: spec.numFmt } : {}),
          ...(spec.showDataAs ? { showDataAs: spec.showDataAs } : {}),
          ...(baseFieldIndex !== undefined ? { baseFieldIndex } : {}),
          ...(spec.base?.baseItemIndex !== undefined
            ? { baseItemIndex: spec.base.baseItemIndex }
            : {}),
          ...(spec.name !== undefined ? { name: spec.name } : {}),
          ...(spec.formula !== undefined
            ? { formula: spec.formula, calcName: spec.calcName! }
            : {}),
        }
      }),
    },
  }
}

export interface PivotArea {
  readonly startRow: number
  readonly startColumn: number
  readonly endRow: number
  readonly endColumn: number
}

export function pivotOutputArea(
  anchor: { row: number; column: number },
  layout: PivotLayout,
): PivotArea {
  return {
    startRow: anchor.row,
    startColumn: anchor.column,
    endRow: anchor.row + layout.height - 1,
    endColumn: anchor.column + layout.width - 1,
  }
}

export function areasOverlap(a: PivotArea, b: PivotArea): boolean {
  return (
    a.startRow <= b.endRow &&
    b.startRow <= a.endRow &&
    a.startColumn <= b.endColumn &&
    b.startColumn <= a.endColumn
  )
}
