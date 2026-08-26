/**
 * Workbook operation appliers for the sheets renderer.
 *
 * AI DSL operations (pivot / table / table-column adds) and pivot dialog
 * helpers applied against the live Univer runtime. Extracted from App.tsx;
 * every function receives its runtime and state explicitly.
 */
import { isDefaultFormat, numfmt } from '@univerjs/core'
import { DATE_1904_OFFSET, isCalendarDatePattern } from './numfmt-fix'
import type {
  AddPivotOperation,
  AddTableColumnOperation,
  AddTableOperation,
  AddTableRowOperation,
  DeleteTableColumnOperation,
  DeleteTableRowOperation,
} from '../domain/workbook-dsl'
import { columnLabel, parseAddress, parseRange } from '../domain/cell-address'
import { renameRefSheet } from '../domain/chart-visual'
import {
  allowedByValueFilter,
  matchesLabelFilter,
  type PivotFilterDef,
} from '../domain/pivot-filters'
import { evaluatePivotFormula, parsePivotFormula } from '../domain/pivot-formula'
import { groupValue, type PivotFieldGrouping } from '../domain/pivot-grouping'
import type { WorkbookVisualObject } from '../shared/desktop-api'
import {
  recordPivotAdd,
  recordPivotCacheRefresh,
  recordPivotRefreshUpdate,
  recordTableAdd,
  updateTableAdd,
} from './edit-journal'
import { t } from './i18n/locale'
import type { OoXmlPivotConfig, PivotField } from './PivotDialog'
import type { PivotDefinition } from '../gateway/xlsx-pivot'
import {
  AGG_CAPTIONS,
  applyFormatPatchToRange,
  nextSessionPivotName,
  nextSessionTableName,
} from './univer-sync'
import type { LazyWorkbookState, UniverRuntime, UniverWorksheet } from './univer-state'

export function applyAiTableAdd(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: AddTableOperation,
): void {
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(op.sheetId)
  if (!worksheet) throw new Error(`Unknown sheet: ${op.sheetId}`)
  const bounds = parseRange(op.range)
  if (bounds.endRow <= bounds.startRow) {
    throw new Error(t('appTableNeedsRows'))
  }
  const width = bounds.endColumn - bounds.startColumn + 1
  if (width > 1_000) throw new Error(t('appTableTooWide'))
  const name = op.name ?? nextSessionTableName(state.editJournal)
  if (
    state.editJournal.tableAdds.some((table) => table.name.toLowerCase() === name.toLowerCase())
  ) {
    throw new Error(t('appTableNameUsed', { name }))
  }
  for (const table of state.editJournal.tableAdds) {
    if (table.sheetId !== op.sheetId) continue
    const apart =
      bounds.endRow < table.area.startRow ||
      table.area.endRow < bounds.startRow ||
      bounds.endColumn < table.area.startColumn ||
      table.area.endColumn < bounds.startColumn
    if (!apart) {
      throw new Error(t('appTableOverlapsSession', { name: table.name }))
    }
  }
  const headerValues =
    worksheet.getRange(bounds.startRow, bounds.startColumn, 1, width).getValues()[0] ?? []
  const columnNames: string[] = []
  const used = new Set<string>()
  for (let index = 0; index < width; index += 1) {
    const raw = String(headerValues[index] ?? '')
      .trim()
      .slice(0, 255)
    const base = raw.length === 0 ? `Column${index + 1}` : raw
    let candidate = base
    for (let suffix = 2; used.has(candidate.toLowerCase()); suffix += 1) {
      candidate = `${base}${suffix}`
    }
    used.add(candidate.toLowerCase())
    columnNames.push(candidate)
    if (candidate !== raw) {
      worksheet.getRange(bounds.startRow, bounds.startColumn + index).setValue(candidate)
    }
  }
  // Rendering is best-effort (the facade call is async); the journal entry
  // below is what the save writes, and the gateway re-checks conflicts.
  void worksheet.addTable(
    name,
    {
      startRow: bounds.startRow,
      startColumn: bounds.startColumn,
      endRow: bounds.endRow,
      endColumn: bounds.endColumn,
    },
    `ai-table-${state.editJournal.tableAdds.length + 1}-${Date.now().toString(36)}`,
  )
  recordTableAdd(state.editJournal, {
    sheetId: op.sheetId,
    area: {
      startRow: bounds.startRow,
      startColumn: bounds.startColumn,
      endRow: bounds.endRow,
      endColumn: bounds.endColumn,
    },
    name,
    columnNames,
    ...(op.style === undefined ? {} : { style: op.style }),
    bandedRows: op.bandedRows ?? true,
  })
}

/// AI add_table_row: inserts one or more rows into a session-added table.
/// The row index is 1-based within the data region (header = row 0).
export function applyAiTableRowAdd(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: AddTableRowOperation,
): void {
  const tableEntry = state.editJournal.tableAdds.find(
    (t) => t.sheetId === op.sheetId && t.name.toLowerCase() === op.tableName.toLowerCase(),
  )
  if (!tableEntry) throw new Error(`Table "${op.tableName}" not found in session journal.`)
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(op.sheetId)
  if (!worksheet) throw new Error(`Unknown sheet: ${op.sheetId}`)
  const count = op.count ?? 1
  const dataRows = tableEntry.area.endRow - tableEntry.area.startRow
  // row is 1-based within data region; default = append at end
  const insertDataRow = op.row ?? dataRows + 1
  // Absolute sheet row: header is tableEntry.area.startRow (0-based),
  // data starts at startRow + 1.  Insert before this row in the sheet.
  const sheetRow = tableEntry.area.startRow + insertDataRow
  worksheet.insertRowsBefore(sheetRow, count)
  // Expand the journal entry's area
  updateTableAdd(state.editJournal, op.sheetId, op.tableName, {
    area: { ...tableEntry.area, endRow: tableEntry.area.endRow + count },
  })
}

/// AI delete_table_row: removes one or more rows from a session-added table.
export function applyAiTableRowDelete(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: DeleteTableRowOperation,
): void {
  const tableEntry = state.editJournal.tableAdds.find(
    (t) => t.sheetId === op.sheetId && t.name.toLowerCase() === op.tableName.toLowerCase(),
  )
  if (!tableEntry) throw new Error(`Table "${op.tableName}" not found in session journal.`)
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(op.sheetId)
  if (!worksheet) throw new Error(`Unknown sheet: ${op.sheetId}`)
  const count = op.count ?? 1
  const sheetRow = tableEntry.area.startRow + op.row // 1-based data row → absolute
  worksheet.deleteRows(sheetRow, count)
  updateTableAdd(state.editJournal, op.sheetId, op.tableName, {
    area: { ...tableEntry.area, endRow: tableEntry.area.endRow - count },
  })
}

/// AI delete_table_column: removes one or more columns from a session-added table.
export function applyAiTableColumnDelete(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: DeleteTableColumnOperation,
): void {
  const tableEntry = state.editJournal.tableAdds.find(
    (t) => t.sheetId === op.sheetId && t.name.toLowerCase() === op.tableName.toLowerCase(),
  )
  if (!tableEntry) throw new Error(`Table "${op.tableName}" not found in session journal.`)
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(op.sheetId)
  if (!worksheet) throw new Error(`Unknown sheet: ${op.sheetId}`)
  const count = op.count ?? 1
  const sheetCol = tableEntry.area.startColumn + op.column - 1
  worksheet.deleteColumns(sheetCol, count)
  const newCols = [...tableEntry.columnNames]
  newCols.splice(op.column - 1, count)
  updateTableAdd(state.editJournal, op.sheetId, op.tableName, {
    area: { ...tableEntry.area, endColumn: tableEntry.area.endColumn - count },
    columnNames: newCols,
  })
}

export function applyAiTableColumnAdd(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: AddTableColumnOperation,
): void {
  const tableEntry = state.editJournal.tableAdds.find(
    (t) => t.sheetId === op.sheetId && t.name.toLowerCase() === op.tableName.toLowerCase(),
  )
  if (!tableEntry) throw new Error(`Table "${op.tableName}" not found in session journal.`)
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(op.sheetId)
  if (!worksheet) throw new Error(`Unknown sheet: ${op.sheetId}`)
  const count = op.count ?? 1
  const tableCols = tableEntry.area.endColumn - tableEntry.area.startColumn + 1
  const insertTableCol = op.column ?? tableCols + 1
  const sheetCol = tableEntry.area.startColumn + insertTableCol - 1
  worksheet.insertColumnsBefore(sheetCol, count)
  // Update column names list (insert placeholder(s) before insertTableCol-1)
  const newCols = [...tableEntry.columnNames]
  for (let i = 0; i < count; i += 1) {
    const colName = i === 0 ? op.columnName : `${op.columnName}${i + 1}`
    newCols.splice(insertTableCol - 1 + i, 0, colName)
  }
  // Write the header cell(s) so the worksheet reflects the new name(s)
  for (let i = 0; i < count; i += 1) {
    worksheet
      .getRange(tableEntry.area.startRow, sheetCol + i)
      .setValue(newCols[insertTableCol - 1 + i] ?? '')
  }
  updateTableAdd(state.editJournal, op.sheetId, op.tableName, {
    area: { ...tableEntry.area, endColumn: tableEntry.area.endColumn + count },
    columnNames: newCols,
  })
}

/// Width-independent pivot source read. getValues() goes through the view
/// interceptors, which clip numbers to the column width (#### fill,
/// budget-limited General) — lies that must not become pivot labels,
/// aggregates, saved sharedItems, or recompute keys (a date column displayed
/// as #### broke timeline binding and slicer recompute). Re-format raw
/// values through each cell's own pattern instead; General numbers stay
/// numeric.
export function readPivotSourceGrid(
  range: {
    getRawValues(): unknown[][]
    getNumberFormats(): string[][]
    getFormulas(): string[][]
  },
  date1904: boolean,
): (string | number | boolean | null)[][] {
  const patterns = range.getNumberFormats()
  // 1904 workbooks: file-loaded static serials count from 1904-01-01, but
  // formula results come from the 1900-based engine — same rule as the
  // display interceptor's calendar-date shift.
  const formulas = date1904 ? range.getFormulas() : null
  return (range.getRawValues() as (string | number | boolean | null)[][]).map((row, rowOffset) =>
    row.map((value, columnOffset) => {
      if (typeof value !== 'number') return value
      const pattern = patterns[rowOffset]?.[columnOffset] ?? ''
      if (pattern === '' || isDefaultFormat(pattern)) return value
      const shift =
        formulas !== null &&
        isCalendarDatePattern(pattern) &&
        (formulas[rowOffset]?.[columnOffset] ?? '') === ''
          ? DATE_1904_OFFSET
          : 0
      try {
        return numfmt.format(pattern, value + shift, { nbsp: true, throws: false })
      } catch {
        return value
      }
    }),
  )
}

export function applyAiPivotAdd(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: AddPivotOperation,
  relayout?: {
    readonly pivotPath: string
    readonly cachePath: string
    readonly oldBounds: {
      startRow: number
      startColumn: number
      endRow: number
      endColumn: number
    }
  },
): void {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const sourceSheet = workbook?.getSheetBySheetId(op.sheetId)
  const targetSheetId = op.targetSheetId ?? op.sheetId
  const targetSheet = workbook?.getSheetBySheetId(targetSheetId)
  if (!sourceSheet) throw new Error(`Unknown sheet: ${op.sheetId}`)
  if (!targetSheet) throw new Error(`Unknown sheet: ${targetSheetId}`)
  if (
    state.editJournal.sheets.added.has(op.sheetId) ||
    state.editJournal.sheets.added.has(targetSheetId)
  ) {
    throw new Error(t('appPivotOnAddedSheet'))
  }
  const source = parseRange(op.sourceRange)
  const sourceRows = source.endRow - source.startRow + 1
  const sourceColumns = source.endColumn - source.startColumn + 1
  if (sourceRows < 2) throw new Error(t('appPivotSourceNeedsRows'))
  if (sourceRows > 10_001) throw new Error(t('appPivotSourceRowLimit'))
  if (sourceColumns > 200) throw new Error(t('appPivotSourceColLimit'))

  // Univer's Nullable<CellValue> narrows to the scalar union we aggregate
  // over; undefined behaves like an empty cell throughout.
  const grid = readPivotSourceGrid(
    sourceSheet.getRange(source.startRow, source.startColumn, sourceRows, sourceColumns),
    state.file.date1904 === true,
  )
  const fieldNames = (grid[0] ?? []).map((value) => String(value ?? '').trim())
  if (fieldNames.some((name) => name.length === 0)) {
    throw new Error(t('appPivotHeaderBlank'))
  }
  if (new Set(fieldNames.map((name) => name.toLowerCase())).size !== fieldNames.length) {
    throw new Error(t('appPivotHeaderDuplicate'))
  }
  const fieldIndex = (label: string): number => {
    const index = fieldNames.indexOf(label)
    if (index < 0) throw new Error(t('appPivotFieldNotHeader', { label }))
    return index
  }
  // rowFields is normalized to an array by the DSL expander
  const rowFieldsArray = Array.isArray(op.rowFields) ? op.rowFields : [op.rowFields]
  const rowFieldIndices = rowFieldsArray.map(fieldIndex)
  const levels = rowFieldIndices.length
  // Column dimensions mirror rows: 0-8 levels, outer first (single level
  // accepts a bare string).
  const columnFieldsArray =
    op.columnField === undefined
      ? []
      : Array.isArray(op.columnField)
        ? op.columnField
        : [op.columnField]
  const columnFieldIndices = columnFieldsArray.map(fieldIndex)
  const colLevels = columnFieldIndices.length
  const pageFieldIndices = (op.pageFields ?? []).map(fieldIndex)
  // Dimension-field grouping (date/numeric ranges): member collection and
  // bucketing both use group labels; on save the rule (fieldIndex form) goes
  // into the journal → OOXML extLst.
  const groupings = (op.groupings ?? []).map(({ field, ...rule }) => ({
    fieldIndex: fieldIndex(field),
    ...rule,
  }))
  const groupingByField = new Map<number, PivotFieldGrouping>()
  for (const { fieldIndex: groupedField, ...rule } of groupings) {
    groupingByField.set(groupedField, rule)
  }
  // Dimension value → member label (grouped fields group first; sort orders
  // the group labels for display).
  const dimGroupValue = (
    fieldIdx: number,
    row: readonly (string | number | boolean | null)[],
  ): { label: string; sort: number | null } => {
    const grouping = groupingByField.get(fieldIdx)
    if (!grouping) return { label: String(row[fieldIdx] ?? ''), sort: null }
    return groupValue(grouping, row[fieldIdx])
  }
  const dimValue = (fieldIdx: number, row: readonly (string | number | boolean | null)[]): string =>
    dimGroupValue(fieldIdx, row).label
  // Calculated fields: field is the new data-field name (must not clash with
  // source headers); the formula is parsed once (a failed reference check
  // throws, with the message surfaced to the dialog/DSL caller).
  const valueSpecs = op.values.map((value) => {
    const isCalc = value.formula !== undefined
    if (
      isCalc &&
      fieldNames.some((name) => name.toLowerCase() === value.field.trim().toLowerCase())
    ) {
      throw new Error(t('appCalcFieldNameClash', { name: value.field }))
    }
    return {
      fieldIndex: isCalc ? -1 : fieldIndex(value.field),
      agg: value.agg,
      // Percent display modes default to 0.00% when no format is given,
      // matching dataField numFmtId=10.
      numFmt: value.numFmt ?? (value.showDataAs !== undefined ? '0.00%' : undefined),
      showDataAs: value.showDataAs,
      ...(isCalc ? { formula: value.formula, calcName: value.field } : {}),
      caption: `${AGG_CAPTIONS[value.agg]} of ${value.field}`,
      ast: isCalc ? parsePivotFormula(value.formula!, fieldNames) : null,
    }
  })
  const calcNameKeys = valueSpecs
    .filter((spec) => spec.calcName !== undefined)
    .map((spec) => spec.calcName!.trim().toLowerCase())
  if (new Set(calcNameKeys).size !== calcNameKeys.length) {
    throw new Error(t('appCalcFieldNameDuplicate'))
  }

  // Per-level member collection: one globally deduplicated list per row/column
  // level (first-seen order, i.e. each level's sharedItems order on save); the
  // combination set enumerates branches that actually exist in the hierarchy —
  // expansion order is the same under every parent group,
  // rather than first-seen within each parent.
  const levelItems: string[][] = rowFieldIndices.map(() => [])
  const comboKeys = new Set<string>()
  const colLevelItems: string[][] = columnFieldIndices.map(() => [])
  const colComboKeys = new Set<string>()
  // Sort keys for grouped-field labels (label → bucket order; pass-through
  // values are null and keep first-seen order).
  const levelSortKeys: Map<string, number | null>[] = rowFieldIndices.map(() => new Map())
  const colLevelSortKeys: Map<string, number | null>[] = columnFieldIndices.map(() => new Map())
  const joinPath = (path: readonly string[]): string => path.join('\u0000')
  // The first pass only collects members (all data rows, including ones later
  // filtered out — they must enter sharedItems and be marked hidden); the
  // combination set is enumerated over rows visible after filtering.
  for (const row of grid.slice(1)) {
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
  // Grouped levels display members in ascending bucket order (e.g. months
  // Jan…Dec, intervals small to large); pass-through values sort last keeping
  // first-seen order; ungrouped levels keep first-seen order untouched.
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
  // rowItems for journal/save = outermost members (alias of rowLevelItems[0])
  const rowItems = levelItems[0] ?? []
  if (levelItems.some((items) => items.length > 10_000)) {
    throw new Error(t('appPivotTooManyRowItems'))
  }
  if (colLevelItems.some((items) => items.length > 1_000)) {
    throw new Error(t('appPivotTooManyColItems'))
  }

  const aggregate = (
    rows: readonly (readonly (string | number | boolean | null)[])[],
    spec: (typeof valueSpecs)[number],
  ): number | null => {
    // Calculated fields: each group sums the formula's referenced source
    // fields first (empty groups count as 0), then evaluates the formula —
    // same criteria as the recompute engine's recomputePivotData.
    if (spec.ast) {
      return evaluatePivotFormula(spec.ast, (name) => {
        const refIndex = fieldIndex(name)
        return rows
          .map((row) => row[refIndex])
          .filter((value): value is number => typeof value === 'number' && Number.isFinite(value))
          .reduce((total, value) => total + value, 0)
      })
    }
    if (spec.agg === 'count') {
      return rows.filter((row) => row[spec.fieldIndex] != null && row[spec.fieldIndex] !== '')
        .length
    }
    const numbers = rows
      .map((row) => row[spec.fieldIndex])
      .filter((value): value is number => typeof value === 'number' && Number.isFinite(value))
    if (numbers.length === 0) return null
    switch (spec.agg) {
      case 'sum':
        return numbers.reduce((total, value) => total + value, 0)
      case 'average':
        return numbers.reduce((total, value) => total + value, 0) / numbers.length
      case 'max':
        return Math.max(...numbers)
      case 'min':
        return Math.min(...numbers)
    }
  }

  // Value/label filters: field caption → source column index; value filters
  // reference a values entry index.
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
  // Label filters pick candidates by member label first; value filters then
  // aggregate each candidate over the label-filtered rows before filtering
  // (same criteria as the refresh engine's reapplyPivotFilters). Filtered-out
  // members stay in the member table (sharedItems) — they are only dropped
  // from layout and aggregation.
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
  const rowVisible = (row: readonly (string | number | boolean | null)[]): boolean => {
    for (const [fieldIdx, hidden] of hiddenByField) {
      if (hidden.has(dimValue(fieldIdx, row))) return false
    }
    return true
  }
  for (const filter of filterEntries) {
    if (filter.kind !== 'value') continue
    const spec = valueSpecs[filter.dataField]
    if (!spec) throw new Error(t('appValueFilterFieldMissing', { index: filter.dataField + 1 }))
    const labelRows = grid.slice(1).filter(rowVisible)
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
  const visibleRows = grid.slice(1).filter(rowVisible)
  // Hidden member indices per level (member-table order), written as pivotField
  // hidden entries on save.
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

  // The combination set is enumerated over visible rows: combinations of
  // filtered-out members naturally never enter the layout.
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

  // Column layout lines (excluding the trailing grand-total column): data
  // columns cover all levels, with a subtotal column appended after each
  // non-leaf member; each line also records its member-label path for value
  // lookup and header baking.
  const colLines: { t: 'data' | 'default'; members: number[] }[] = []
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
  if (colLines.length > 1_000) {
    throw new Error(t('appPivotTooManyColLines'))
  }

  // Collect data rows matching (row-member prefix × column-member prefix); an
  // empty prefix array = unbounded (i.e. the various grand totals; column
  // subtotals map to shorter column prefixes). Filtered-out rows join no
  // aggregation.
  const bucketRows = (
    prefix: readonly string[],
    colPrefix: readonly string[],
  ): (string | number | boolean | null)[][] =>
    visibleRows.filter((row) => {
      for (let level = 0; level < prefix.length; level += 1) {
        if (dimValue(rowFieldIndices[level]!, row) !== prefix[level]) return false
      }
      for (let level = 0; level < colPrefix.length; level += 1) {
        if (dimValue(columnFieldIndices[level]!, row) !== colPrefix[level]) return false
      }
      return true
    })

  // Second-pass normalization for "show values as" (percent) modes: the
  // denominator re-aggregates by row prefix/column members, same criteria as
  // the recompute engine's recomputePivotData (grand-total/subtotal cells
  // included).
  const applyShowDataAs = (
    raw: number | null,
    spec: (typeof valueSpecs)[number],
    prefix: readonly string[],
    colPrefix: readonly string[],
  ): number | null => {
    if (spec.showDataAs === undefined || raw === null) return raw
    const base =
      spec.showDataAs === 'percentOfTotal'
        ? aggregate(bucketRows([], []), spec)
        : spec.showDataAs === 'percentOfRow'
          ? aggregate(bucketRows(prefix, []), spec)
          : aggregate(bucketRows([], colPrefix), spec)
    // Blank when the denominator is empty or 0 (Excel shows #DIV/0!; baking
    // chooses to blank it).
    return base === null || base === 0 ? null : raw / base
  }

  // One row's value cells: no column dimension = one column per data field;
  // with column dimensions = one column per column line (data/subtotal) + the
  // row grand-total column (the DSL guarantees a single data field here).
  const valueCells = (prefix: readonly string[]): (number | null)[] => {
    if (colLevels === 0) {
      return valueSpecs.map((spec) =>
        applyShowDataAs(aggregate(bucketRows(prefix, []), spec), spec, prefix, []),
      )
    }
    const spec = valueSpecs[0]
    if (!spec) throw new Error(t('appPivotNeedsValues'))
    return [
      ...colLinePaths.map((colPrefix) =>
        applyShowDataAs(aggregate(bucketRows(prefix, colPrefix), spec), spec, prefix, colPrefix),
      ),
      applyShowDataAs(aggregate(bucketRows(prefix, []), spec), spec, prefix, []),
    ]
  }

  // Bake the output grid: tabular layout, one column per row level; with
  // multiple levels, non-leaf levels expand recursively with a subtotal row
  // appended after their children. rowLines map one-to-one to data rows and
  // are converted to <rowItems> by xlsx-pivot-add on save.
  // Headers: 1 row without column dimensions; with them, one row per column
  // level (row-field names go in the last row's label columns), and data
  // columns write members level by level — the common prefix shared with the
  // previous data column stays blank (mirroring colItems' r encoding), subtotal
  // columns write the subtotal label on the next level's row, and the
  // grand-total column writes on the first row.
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
        // Subtotal columns: write all fixed members, with the subtotal label on
        // the next level's row.
        path.forEach((member, level) => {
          headers[level]![colOffset] = member
        })
        if (path.length < colLevels) headers[path.length]![colOffset] = t('appPivotSubtotal')
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
    headers[0]![levels + colLines.length] = 'Grand Total'
    matrix.push(...headers)
  }
  const rowLines: { t: 'data' | 'default'; members: number[] }[] = []
  // After entering a non-leaf member, its label shows on the first following
  // leaf row.
  const pendingLabels: (string | null)[] = new Array(Math.max(0, levels - 1)).fill(null)
  const emitLevel = (path: string[], indices: number[], depth: number): void => {
    levelItems[depth]!.forEach((member, memberIndex) => {
      if (!comboKeys.has(joinPath([...path, member]))) return
      if (depth === levels - 1) {
        const labels: (string | null)[] = [...pendingLabels, member]
        pendingLabels.fill(null)
        matrix.push([...labels, ...valueCells([...path, member])])
        rowLines.push({ t: 'data', members: [...indices, memberIndex] })
      } else {
        pendingLabels[depth] = member
        emitLevel([...path, member], [...indices, memberIndex], depth + 1)
        // Subtotal rows: fix the first depth+1 levels' members and write the
        // subtotal label in the next label column.
        const subtotalLabels: (string | null)[] = new Array(levels).fill(null)
        for (let level = 0; level < depth; level += 1) subtotalLabels[level] = path[level]!
        subtotalLabels[depth] = member
        subtotalLabels[depth + 1] = '\u5c0f\u8ba1'
        matrix.push([...subtotalLabels, ...valueCells([...path, member])])
        rowLines.push({ t: 'default', members: [...indices, memberIndex] })
      }
    })
  }
  emitLevel([], [], 0)
  if (rowLines.length > 20_000) {
    throw new Error(t('appPivotTooManyRowLines'))
  }
  matrix.push([...rowFieldsArray.map((_, i) => (i === 0 ? 'Grand Total' : '')), ...valueCells([])])

  const anchor = parseAddress(op.targetCell)
  const height = matrix.length
  const width = matrix[0]?.length ?? 0
  const location = {
    startRow: anchor.row,
    startColumn: anchor.column,
    endRow: anchor.row + height - 1,
    endColumn: anchor.column + width - 1,
  }
  if (
    targetSheetId === op.sheetId &&
    location.startRow <= source.endRow &&
    source.startRow <= location.endRow &&
    location.startColumn <= source.endColumn &&
    source.startColumn <= location.endColumn
  ) {
    throw new Error(t('appPivotOverlapSource'))
  }
  const targetMeta = state.file.sheets.find((sheet) => sheet.id === targetSheetId)
  for (const existing of targetMeta?.pivotRanges ?? []) {
    // When editing itself, overlapping the old output area is expected.
    if (
      relayout &&
      existing.startRow === relayout.oldBounds.startRow &&
      existing.startColumn === relayout.oldBounds.startColumn &&
      existing.endRow === relayout.oldBounds.endRow &&
      existing.endColumn === relayout.oldBounds.endColumn
    ) {
      continue
    }
    if (
      location.startRow <= existing.endRow &&
      existing.startRow <= location.endRow &&
      location.startColumn <= existing.endColumn &&
      existing.startColumn <= location.endColumn
    ) {
      throw new Error(t('appPivotOverlapExisting'))
    }
  }
  for (const pivot of state.editJournal.pivotAdds) {
    if (pivot.sheetId !== targetSheetId) continue
    if (
      location.startRow <= pivot.location.endRow &&
      pivot.location.startRow <= location.endRow &&
      location.startColumn <= pivot.location.endColumn &&
      pivot.location.startColumn <= location.endColumn
    ) {
      throw new Error(t('appPivotOverlapSession', { name: pivot.name }))
    }
  }

  // When editing an existing pivot, name is just a placeholder — the saver
  // keeps the original name from the file.
  const name = relayout ? 'PivotEdit' : (op.name ?? nextSessionPivotName(state.editJournal))
  if (
    !relayout &&
    state.editJournal.pivotAdds.some((pivot) => pivot.name.toLowerCase() === name.toLowerCase())
  ) {
    throw new Error(t('appPivotNameUsed', { name }))
  }

  if (relayout) {
    const old = relayout.oldBounds
    // Grown parts must be empty (same criteria as refresh growth); shrunken
    // parts are cleared via the union write.
    const assertEmpty = (row: number, column: number, rows: number, columns: number): void => {
      if (rows <= 0 || columns <= 0) return
      const values = targetSheet.getRange(row, column, rows, columns).getValues() as (
        string | number | boolean | null | undefined
      )[][]
      if (
        values.some((line) =>
          line.some((value) => value !== null && value !== undefined && value !== ''),
        )
      ) {
        throw new Error(t('appPivotRelayoutOverlap'))
      }
    }
    assertEmpty(old.endRow + 1, old.startColumn, location.endRow - old.endRow, width)
    assertEmpty(
      old.startRow,
      old.endColumn + 1,
      Math.min(old.endRow, location.endRow) - old.startRow + 1,
      location.endColumn - old.endColumn,
    )
    const totalRows = Math.max(height, old.endRow - old.startRow + 1)
    const totalColumns = Math.max(width, old.endColumn - old.startColumn + 1)
    const padded = Array.from({ length: totalRows }, (_, rowIndex) => {
      const row = matrix[rowIndex] ?? []
      return Array.from({ length: totalColumns }, (_, colIndex) => row[colIndex] ?? null)
    })
    targetSheet
      .getRange(anchor.row, anchor.column, totalRows, totalColumns)
      .setValues(
        padded as unknown as Parameters<ReturnType<UniverWorksheet['getRange']>['setValues']>[0],
      )
  } else {
    targetSheet
      .getRange(anchor.row, anchor.column, height, width)
      .setValues(
        matrix as unknown as Parameters<ReturnType<UniverWorksheet['getRange']>['setValues']>[0],
      )
  }

  // Apply numFmt to value columns (data rows, skip header and grand total).
  // Only applicable when there is no column field.
  if (colLevels === 0) {
    for (let vi = 0; vi < valueSpecs.length; vi += 1) {
      const fmt = valueSpecs[vi]?.numFmt
      if (!fmt) continue
      const dataHeight = height - 2 // exclude header row and grand total row
      if (dataHeight <= 0) continue
      const colIdx = anchor.column + levels + vi
      const startR = anchor.row + 1
      const rangeStr = `${columnLabel(colIdx)}${startR + 1}:${columnLabel(colIdx)}${startR + dataHeight}`
      applyFormatPatchToRange(targetSheet.getRange(rangeStr), { numberFormat: fmt })
    }
  }

  const additionPayload: Parameters<typeof recordPivotAdd>[1] = {
    sheetId: targetSheetId,
    sourceSheetId: op.sheetId,
    sourceArea: {
      startRow: source.startRow,
      startColumn: source.startColumn,
      endRow: source.endRow,
      endColumn: source.endColumn,
    },
    location,
    name,
    fieldNames,
    rowFieldIndices,
    // Single-level columns keep the legacy columnFieldIndex/columnItems form;
    // multi-level columns record columnFieldIndices + per-level members +
    // column layout lines (symmetric with the row axis).
    ...(colLevels === 1 ? { columnFieldIndex: columnFieldIndices[0]! } : {}),
    ...(pageFieldIndices.length > 0 ? { pageFieldIndices } : {}),
    rowItems,
    rowLevelItems: levelItems,
    rowLines,
    ...(colLevels === 1 ? { columnItems: colLevelItems[0]! } : {}),
    ...(colLevels >= 2 ? { columnFieldIndices, colLevelItems, colLines } : {}),
    ...(groupings.length > 0 ? { groupings } : {}),
    ...(filterEntries.length > 0 ? { filters: filterEntries, rowHiddenItems, colHiddenItems } : {}),
    values: valueSpecs.map((spec) => ({
      fieldIndex: spec.fieldIndex,
      agg: spec.agg,
      ...(spec.numFmt ? { numFmt: spec.numFmt } : {}),
      ...(spec.showDataAs ? { showDataAs: spec.showDataAs } : {}),
      ...(spec.formula !== undefined ? { formula: spec.formula, calcName: spec.calcName! } : {}),
    })),
  }
  if (relayout) {
    const newOutputRef =
      `${columnLabel(location.startColumn)}${location.startRow + 1}` +
      `:${columnLabel(location.endColumn)}${location.endRow + 1}`
    recordPivotCacheRefresh(state.editJournal, relayout.cachePath)
    recordPivotRefreshUpdate(
      state.editJournal,
      relayout.cachePath,
      targetSheetId,
      newOutputRef,
      additionPayload,
    )
    // The old in-memory definition is stale: after deletion, refresh/slicers
    // are disabled for it until save-and-reopen; saving reopens the session and
    // re-parses the newly written definition.
    state.pivotDefinitions.delete(relayout.pivotPath)
    const staleRange = targetMeta?.pivotRanges.find(
      (range) =>
        range.startRow === relayout.oldBounds.startRow &&
        range.startColumn === relayout.oldBounds.startColumn &&
        range.endRow === relayout.oldBounds.endRow &&
        range.endColumn === relayout.oldBounds.endColumn,
    )
    if (staleRange) {
      staleRange.endRow = location.endRow
      staleRange.endColumn = location.endColumn
    }
    const pivotMeta = targetMeta?.pivotTables.find((entry) => entry.path === relayout.pivotPath)
    if (pivotMeta) pivotMeta.outputRef = newOutputRef
    return
  }
  recordPivotAdd(state.editJournal, additionPayload)
}

export function pivotConfigToOpParts(
  config: OoXmlPivotConfig,
  fields: readonly PivotField[],
):
  | string
  | Pick<AddPivotOperation, 'rowFields' | 'columnField' | 'groupings' | 'filters' | 'values'> {
  // Multi-level rows: mapped to field captions in the dialog's order (outer
  // first).
  const rowFieldLabels: string[] = []
  for (const rowFieldIndex of config.rowFieldIndices) {
    const label = fields[rowFieldIndex]?.label
    if (!label) return t('appInvalidRowField')
    rowFieldLabels.push(label)
  }
  if (rowFieldLabels.length === 0) return t('appNeedRowField')
  // Multi-level columns: mapped to field captions in the dialog's order (outer
  // first); empty array = no column dimension.
  const columnFieldLabels: string[] = []
  for (const colFieldIndex of config.colFieldIndices) {
    const label = fields[colFieldIndex]?.label
    if (!label) return t('appInvalidColumnField')
    columnFieldLabels.push(label)
  }
  // Grouping modes: field index → field caption (DSL groupings reference
  // fields by caption).
  const groupingEntries: NonNullable<AddPivotOperation['groupings']> = []
  for (const { fieldIndex: groupedField, rule } of config.groupings) {
    const label = fields[groupedField]?.label
    if (!label) return t('appInvalidGroupField')
    groupingEntries.push(
      rule.kind === 'date'
        ? { kind: 'date', field: label, dateUnit: rule.dateUnit }
        : { kind: 'range', field: label, rangeStep: rule.rangeStep },
    )
  }
  // Filters: label filters by field caption; value filters apply to the
  // level-1 row field.
  const filterOps: NonNullable<AddPivotOperation['filters']> = []
  for (const { fieldIndex: filteredField, rule } of config.labelFilters) {
    const label = fields[filteredField]?.label
    if (!label) return t('appInvalidLabelFilterField')
    filterOps.push({ kind: 'label', field: label, op: rule.op, value: rule.value.trim() })
  }
  for (const { valueIndex, rule } of config.valueFilters) {
    filterOps.push({
      kind: 'value',
      field: rowFieldLabels[0]!,
      valueIndex,
      op: rule.op,
      ...(rule.count !== undefined ? { count: rule.count } : {}),
      ...(rule.from !== undefined ? { from: rule.from } : {}),
      ...(rule.to !== undefined ? { to: rule.to } : {}),
    })
  }
  let valueFields: AddPivotOperation['values']
  try {
    valueFields = config.values.map((v) => {
      // Calculated fields: field is the new field name, aggregation fixed to
      // sum, formula passed to the DSL verbatim.
      if (v.formula !== undefined) {
        return {
          field: (v.calcName ?? '').trim(),
          agg: 'sum' as const,
          formula: v.formula.trim(),
          ...(v.showDataAs !== undefined ? { showDataAs: v.showDataAs } : {}),
        }
      }
      const label = fields[v.fieldIndex]?.label
      if (!label) throw new Error(t('appInvalidValueFieldIndex', { index: v.fieldIndex }))
      return {
        field: label,
        agg: v.agg,
        ...(v.showDataAs !== undefined ? { showDataAs: v.showDataAs } : {}),
      }
    })
  } catch (error) {
    return error instanceof Error ? error.message : t('appInvalidValueField')
  }
  return {
    rowFields: rowFieldLabels,
    ...(columnFieldLabels.length > 0 ? { columnField: columnFieldLabels } : {}),
    ...(groupingEntries.length > 0 ? { groupings: groupingEntries } : {}),
    ...(filterOps.length > 0 ? { filters: filterOps } : {}),
    values: valueFields,
  }
}

export function renameChartRefsForSheet(
  state: LazyWorkbookState,
  oldName: string,
  newName: string,
): void {
  const renameSeries = <
    T extends {
      valuesRef?: string | undefined
      categoriesRef?: string | undefined
    },
  >(
    entry: T,
  ): T => ({
    ...entry,
    ...(entry.valuesRef === undefined
      ? {}
      : { valuesRef: renameRefSheet(entry.valuesRef, oldName, newName) }),
    ...(entry.categoriesRef === undefined
      ? {}
      : { categoriesRef: renameRefSheet(entry.categoriesRef, oldName, newName) }),
  })
  const renameVisual = (visual: WorkbookVisualObject): WorkbookVisualObject =>
    visual.chart
      ? { ...visual, chart: { ...visual.chart, series: visual.chart.series.map(renameSeries) } }
      : visual
  state.file.visuals.forEach((visual, at) => {
    state.file.visuals[at] = renameVisual(visual)
  })
  state.editJournal.visualAdds.forEach((visual, at) => {
    state.editJournal.visualAdds[at] = renameVisual(visual)
  })
  for (const [chartPath, edit] of state.editJournal.chartEdits) {
    if (edit.series === undefined && edit.seriesSet === undefined) continue
    state.editJournal.chartEdits.set(chartPath, {
      ...edit,
      ...(edit.series === undefined ? {} : { series: edit.series.map(renameSeries) }),
      ...(edit.seriesSet === undefined ? {} : { seriesSet: edit.seriesSet.map(renameSeries) }),
    })
  }
}

export function pivotMemberCaption(
  definition: PivotDefinition,
  field: number,
  member: number,
): string {
  const item = definition.fieldItems[field]?.[member]
  if (!item || item.x === null) return ''
  return String(definition.fields[field]?.sharedItems[item.x] ?? '')
}

export function pivotColLineLabel(
  definition: PivotDefinition,
  line: PivotDefinition['colLines'][number],
): string {
  if (line.t === 'grand') return 'Grand Total'
  if (line.t === 'blank') return ''
  const levels = line.t === 'data' ? definition.colFields.length : line.depth
  const parts: string[] = []
  for (let level = 0; level < levels; level += 1) {
    const field = definition.colFields[level]
    const member = line.members[level]
    if (field === undefined || member === null || member === undefined) continue
    parts.push(
      field === -2
        ? (definition.dataFields[member]?.name ?? '')
        : pivotMemberCaption(definition, field, member),
    )
  }
  if (line.t === 'default') parts.push(t('appPivotSubtotal'))
  // Without column dimensions, the single data column's caption is the
  // data-field name.
  if (parts.length === 0) return definition.dataFields[line.dataField]?.name ?? ''
  return parts.join(' ')
}

export function applyGrownPivotOutput(
  target: UniverWorksheet,
  grown: PivotDefinition,
  data: readonly (readonly (number | null)[])[],
  oldBounds: { startRow: number; startColumn: number; endRow: number; endColumn: number },
): string {
  const labelCols = Math.max(1, grown.rowFields.length)
  if (grown.firstDataRow !== 1 || grown.firstDataCol !== labelCols) {
    throw new Error(t('appPivotGrowUnsupported'))
  }
  const width = labelCols + (data[0]?.length ?? 0)
  const height = 1 + data.length

  // Newly occupied areas (the downward/rightward growth) must be empty to
  // avoid overwriting user content.
  const assertAreaEmpty = (row: number, column: number, rows: number, columns: number): void => {
    if (rows <= 0 || columns <= 0) return
    const values = target.getRange(row, column, rows, columns).getValues() as (
      string | number | boolean | null | undefined
    )[][]
    if (
      values.some((line) =>
        line.some((value) => value !== null && value !== undefined && value !== ''),
      )
    ) {
      throw new Error(t('appPivotGrowConflict'))
    }
  }
  const newEndRow = oldBounds.startRow + height - 1
  const newEndColumn = oldBounds.startColumn + width - 1
  assertAreaEmpty(oldBounds.endRow + 1, oldBounds.startColumn, newEndRow - oldBounds.endRow, width)
  assertAreaEmpty(
    oldBounds.startRow,
    oldBounds.endColumn + 1,
    Math.min(oldBounds.endRow, newEndRow) - oldBounds.startRow + 1,
    newEndColumn - oldBounds.endColumn,
  )

  // Header row: row-field names + each column line's caption.
  const headerRow: (string | number | null)[] = [
    ...(grown.rowFields.length === 0
      ? ['']
      : grown.rowFields.map((field) =>
          field === -2 ? 'Values' : (grown.fields[field]?.name ?? ''),
        )),
    ...grown.colLines.map((line) => pivotColLineLabel(grown, line)),
  ]

  // Row labels: data rows match the baked form (prefix members equal to the
  // previous data row stay blank), subtotal rows write the subtotal label at
  // their depth, and the grand-total row writes "Grand Total" in column 1.
  const block: (string | number | null)[][] = [headerRow]
  let previousMembers: readonly (number | null)[] | null = null
  grown.rowLines.forEach((line, index) => {
    const labels: (string | null)[] = new Array(labelCols).fill(null)
    if (line.t === 'grand') {
      labels[0] = 'Grand Total'
      previousMembers = null
    } else if (line.t === 'default') {
      for (let level = 0; level < line.depth && level < labelCols; level += 1) {
        const field = grown.rowFields[level]
        const member = line.members[level]
        if (field === undefined || member === null || member === undefined) continue
        labels[level] =
          field === -2
            ? (grown.dataFields[member]?.name ?? '')
            : pivotMemberCaption(grown, field, member)
      }
      if (line.depth < labelCols) labels[line.depth] = t('appPivotSubtotal')
    } else if (line.t === 'data') {
      for (let level = 0; level < labelCols; level += 1) {
        const field = grown.rowFields[level]
        const member = line.members[level]
        if (field === undefined || member === null || member === undefined) continue
        if (previousMembers !== null && level < labelCols - 1) {
          let samePrefix = true
          for (let k = 0; k <= level; k += 1) {
            if (previousMembers[k] !== line.members[k]) {
              samePrefix = false
              break
            }
          }
          if (samePrefix) continue
        }
        labels[level] =
          field === -2
            ? (grown.dataFields[member]?.name ?? '')
            : pivotMemberCaption(grown, field, member)
      }
      previousMembers = line.members
    }
    block.push([...labels, ...(data[index] ?? [])])
  })

  // The write area is the union of old and new: when the layout shrinks, the
  // old area's extra rows/columns are cleared too.
  const totalRows = Math.max(height, oldBounds.endRow - oldBounds.startRow + 1)
  const totalColumns = Math.max(width, oldBounds.endColumn - oldBounds.startColumn + 1)
  const padded = Array.from({ length: totalRows }, (_, rowIndex) => {
    const row = block[rowIndex] ?? []
    return Array.from({ length: totalColumns }, (_, colIndex) => row[colIndex] ?? null)
  })
  target
    .getRange(oldBounds.startRow, oldBounds.startColumn, totalRows, totalColumns)
    .setValues(
      padded as unknown as Parameters<ReturnType<UniverWorksheet['getRange']>['setValues']>[0],
    )

  return (
    `${columnLabel(oldBounds.startColumn)}${oldBounds.startRow + 1}` +
    `:${columnLabel(newEndColumn)}${newEndRow + 1}`
  )
}
