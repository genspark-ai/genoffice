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
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { columnLabel, parseAddress, parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import { renameRefSheet } from '@genoffice/xlsx-gateway/domain/chart-visual'
import {
  areasOverlap,
  buildPivotLayout,
  PIVOT_SOURCE_COL_LIMIT,
  PIVOT_SOURCE_ROW_LIMIT,
  PivotLayoutError,
  pivotOutputArea,
  type PivotLayout,
  type PivotLayoutErrorCode,
} from '@genoffice/xlsx-gateway/domain/pivot-layout'
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
import type { PivotDefinition } from '@genoffice/xlsx-gateway/gateway/xlsx-pivot'
import { applyFormatPatchToRange, nextSessionPivotName, nextSessionTableName } from './univer-sync'
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

const LAYOUT_ERROR_KEYS = {
  sourceNeedsRows: 'appPivotSourceNeedsRows',
  sourceRowLimit: 'appPivotSourceRowLimit',
  sourceColLimit: 'appPivotSourceColLimit',
  headerBlank: 'appPivotHeaderBlank',
  headerDuplicate: 'appPivotHeaderDuplicate',
  fieldNotHeader: 'appPivotFieldNotHeader',
  calcFieldNameClash: 'appCalcFieldNameClash',
  calcFieldNameDuplicate: 'appCalcFieldNameDuplicate',
  valueFilterFieldMissing: 'appValueFilterFieldMissing',
  tooManyRowItems: 'appPivotTooManyRowItems',
  tooManyColItems: 'appPivotTooManyColItems',
  tooManyColLines: 'appPivotTooManyColLines',
  tooManyRowLines: 'appPivotTooManyRowLines',
  needsValues: 'appPivotNeedsValues',
} satisfies Record<PivotLayoutErrorCode, Parameters<typeof t>[0]>

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
  if (sourceRows > PIVOT_SOURCE_ROW_LIMIT) throw new Error(t('appPivotSourceRowLimit'))
  if (sourceColumns > PIVOT_SOURCE_COL_LIMIT) throw new Error(t('appPivotSourceColLimit'))

  const grid = readPivotSourceGrid(
    sourceSheet.getRange(source.startRow, source.startColumn, sourceRows, sourceColumns),
    state.file.date1904 === true,
  )
  let layout: PivotLayout
  try {
    layout = buildPivotLayout(grid, op, {
      subtotal: t('appPivotSubtotal'),
      grandTotal: 'Grand Total',
    })
  } catch (err) {
    if (err instanceof PivotLayoutError) {
      throw new Error(t(LAYOUT_ERROR_KEYS[err.code], err.params), { cause: err })
    }
    throw err
  }
  const { matrix, width, height } = layout

  const anchor = parseAddress(op.targetCell)
  const location = pivotOutputArea(anchor, layout)
  if (targetSheetId === op.sheetId && areasOverlap(location, source)) {
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
    if (areasOverlap(location, existing)) throw new Error(t('appPivotOverlapExisting'))
  }
  for (const pivot of state.editJournal.pivotAdds) {
    if (pivot.sheetId !== targetSheetId) continue
    if (areasOverlap(location, pivot.location)) {
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

  // value columns' numFmt covers the data rows only (no header, no grand total)
  const dataHeight = height - 2
  for (const { columnOffset, format } of layout.numberFormats) {
    if (dataHeight <= 0) break
    const colIdx = anchor.column + columnOffset
    const startR = anchor.row + 1
    const rangeStr = `${columnLabel(colIdx)}${startR + 1}:${columnLabel(colIdx)}${startR + dataHeight}`
    applyFormatPatchToRange(targetSheet.getRange(rangeStr), { numberFormat: format })
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
    ...layout.definition,
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
