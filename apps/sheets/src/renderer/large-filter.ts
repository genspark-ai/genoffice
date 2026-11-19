/**
 * Filtering for streamed workbooks past the Full Load budget. The filter
 * panel's value list and the hidden-row set are computed over a per-column
 * projection (one dense array per filtered column, read from the sidecar in
 * screen coordinates) instead of over the Univer model, which only holds the
 * viewport. Results land in Univer's FilterModel through the same setters the
 * open-time restore uses, so rendering and the save snapshot are unchanged.
 */
import {
  CellValueType,
  createIdentifier,
  extractPureTextFromCell,
  type ICellData,
  type IRange,
  isNumeric,
  type IStyleData,
  LocaleService,
  numfmt,
  RANGE_TYPE,
} from '@univerjs/core'
import type { IFilterColumn } from '@univerjs/preset-sheets-filter'
import {
  type CustomFilterOperator,
  FilterBy,
  type FilterColumn,
  type FilterModel,
  getCustomFilterFn,
  notEquals,
  SheetsFilterService,
} from '@univerjs/sheets-filter'

import { OPEN_FILTER_PANEL_OPERATION } from './app-constants'
import { journalCellContentAt, recordFilterChange } from './edit-journal'
import { type ColumnProjection, largeFilterSheet, type ProjectionValue } from './large-filter-state'
import { fitsFullLoad, workbookCellCounts } from './load-budget'
import { readSheetRangeMapped, toUniverStyle } from './univer-sync'
import {
  journalSuppression,
  type LazyWorkbookState,
  lazySheetMeta,
  lazySheetScreenExtent,
  type UniverRuntime,
  type UniverWorksheet,
} from './univer-state'

export type FilterPredicate = (value: ProjectionValue) => boolean

/// Text the by-values filter compares (Univer: extractPureTextFromCell).
export function projectionText(value: ProjectionValue): string {
  if (value === null) return ''
  if (typeof value === 'string') return value.replace(/[\r\n]/g, '')
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  return value.toString()
}

/// Value the condition filter compares (Univer: getFilterValueForConditionalFiltering).
function conditionValue(value: ProjectionValue): string | number | null {
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  return value
}

function customFilterFn(filter: {
  val: string | number
  operator?: CustomFilterOperator | undefined
}): (value: string | number | null) => boolean {
  const compare = filter.val
  const numericCompare =
    typeof compare === 'number' || (typeof compare === 'string' && isNumeric(compare))
  if (filter.operator === 'notEqual' && !numericCompare) {
    return (value) => notEquals.fn(value as never, compare as never)
  }
  if (filter.operator) {
    if (!numericCompare) return () => false
    const fn = getCustomFilterFn(filter.operator)
    const ensured = Number(compare)
    return (value) => fn.fn(value as never, ensured as never)
  }
  const fn = getCustomFilterFn(undefined)
  return (value) => fn.fn(value as never, compare as never)
}

/// Mirrors Univer's generateFilterFn; color filters are not evaluated here.
export function buildFilterPredicate(column: IFilterColumn): FilterPredicate | null {
  if (column.filters) {
    const includeBlank = !!column.filters.blank
    const values = new Set(column.filters.filters ?? [])
    return (value) => {
      const text = projectionText(value)
      return text === '' ? includeBlank : values.has(text)
    }
  }
  if (column.customFilters) {
    const fns = column.customFilters.customFilters.map((filter) => customFilterFn(filter))
    const [first, second] = fns
    if (!first) return null
    const combined =
      second === undefined
        ? first
        : column.customFilters.and
          ? (value: string | number | null) => first(value) && second(value)
          : (value: string | number | null) => first(value) || second(value)
    return (value) => combined(conditionValue(value))
  }
  return null
}

export interface ProjectionReader {
  (row: number): ProjectionValue
}

/// Session edits win over the file projection.
export function projectionReader(
  state: LazyWorkbookState,
  sheetId: string,
  column: number,
  projection: ColumnProjection,
): ProjectionReader {
  return (row) => {
    const edited = journalCellContentAt(state.editJournal, sheetId, row, column)
    if (edited.found) return edited.value
    const value = projection.values[row - projection.startRow]
    return value === undefined ? null : value
  }
}

export function hiddenRowsForColumn(
  read: ProjectionReader,
  predicate: FilterPredicate,
  firstDataRow: number,
  lastRow: number,
): Set<number> {
  const hidden = new Set<number>()
  for (let row = firstDataRow; row <= lastRow; row += 1) {
    if (!predicate(read(row))) hidden.add(row)
  }
  return hidden
}

const PROJECTION_INDEX_WAIT_MS = 500
const PROJECTION_INDEX_DEADLINE_MS = 180_000

async function loadColumnProjection(
  state: LazyWorkbookState,
  sheetId: string,
  column: number,
  startRow: number,
  endRow: number,
): Promise<{ projection: ColumnProjection; complete: boolean } | null> {
  const sheet = lazySheetMeta(state, sheetId)
  if (!sheet) return null
  let complete = false
  const values: ProjectionValue[] = new Array<ProjectionValue>(endRow - startRow + 1).fill(null)
  const styleIndex: (number | undefined)[] = new Array<number | undefined>(
    endRow - startRow + 1,
  ).fill(undefined)
  const deadline = Date.now() + PROJECTION_INDEX_DEADLINE_MS
  let from = startRow
  while (from <= endRow) {
    const range: IRange = { startRow: from, endRow, startColumn: column, endColumn: column }
    const mapped = await readSheetRangeMapped(state, sheetId, range, sheet)
    if (!mapped) break
    for (const cell of mapped.screen.cells) {
      if (cell.row < startRow || cell.row > endRow) continue
      values[cell.row - startRow] = cell.value
      styleIndex[cell.row - startRow] = cell.styleIndex
    }
    const through = mapped.indexedThroughScreen
    if (mapped.raw.indexingComplete || (through !== null && through >= endRow)) {
      complete = true
      break
    }
    // The sidecar is still indexing (a null cutoff means it has not reached
    // this sheet yet); poll for the remainder.
    if (through !== null) from = Math.max(from, through + 1)
    if (Date.now() > deadline) break
    await new Promise((resolve) => setTimeout(resolve, PROJECTION_INDEX_WAIT_MS))
  }
  return { projection: { startRow, values, styleIndex }, complete }
}

export async function ensureColumnProjection(
  state: LazyWorkbookState,
  sheetId: string,
  column: number,
  range: IRange,
): Promise<ColumnProjection | null> {
  const entry = largeFilterSheet(state, sheetId)
  const opsStamp = state.editJournal.structuralOps.get(sheetId)?.length ?? 0
  if (entry.opsStamp !== opsStamp) {
    entry.projections.clear()
    entry.opsStamp = opsStamp
  }
  const startRow = range.startRow + 1
  const cached = entry.projections.get(column)
  if (
    cached &&
    cached.startRow === startRow &&
    cached.values.length >= range.endRow - startRow + 1
  ) {
    return cached
  }
  const loaded = await loadColumnProjection(state, sheetId, column, startRow, range.endRow)
  if (!loaded) return null
  // A projection cut short by the indexing deadline reads unindexed rows as
  // blanks; use it for this evaluation but never cache it.
  if (loaded.complete) entry.projections.set(column, loaded.projection)
  return loaded.projection
}

function filterModelFor(runtime: UniverRuntime, worksheet: UniverWorksheet): FilterModel | null {
  return (
    runtime.univer
      .__getInjector()
      .get(SheetsFilterService)
      .getFilterModel(worksheet.getSheet().getUnitId(), worksheet.getSheetId()) ?? null
  )
}

interface FilterModelInternals {
  _emitHasCriteria?: () => void
  __largeFilterPatched?: true
}

interface FilterColumnInternals {
  _filteredOutRows: Set<number> | null
}

/// The panel asks the model which rows other columns already hide; stock
/// Univer recomputes that from the (viewport-only) model, so the answer comes
/// from the evaluated sets instead.
function patchFilterModel(model: FilterModel, state: LazyWorkbookState, sheetId: string): void {
  const internals = model as unknown as FilterModelInternals
  if (internals.__largeFilterPatched) return
  internals.__largeFilterPatched = true
  model.getFilteredOutRowsExceptCol = (col: number) => {
    const union = new Set<number>()
    for (const [column, rows] of largeFilterSheet(state, sheetId).hiddenByColumn) {
      if (column === col) continue
      for (const row of rows) union.add(row)
    }
    return union
  }
}

/// Only a mutation rebuilds the skeleton (and the header buttons); a no-op
/// unhide of the header row does that without touching cell data or the
/// journal, like the open-time restore.
function refreshRows(runtime: UniverRuntime, worksheet: UniverWorksheet, headerRow: number): void {
  journalSuppression.active = true
  try {
    runtime.univerAPI.syncExecuteCommand('sheet.mutation.set-row-visible', {
      unitId: worksheet.getSheet().getUnitId(),
      subUnitId: worksheet.getSheetId(),
      ranges: [
        {
          startRow: headerRow,
          endRow: headerRow,
          startColumn: 0,
          endColumn: worksheet.getSheet().getColumnCount() - 1,
          rangeType: RANGE_TYPE.ROW,
        },
      ],
    })
  } finally {
    journalSuppression.active = false
  }
}

/// Evaluates every criteria column of the sheet's filter over projections and
/// writes the result into the model. Serialized per sheet.
export function applyLargeFilter(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  worksheet: UniverWorksheet,
  options: { record?: boolean } = {},
): Promise<void> {
  const sheetId = worksheet.getSheetId()
  const entry = largeFilterSheet(state, sheetId)
  const run = async () => {
    if (entry.pending) await entry.pending.catch(() => undefined)
    const model = filterModelFor(runtime, worksheet)
    if (!model) return
    const generation = entry.generation
    const superseded = () =>
      entry.generation !== generation || filterModelFor(runtime, worksheet) !== model
    patchFilterModel(model, state, sheetId)
    const range = model.getRange()
    const evaluated: [FilterColumn, number, Set<number>][] = []
    for (const [column, filterColumn] of model.getAllFilterColumns()) {
      const predicate = buildFilterPredicate(filterColumn.getColumnData())
      if (!predicate) continue
      const projection = await ensureColumnProjection(state, sheetId, column, range)
      if (superseded()) return
      if (!projection) continue
      const hidden = hiddenRowsForColumn(
        projectionReader(state, sheetId, column, projection),
        predicate,
        range.startRow + 1,
        range.endRow,
      )
      evaluated.push([filterColumn, column, hidden])
    }
    // Commit only once every column is in, so a failed sidecar read leaves
    // the previous evaluation in place rather than an empty one.
    const union = new Set<number>()
    entry.hiddenByColumn.clear()
    for (const [filterColumn, column, hidden] of evaluated) {
      ;(filterColumn as unknown as FilterColumnInternals)._filteredOutRows = hidden
      entry.hiddenByColumn.set(column, hidden)
      for (const row of hidden) union.add(row)
    }
    entry.active = true
    // The public setter seeds the cache and notifies the render controller.
    model.filteredOutRows = union
    ;(model as unknown as FilterModelInternals)._emitHasCriteria?.()
    refreshRows(runtime, worksheet, range.startRow)
    if (options.record !== false) recordFilterChange(state.editJournal, sheetId)
  }
  const pending = run().finally(() => {
    if (entry.pending === pending) entry.pending = null
  })
  entry.pending = pending
  return pending
}

export async function setLargeFilterCriteria(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  worksheet: UniverWorksheet,
  col: number,
  criteria: IFilterColumn | null,
): Promise<void> {
  const model = filterModelFor(runtime, worksheet)
  if (!model) return
  model.setCriteria(col, criteria, false)
  await applyLargeFilter(runtime, state, worksheet)
}

export async function clearLargeFilterCriteria(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  worksheet: UniverWorksheet,
): Promise<void> {
  const model = filterModelFor(runtime, worksheet)
  if (!model) return
  for (const [column] of model.getAllFilterColumns()) model.setCriteria(column, null, false)
  await applyLargeFilter(runtime, state, worksheet)
}

/// A FilterModel restored from the file carries criteria but no evaluated
/// sets; evaluate before the panel reads "rows hidden by other columns".
async function ensureEvaluated(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  worksheet: UniverWorksheet,
): Promise<void> {
  const model = filterModelFor(runtime, worksheet)
  if (!model) return
  const entry = largeFilterSheet(state, worksheet.getSheetId())
  const stale = model.getAllFilterColumns().some(([column]) => !entry.hiddenByColumn.has(column))
  if (stale || entry.pending) await applyLargeFilter(runtime, state, worksheet, { record: false })
}

interface VirtualCell {
  readonly row: number
  readonly col: number
  readonly rowSpan: number
  readonly colSpan: number
  readonly value: ICellData | null
}

/// The slice of Worksheet the stock value-list builder touches, served from
/// a projection so the stock code (date grouping, counts, locale) is reused.
export function virtualWorksheet(
  read: ProjectionReader,
  styleAt: (row: number) => IStyleData | undefined,
): {
  iterateByColumn(range: IRange): Iterable<VirtualCell>
  getCellRaw(row: number, col: number): ICellData | null
} {
  const cellAt = (row: number): ICellData | null => {
    const value = read(row)
    if (value === null) return null
    const style = styleAt(row)
    if (typeof value === 'boolean') {
      return { v: value ? 1 : 0, t: CellValueType.BOOLEAN, ...(style ? { s: style } : {}) }
    }
    return { v: value, ...(style ? { s: style } : {}) }
  }
  return {
    *iterateByColumn(range) {
      for (let row = range.startRow; row <= range.endRow; row += 1) {
        yield { row, col: range.startColumn, rowSpan: 1, colSpan: 1, value: cellAt(row) }
      }
    },
    getCellRaw: (row) => cellAt(row),
  }
}

const VALUES_SERVICE_NAME = 'sheets-filter.generate-filter-values.service'

export interface FilterValueTreeItem {
  title: string
  key: string
  count: number
  leaf: boolean
  checked: boolean
  children?: FilterValueTreeItem[]
  originValues?: Set<string>
}

export interface FilterValueTree {
  filterTreeItems: FilterValueTreeItem[]
  filterTreeMapCache: Map<string, string[]>
}

export interface FilterValueSource {
  iterateByColumn(
    range: IRange,
  ): Iterable<{ row: number; rowSpan?: number; value: ICellData | null }>
  getCellRaw(row: number, col: number): ICellData | null | undefined
}

export interface FilterValueTreeInput {
  readonly source: FilterValueSource
  readonly iterateRange: IRange
  readonly styles: {
    get(id: string | IStyleData | undefined): IStyleData | null | undefined | void
  }
  readonly filteredOutRowsByOtherColumns: ReadonlySet<number>
  readonly hasFilters: boolean
  readonly filterBy: number | undefined
  readonly filteredOutCount: number
  readonly alreadyChecked: ReadonlySet<string>
  readonly blankChecked: boolean
  readonly t: (key: string) => string
}

const EMPTY_KEY = 'empty'
const DATE_PATTERN = 'yyyy-mm-dd'

function compareTitles(a: string, b: string): number {
  const aNumeric = isNumeric(a)
  const bNumeric = isNumeric(b)
  if (aNumeric && bNumeric) return Number.parseFloat(a) - Number.parseFloat(b)
  if (aNumeric) return -1
  if (bNumeric) return 1
  return a.localeCompare(b)
}

function monthIndex(key: string): number {
  return Number.parseInt(key.split('-')[1] ?? '0', 10)
}

function dayIndex(key: string): number {
  return Number.parseInt(key.split('-')[2] ?? '0', 10)
}

/// Mirrors the panel's local value-list builder (getFilterTreeByValueItems in
/// sheets-filter-ui, not exported) over any cell source: the real worksheet
/// for resident sheets, a projection for large ones.
export function buildFilterValueTree(input: FilterValueTreeInput): FilterValueTree {
  const items = new Map<string, FilterValueTreeItem>()
  const treeMap = new Map<string, string[]>()
  // Color / condition filters leave no value checklist behind: start unchecked.
  const clearChecked =
    !input.hasFilters &&
    (input.filterBy === FilterBy.COLORS || input.filterBy === FilterBy.CONDITIONS) &&
    input.filteredOutCount > 0
  const checkedFor = (value: string) =>
    clearChecked
      ? false
      : input.alreadyChecked.size
        ? input.alreadyChecked.has(value)
        : !input.blankChecked
  let emptyCount = 0
  for (const cell of input.source.iterateByColumn(input.iterateRange)) {
    const { row, rowSpan = 1 } = cell
    let rowIndex = 0
    while (rowIndex < rowSpan) {
      const targetRow = row + rowIndex
      if (input.filteredOutRowsByOtherColumns.has(targetRow)) {
        rowIndex += 1
        continue
      }
      const value = cell.value ? extractPureTextFromCell(cell.value) : ''
      if (!value) {
        emptyCount += 1
        rowIndex += rowSpan
        continue
      }
      const pattern =
        cell.value?.v !== undefined && cell.value?.v !== null && !cell.value.p
          ? input.styles.get(cell.value.s as string | IStyleData | undefined)?.n?.pattern
          : undefined
      const dateInfo = pattern ? numfmt.getFormatInfo(pattern) : null
      let datePart = false
      if (dateInfo?.isDate) {
        const { year, month, day } = numfmt.getFormatDateInfo(pattern as string)
        datePart = Boolean(year || month || day)
      }
      if (pattern && dateInfo?.isDate && datePart) {
        const origin = input.source.getCellRaw(row, input.iterateRange.startColumn)?.v
        if (!origin) {
          rowIndex += 1
          continue
        }
        const [year, month, day] = numfmt
          .format(DATE_PATTERN, Number(origin))
          .split('-')
          .map(Number)
        const yearKey = `${year}`
        let yearItem = items.get(yearKey)
        if (!yearItem) {
          yearItem = {
            title: yearKey,
            key: yearKey,
            children: [],
            count: 0,
            leaf: false,
            checked: false,
          }
          items.set(yearKey, yearItem)
          treeMap.set(yearKey, [yearKey])
        }
        const monthKey = `${year}-${month}`
        let monthItem = yearItem.children?.find((item) => item.key === monthKey)
        if (!monthItem) {
          monthItem = {
            title: input.t(`sheets-filter-ui.date.${month}`),
            key: monthKey,
            children: [],
            count: 0,
            leaf: false,
            checked: false,
          }
          yearItem.children?.push(monthItem)
          treeMap.set(monthKey, [yearKey, monthKey])
        }
        const dayKey = `${year}-${month}-${day}`
        const dayItem = monthItem.children?.find((item) => item.key === dayKey)
        if (!dayItem) {
          monthItem.children?.push({
            title: `${day}`,
            key: dayKey,
            count: 1,
            originValues: new Set([value]),
            leaf: true,
            checked: checkedFor(value),
          })
          monthItem.count += 1
          yearItem.count += 1
          treeMap.set(dayKey, [yearKey, monthKey, dayKey])
        } else {
          dayItem.originValues?.add(value)
          dayItem.count += 1
          monthItem.count += 1
          yearItem.count += 1
        }
      } else {
        const item = items.get(value)
        if (!item) {
          items.set(value, {
            title: value,
            leaf: true,
            checked: checkedFor(value),
            key: value,
            count: 1,
          })
          treeMap.set(value, [value])
        } else {
          item.count += 1
        }
      }
      rowIndex += 1
    }
  }
  if (emptyCount > 0) {
    items.set(EMPTY_KEY, {
      title: input.t('sheets-filter-ui.panel.empty'),
      count: emptyCount,
      leaf: true,
      checked: clearChecked ? false : input.hasFilters ? input.blankChecked : true,
      key: EMPTY_KEY,
    })
    treeMap.set(EMPTY_KEY, [EMPTY_KEY])
  }
  const sorted = [...items.values()]
    .sort((a, b) => {
      if (a.children && !b.children) return -1
      if (!a.children && b.children) return 1
      return compareTitles(a.title, b.title)
    })
    .map((yearItem) => {
      yearItem.children
        ?.sort((a, b) => monthIndex(a.key) - monthIndex(b.key))
        .forEach((monthItem) =>
          monthItem.children?.sort((a, b) => dayIndex(a.key) - dayIndex(b.key)),
        )
      return yearItem
    })
  return { filterTreeItems: sorted, filterTreeMapCache: treeMap }
}

interface GenerateFilterValuesParams {
  unitId: string
  subUnitId: string
  filteredOutRowsByOtherColumns: number[]
  filterColumn: { filterBy?: number; filteredOutRows?: Set<number> | null } | null | undefined
  filters: boolean
  blankChecked: boolean
  iterateRange: IRange
  alreadyChecked: string[]
}

interface GenerateFilterValuesService {
  getFilterValues(params: GenerateFilterValuesParams): Promise<FilterValueTree>
}

interface InjectorLike {
  has(id: unknown): boolean
  get(id: unknown): unknown
  add(dependency: [unknown, { useValue: unknown }]): void
}

/// The panel resolves its generator service by a module-private identifier;
/// redi hands back the cached decorator for a known name (and logs an error
/// while doing so, muted here).
function valuesServiceIdentifier(): unknown {
  const error = console.error
  console.error = () => undefined
  try {
    return createIdentifier(VALUES_SERVICE_NAME)
  } finally {
    console.error = error
  }
}

const installedRuntimes = new WeakSet<object>()

/// Registers the panel's value-list generator. Desktop Univer ships none and
/// builds the list from the resident worksheet; this one does the same for
/// resident sheets and reads a projection for large ones. Called lazily from
/// the command gate, after the filter UI plugin has started.
export function installLargeFilterValues(
  runtime: UniverRuntime,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
  isLargeFilterSheet: (state: LazyWorkbookState, sheetId: string) => boolean,
  onLoading: (loading: boolean) => void,
): boolean {
  if (installedRuntimes.has(runtime)) return true
  const injector = runtime.univer.__getInjector() as unknown as InjectorLike
  const identifier = valuesServiceIdentifier()
  if (injector.has(identifier)) {
    // A remote generator is configured; large sheets are not expected there.
    installedRuntimes.add(runtime)
    return false
  }
  const localeService = injector.get(LocaleService) as LocaleService
  const service: GenerateFilterValuesService = {
    async getFilterValues(params) {
      const workbook = runtime.univerAPI.getActiveWorkbook()
      const worksheet = workbook?.getSheetBySheetId(params.subUnitId)
      if (!workbook || !worksheet) return { filterTreeItems: [], filterTreeMapCache: new Map() }
      const common = {
        iterateRange: params.iterateRange,
        styles: workbook.getWorkbook().getStyles(),
        filteredOutRowsByOtherColumns: new Set(params.filteredOutRowsByOtherColumns),
        hasFilters: params.filters,
        filterBy: params.filterColumn?.filterBy,
        filteredOutCount: params.filterColumn?.filteredOutRows?.size ?? 0,
        alreadyChecked: new Set(params.alreadyChecked.map(String)),
        blankChecked: params.blankChecked,
        t: (key: string) => localeService.t(key),
      }
      const state = lazyWorkbookRef.current
      if (!state || !isLargeFilterSheet(state, params.subUnitId)) {
        return buildFilterValueTree({
          ...common,
          source: worksheet.getSheet() as unknown as FilterValueSource,
        })
      }
      onLoading(true)
      try {
        await ensureEvaluated(runtime, state, worksheet)
        const model = filterModelFor(runtime, worksheet)
        const column = params.iterateRange.startColumn
        const projection = model
          ? await ensureColumnProjection(state, params.subUnitId, column, model.getRange())
          : null
        if (!projection) {
          return buildFilterValueTree({
            ...common,
            source: worksheet.getSheet() as unknown as FilterValueSource,
          })
        }
        const fileStyles = state.file.styles
        const source = virtualWorksheet(
          projectionReader(state, params.subUnitId, column, projection),
          (row) => {
            const index = projection.styleIndex[row - projection.startRow]
            const style = index === undefined ? undefined : fileStyles[index]
            return style?.numberFormat ? toUniverStyle(style) : undefined
          },
        )
        return buildFilterValueTree({ ...common, source })
      } finally {
        onLoading(false)
      }
    },
  }
  injector.add([identifier, { useValue: service }])
  installedRuntimes.add(runtime)
  return true
}

/// Univer sizes a new filter from the resident cells, which on a streamed
/// sheet is the viewport: a 298k-row sheet got a 267-row filter. When the
/// range was auto-detected (single cell or single row selected) extend it
/// to the sheet's data extent; a multi-row selection is kept as the user
/// drew it, as Univer and Excel do. Columns are left alone: the extent's
/// column count includes formatted-only trailing columns, and the viewport
/// already spans the data columns of ordinary sheets.
export function extendLargeFilterRange(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  worksheet: UniverWorksheet,
): boolean {
  const model = filterModelFor(runtime, worksheet)
  const extent = lazySheetScreenExtent(state, worksheet.getSheetId())
  if (!model || !extent) return false
  const selection = worksheet.getSelection()?.getActiveRange()?.getRange()
  if (selection && selection.startRow !== selection.endRow) return false
  const range = model.getRange()
  if (range.endRow >= extent.rows - 1) return false
  model.setRange({ ...range, endRow: extent.rows - 1 })
  refreshRows(runtime, worksheet, range.startRow)
  return true
}

/// Streamed, past the Full Load budget, and not a sheet added this session
/// (those live entirely in the model).
export function isLargeFilterSheet(state: LazyWorkbookState, sheetId: string): boolean {
  return (
    !state.flags.preloadComplete &&
    !state.editJournal.sheets.added.has(sheetId) &&
    !fitsFullLoad(workbookCellCounts(state.file.sheets))
  )
}

export type LargeFilterAction = (runtime: UniverRuntime, state: LazyWorkbookState) => Promise<void>

const CLOSE_FILTER_PANEL_OPERATION = 'sheet.operation.close-filter-panel'

/// Commands that evaluate criteria are taken over; the ones that only shape
/// the filter (open panel, set range, toggle, remove) return null and run in
/// Univer as usual.
export function largeFilterAction(
  commandId: string,
  params: unknown,
  worksheet: UniverWorksheet,
): LargeFilterAction | null {
  if (commandId === OPEN_FILTER_PANEL_OPERATION) return null
  switch (commandId) {
    case 'sheet.command.set-filter-criteria': {
      const { col, criteria } = (params ?? {}) as { col?: number; criteria?: IFilterColumn | null }
      if (typeof col !== 'number') return null
      return async (runtime, state) => {
        await setLargeFilterCriteria(runtime, state, worksheet, col, criteria ?? null)
        await runtime.univerAPI.executeCommand(CLOSE_FILTER_PANEL_OPERATION)
      }
    }
    case 'sheet.command.clear-filter-criteria':
      return (runtime, state) => clearLargeFilterCriteria(runtime, state, worksheet)
    case 'sheet.command.re-calc-filter':
      return (runtime, state) => applyLargeFilter(runtime, state, worksheet)
    default:
      return null
  }
}
