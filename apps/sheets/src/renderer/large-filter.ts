/** Rust evaluates AutoFilter over the streamed file plus journal edits. */
import {
  extractPureTextFromCell,
  type ICellData,
  type IRange,
  type IStyleData,
  isNumeric,
  numfmt,
  RANGE_TYPE,
} from '@univerjs/core'
import type { IFilterColumn } from '@univerjs/preset-sheets-filter'
import { FilterBy, type FilterModel, SheetsFilterService } from '@univerjs/sheets-filter'
import type { WorkbookFilterRequest, WorkbookFilterResult } from '../shared/desktop-api'
import { OPEN_FILTER_PANEL_OPERATION } from './app-constants'
import { recordFilterChange } from './edit-journal'
import { largeFilterSheet } from './large-filter-state'
import { fitsFullLoad, workbookCellCounts } from './load-budget'
import {
  journalSuppression,
  lazyFileSheetId,
  type LazyWorkbookState,
  lazySheetMeta,
  lazySheetScreenExtent,
  type UniverRuntime,
  type UniverWorksheet,
} from './univer-state'
import { screenFileRowSpans, screenToFile } from './view-transform'

export function filterModelFor(
  runtime: UniverRuntime,
  worksheet: UniverWorksheet,
): FilterModel | null {
  return (
    runtime.univer
      .__getInjector()
      .get(SheetsFilterService)
      .getFilterModel(worksheet.getSheet().getUnitId(), worksheet.getSheetId()) ?? null
  )
}

type Column = WorkbookFilterRequest['criteria'][number]
type Criteria = Map<number, IFilterColumn>
export type FilterPage = NonNullable<WorkbookFilterResult['page']>

function criteriaFor(model: FilterModel): Criteria {
  return new Map(
    model.getAllFilterColumns().map(([column, filter]) => [column, filter.getColumnData()]),
  )
}

export function filterRequest(
  state: LazyWorkbookState,
  sheetId: string,
  range: IRange,
  criteria: ReadonlyMap<number, IFilterColumn>,
  page?: { column: number; query: string; offset: number; limit: number },
): WorkbookFilterRequest {
  const sheet = lazySheetMeta(state, sheetId)
  const startRow = range.startRow + 1
  const endRow = range.endRow
  const ops = state.editJournal.structuralOps.get(sheetId) ?? []
  const overrides = new Map<number, Column['overrides']>()
  const columns = new Set(criteria.keys())
  if (page) columns.add(page.column)
  for (const column of columns) {
    const fills = state.editJournal.bulkConstantFills.get(sheetId) ?? []
    overrides.set(
      column,
      fills.flatMap((fill) => {
        if (
          column < fill.startColumn ||
          column > fill.endColumn ||
          fill.endRow < startRow ||
          fill.startRow > endRow
        )
          return []
        return [
          {
            startRow: Math.max(startRow, fill.startRow),
            endRow: Math.min(endRow, fill.endRow),
            value: fill.value,
          },
        ]
      }),
    )
  }
  for (const entry of state.editJournal.cells.get(sheetId)?.values() ?? []) {
    if (entry.hasValue && entry.row >= startRow && entry.row <= endRow) {
      overrides
        .get(entry.column)
        ?.push({ startRow: entry.row, endRow: entry.row, value: entry.value })
    }
  }
  const forColumn = (column: number): Column => ({
    fileColumn: screenToFile(ops, 'column', column),
    overrides: overrides.get(column) ?? [],
  })
  return {
    sessionId: state.file.sessionId,
    sheetId: lazyFileSheetId(state, sheetId),
    startRow,
    endRow,
    spans: screenFileRowSpans(ops, startRow, endRow, sheet?.rowCount ?? 0),
    criteria: [...criteria].flatMap(([column, data]): Column[] => {
      if (column === page?.column) return []
      const base = forColumn(column)
      if (data.filters)
        return [
          {
            ...base,
            predicate: {
              kind: 'values',
              values: data.filters.filters ?? [],
              blank: !!data.filters.blank,
            },
          },
        ]
      if (data.customFilters)
        return [
          {
            ...base,
            predicate: {
              kind: 'conditions',
              and: !!data.customFilters.and,
              conditions: data.customFilters.customFilters.map(({ val, operator }) => ({
                value: val,
                ...(operator === undefined ? {} : { operator }),
              })),
            },
          },
        ]
      return []
    }),
    ...(page ? { page: { ...page, column: forColumn(page.column) } } : {}),
  }
}

async function requestFilter(
  request: WorkbookFilterRequest,
  current: () => boolean,
): Promise<WorkbookFilterResult | undefined> {
  const deadline = Date.now() + 180_000
  while (current()) {
    const result = await window.desktopApi.filterWorkbookRows(request)
    if (!current()) return
    if (result.ready) return result
    if (Date.now() >= deadline) throw new Error('Worksheet indexing has not finished.')
    await new Promise((resolve) => setTimeout(resolve, 250))
  }
}

export async function readLargeFilterPage(
  state: LazyWorkbookState,
  model: FilterModel,
  column: number,
  query: string,
  offset: number,
  limit: number,
  current: () => boolean,
): Promise<FilterPage | undefined> {
  const range = model.getRange()
  if (range.endRow <= range.startRow) return { values: [], total: 0 }
  const request = filterRequest(state, model.subUnitId, range, criteriaFor(model), {
    column,
    query,
    offset,
    limit,
  })
  return (await requestFilter(request, current))?.page
}

export function decodeFilterHiddenRows(
  encoded: string,
  startRow: number,
  endRow: number,
): Set<number> {
  const bytes = atob(encoded)
  if (bytes.length !== Math.ceil((endRow - startRow + 1) / 8)) {
    throw new Error('Invalid worksheet filter mask.')
  }
  const hidden = new Set<number>()
  for (let i = 0; i < bytes.length; i += 1) {
    let byte = bytes.charCodeAt(i)
    while (byte) {
      const row = startRow + i * 8 + 31 - Math.clz32(byte & -byte)
      if (row <= endRow) hidden.add(row)
      byte &= byte - 1
    }
  }
  return hidden
}

/// Rebuild the grid geometry and the filter header icons through Univer's
/// standard render notification. Keep this bookkeeping out of the journal.
function refreshRows(runtime: UniverRuntime, worksheet: UniverWorksheet, headerRow: number): void {
  const suppressed = journalSuppression.active
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
    journalSuppression.active = suppressed
  }
}

const fastVisibility = new WeakSet<object>()

/// The row layout asks about all physical rows, even when the filter hides
/// almost every one. Short-circuit known filtered rows before Univer builds
/// interceptor contexts; raw hides and table filters still run normally.
function installFastRowVisibility(
  worksheet: UniverWorksheet,
  entry: ReturnType<typeof largeFilterSheet>,
): void {
  const sheet = worksheet.getSheet()
  if (fastVisibility.has(sheet)) return
  fastVisibility.add(sheet)
  const getRowFiltered = sheet.getRowFiltered.bind(sheet)
  const isRowFiltered = sheet.isRowFiltered.bind(sheet)
  sheet.getRowFiltered = (row) =>
    entry.visibility?.filteredOutRows.has(row) === true || getRowFiltered(row)
  sheet.isRowFiltered = (row) =>
    entry.visibility?.filteredOutRows.has(row) === true || isRowFiltered(row)
}

export function applyLargeFilter(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  worksheet: UniverWorksheet,
  options: {
    record?: boolean
    change?: (criteria: Criteria) => void
    current?: () => boolean
  } = {},
): Promise<void> {
  const sheetId = worksheet.getSheetId()
  const entry = largeFilterSheet(state, sheetId)
  const previous = entry.pending
  const model = filterModelFor(runtime, worksheet)
  const generation = entry.generation
  const current = () =>
    (options.current?.() ?? true) &&
    entry.generation === generation &&
    filterModelFor(runtime, worksheet) === model
  const run = async () => {
    await previous?.catch(() => undefined)
    if (!model || !current()) return
    const original = criteriaFor(model)
    const criteria = new Map(original)
    options.change?.(criteria)
    const range = model.getRange()
    const request = filterRequest(state, sheetId, range, criteria)
    let hidden = new Set<number>()
    if (request.criteria.length && request.startRow <= request.endRow) {
      const result = await requestFilter(request, current)
      if (!result) return
      hidden = decodeFilterHiddenRows(result.hidden ?? '', request.startRow, request.endRow)
    }
    if (!current()) return
    for (const column of original.keys()) {
      if (!criteria.has(column)) model.setCriteria(column, null, false)
    }
    for (const [column, data] of criteria) {
      if (original.get(column) !== data) model.setCriteria(column, data, false)
    }
    entry.visibility = model
    installFastRowVisibility(worksheet, entry)
    model.filteredOutRows = hidden
    ;(model as unknown as { _emitHasCriteria?: () => void })._emitHasCriteria?.()
    refreshRows(runtime, worksheet, range.startRow)
    if (options.record !== false) recordFilterChange(state.editJournal, sheetId)
  }
  const pending = run().finally(() => {
    if (entry.pending === pending) entry.pending = null
  })
  entry.pending = pending
  return pending
}

export function setLargeFilterCriteria(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  worksheet: UniverWorksheet,
  col: number,
  data: IFilterColumn | null,
  current?: () => boolean,
): Promise<void> {
  return applyLargeFilter(runtime, state, worksheet, {
    ...(current ? { current } : {}),
    change: (criteria) => {
      if (data) criteria.set(col, data)
      else criteria.delete(col)
    },
  })
}

export function clearLargeFilterCriteria(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  worksheet: UniverWorksheet,
  current?: () => boolean,
): Promise<void> {
  return applyLargeFilter(runtime, state, worksheet, {
    ...(current ? { current } : {}),
    change: (criteria) => criteria.clear(),
  })
}

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

export type LargeFilterAction = (
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  current: () => boolean,
) => Promise<void>

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
      return async (runtime, state, current) => {
        await setLargeFilterCriteria(runtime, state, worksheet, col, criteria ?? null, current)
        if (current()) await runtime.univerAPI.executeCommand(CLOSE_FILTER_PANEL_OPERATION)
      }
    }
    case 'sheet.command.clear-filter-criteria':
      return (runtime, state, current) =>
        clearLargeFilterCriteria(runtime, state, worksheet, current)
    case 'sheet.command.re-calc-filter':
      return (runtime, state, current) => applyLargeFilter(runtime, state, worksheet, { current })
    default:
      return null
  }
}
