/**
 * Table Design tab + Ctrl+T: Excel ListObject editing over both session
 * tables (journal tableAdds, rendered by Univer's table plugin) and tables
 * that came with the file (banding baked at load time, journal tableEdits).
 */
import { parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'

import type { WorkbookFile, WorkbookTableAdd, WorkbookTableEdit } from '../shared/desktop-api'
import {
  isSheetRemoved,
  journalSize,
  recordTableEdit,
  removeTableAdd,
  tableEditFor,
  updateTableAdd,
} from './edit-journal'
import { fileTableId, planFileTableRegistrations } from './file-tables'
import { t } from './i18n/locale'
import { inferContinuousRegion, type TableRegion } from './table-actions'
import {
  TABLE_OPTION_KEYS,
  areaToA1,
  areasIntersect,
  containsCell,
  isValidTableName,
  optionsOfFileTable,
  optionsOfTableAdd,
  resizeProblem,
  resolveCreateTableRange,
  totalsFunctionFor,
  uniqueColumnNames,
  type TableOptionKey,
  type TableStyleOptionFlags,
} from './table-design'
import {
  DEFAULT_TABLE_STYLE,
  TABLE_PALETTE_KEYS,
  builtinTablePalette,
  isBuiltinTableStyle,
} from './table-styles'
import { ensureLazyRangeLoaded, loadVisibleRange } from './univer-sync'
import {
  lazySheetScreenExtent,
  type LazyWorkbookState,
  type UniverRuntime,
  type UniverWorksheet,
} from './univer-state'
import { applyAiTableAdd, applySessionTableTheme } from './workbook-ops'

type FileTable = WorkbookFile['sheets'][number]['tables'][number]

export interface TableDesignContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  setMessage: (message: string) => void
  setPendingEdits: (count: number) => void
  /// Re-reads the selection so the ribbon echo follows the table change.
  refreshSelection: () => void
  /// Fires after Ctrl+T succeeds so the shell can activate Table Design.
  onTableCreated?: () => void
}

/// What the ribbon shows for the table under the selection.
export interface SelectedTableRibbon extends TableStyleOptionFlags {
  readonly key: string
  readonly name: string
  readonly rangeA1: string
  readonly style: string | null
  readonly isFile: boolean
}

type TableRef =
  | { readonly kind: 'session'; readonly sheetId: string; readonly add: WorkbookTableAdd }
  | {
      readonly kind: 'file'
      readonly sheetId: string
      readonly index: number
      readonly table: FileTable
      readonly originalName: string
      readonly edit: WorkbookTableEdit | undefined
    }

interface TableFacade {
  addTable(
    name: string,
    range: TableRegion,
    id: string,
    options?: { columns: { id: string; displayName: string }[] },
  ): Promise<boolean> | boolean
  removeTable(id: string): Promise<boolean>
  setTableName(id: string, name: string): Promise<boolean> | boolean
  setTableRange(id: string, range: TableRegion): Promise<boolean>
  addTableTheme(id: string, theme: unknown): Promise<boolean>
}

function tableApi(worksheet: UniverWorksheet): TableFacade {
  return worksheet as unknown as TableFacade
}

function sessionTableId(runtime: UniverRuntime, name: string): string | null {
  const workbook = runtime.univerAPI.getActiveWorkbook() as unknown as {
    getTableInfoByName?: (tableName: string) => { id: string } | undefined
  } | null
  return workbook?.getTableInfoByName?.(name)?.id ?? null
}

/// Open-time displayName of every file table, by `${sheetId}:${index}`: the
/// save locates the part by it after an in-session rename.
const originalNames = new WeakMap<LazyWorkbookState, Map<string, string>>()

function originalFileTableName(
  state: LazyWorkbookState,
  sheetId: string,
  index: number,
  table: FileTable,
): string {
  let names = originalNames.get(state)
  if (!names) {
    names = new Map()
    originalNames.set(state, names)
  }
  const key = `${sheetId}:${index}`
  const known = names.get(key)
  if (known !== undefined) return known
  const name = table.name ?? ''
  names.set(key, name)
  return name
}

function fileTableArea(table: FileTable): TableRegion {
  return table.range
}

export function findTableAt(
  state: LazyWorkbookState,
  sheetId: string,
  row: number,
  column: number,
): TableRef | null {
  for (const add of state.editJournal.tableAdds) {
    if (add.sheetId === sheetId && containsCell(add.area, row, column)) {
      return { kind: 'session', sheetId, add }
    }
  }
  const sheet = state.file.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheet) return null
  for (let index = 0; index < sheet.tables.length; index += 1) {
    const table = sheet.tables[index]!
    if (!containsCell(fileTableArea(table), row, column)) continue
    const originalName = originalFileTableName(state, sheetId, index, table)
    const edit = tableEditFor(state.editJournal, sheetId, originalName)
    if (edit?.remove) continue
    return { kind: 'file', sheetId, index, table, originalName, edit }
  }
  return null
}

function tablesOnSheet(state: LazyWorkbookState, sheetId: string): TableRef[] {
  const refs: TableRef[] = []
  for (const add of state.editJournal.tableAdds) {
    if (add.sheetId === sheetId) refs.push({ kind: 'session', sheetId, add })
  }
  const sheet = state.file.sheets.find((candidate) => candidate.id === sheetId)
  sheet?.tables.forEach((table, index) => {
    const originalName = originalFileTableName(state, sheetId, index, table)
    const edit = tableEditFor(state.editJournal, sheetId, originalName)
    if (!edit?.remove) refs.push({ kind: 'file', sheetId, index, table, originalName, edit })
  })
  return refs
}

function refArea(ref: TableRef): TableRegion {
  return ref.kind === 'session' ? ref.add.area : fileTableArea(ref.table)
}

function refName(ref: TableRef): string {
  return ref.kind === 'session' ? ref.add.name : (ref.table.name ?? ref.originalName)
}

function refOptions(ref: TableRef): TableStyleOptionFlags {
  return ref.kind === 'session'
    ? optionsOfTableAdd(ref.add)
    : optionsOfFileTable(ref.table, ref.edit)
}

function refStyle(ref: TableRef): string | null {
  if (ref.kind === 'session') return ref.add.style ?? DEFAULT_TABLE_STYLE
  return ref.table.styleName ?? null
}

interface ActiveTarget {
  readonly runtime: UniverRuntime
  readonly state: LazyWorkbookState
  readonly worksheet: UniverWorksheet
  readonly sheetId: string
}

function activeTarget(ctx: TableDesignContext): ActiveTarget | null {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  const worksheet = runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!runtime || !state || !worksheet) return null
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return null
  return { runtime, state, worksheet, sheetId }
}

function selectedRef(ctx: TableDesignContext): (ActiveTarget & { ref: TableRef }) | null {
  const target = activeTarget(ctx)
  if (!target) return null
  let range: ReturnType<UniverWorksheet['getRange']> | null
  try {
    range = target.runtime.univerAPI.getActiveWorkbook()?.getActiveRange() ?? null
  } catch {
    return null
  }
  if (!range) return null
  const ref = findTableAt(target.state, target.sheetId, range.getRow(), range.getColumn())
  return ref ? { ...target, ref } : null
}

export function selectedTableEcho(ctx: TableDesignContext): SelectedTableRibbon | null {
  const selected = selectedRef(ctx)
  if (!selected) return null
  const { ref } = selected
  const key =
    ref.kind === 'session'
      ? `session:${ref.sheetId}:${ref.add.name}`
      : `file:${ref.sheetId}:${ref.index}`
  return {
    key,
    name: refName(ref),
    rangeA1: areaToA1(refArea(ref)),
    style: refStyle(ref),
    isFile: ref.kind === 'file',
    ...refOptions(ref),
  }
}

export function selectedTableEquals(
  a: SelectedTableRibbon | null,
  b: SelectedTableRibbon | null,
): boolean {
  if (a === null || b === null) return a === b
  return (
    a.key === b.key &&
    a.name === b.name &&
    a.rangeA1 === b.rangeA1 &&
    a.style === b.style &&
    TABLE_OPTION_KEYS.every((option) => a[option] === b[option])
  )
}

/// Default range for the Create Table dialog: the selection, widened to its
/// data block from a single cell.
export function createTableDefaultRange(ctx: TableDesignContext): string {
  const target = activeTarget(ctx)
  const range = target?.runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
  if (!target || !range) return ''
  const selection: TableRegion = {
    startRow: range.getRow(),
    startColumn: range.getColumn(),
    endRow: range.getRow() + range.getHeight() - 1,
    endColumn: range.getColumn() + range.getWidth() - 1,
  }
  const single =
    selection.startRow === selection.endRow && selection.startColumn === selection.endColumn
  const region = single
    ? inferContinuousRegion(target.worksheet, selection.startRow, selection.startColumn)
    : null
  return areaToA1(
    region ??
      resolveCreateTableRange(selection, () => false, {
        row: selection.endRow,
        column: selection.endColumn,
      }),
  )
}

/// Returns true when the command belonged to Table Design.
export function handleTableDesignCommand(ctx: TableDesignContext, command: string): boolean {
  if (command.startsWith('table-create:')) {
    const rest = command.slice('table-create:'.length)
    const cut = rest.lastIndexOf(':')
    const ref = cut < 0 ? rest : rest.slice(0, cut)
    const headers = cut < 0 ? '1' : rest.slice(cut + 1)
    runGuarded(ctx, () => createTable(ctx, ref, headers === '1'))
    return true
  }
  if (command.startsWith('table-rename:')) {
    runGuarded(ctx, () => renameTable(ctx, command.slice('table-rename:'.length)))
    return true
  }
  if (command.startsWith('table-style:')) {
    runGuarded(ctx, () => restyleTable(ctx, command.slice('table-style:'.length)))
    return true
  }
  if (command.startsWith('table-option:')) {
    const [, key = '', value = ''] = command.split(':')
    if ((TABLE_OPTION_KEYS as readonly string[]).includes(key)) {
      runGuarded(ctx, () => toggleOption(ctx, key as TableOptionKey, value === '1'))
    }
    return true
  }
  if (command.startsWith('table-resize:')) {
    runGuarded(ctx, () => resizeTable(ctx, command.slice('table-resize:'.length)))
    return true
  }
  if (command === 'table-convert-range') {
    runGuarded(ctx, () => convertToRange(ctx))
    return true
  }
  return false
}

function runGuarded(
  ctx: TableDesignContext,
  action: () => string | void | Promise<string | void>,
): void {
  void (async () => {
    try {
      const message = await action()
      const state = ctx.lazyWorkbookRef.current
      if (state) ctx.setPendingEdits(journalSize(state.editJournal))
      if (message) ctx.setMessage(message)
    } catch (error: unknown) {
      ctx.setMessage(error instanceof Error ? error.message : t('appTableCreateFailed'))
    }
    ctx.refreshSelection()
  })()
}

/// Streamed sheets may not have the rows around the table resident yet; an
/// absent cell must not read as a free one. The padding rows are clamped to
/// the sheet extent (the loader rejects ranges past it) and a load that does
/// not complete aborts the action.
async function ensureRowsLoaded(
  ctx: TableDesignContext,
  target: ActiveTarget,
  area: TableRegion,
): Promise<void> {
  const extent = lazySheetScreenExtent(target.state, target.sheetId)
  if (!extent) return
  const range = {
    startRow: Math.max(0, area.startRow - 1),
    endRow: Math.min(extent.rows - 1, area.endRow + 1),
    startColumn: area.startColumn,
    endColumn: Math.min(extent.columns - 1, area.endColumn),
  }
  if (range.endRow < range.startRow || range.endColumn < range.startColumn) return
  const loaded = await ensureLazyRangeLoaded(
    target.runtime,
    ctx.lazyWorkbookRef,
    target.worksheet,
    range,
    ctx.setMessage,
  )
  if (!loaded) throw new Error(t('appTableRowsNotLoaded'))
}

function parseA1(ref: string): TableRegion {
  try {
    const bounds = parseRange(ref)
    return {
      startRow: bounds.startRow,
      startColumn: bounds.startColumn,
      endRow: bounds.endRow,
      endColumn: bounds.endColumn,
    }
  } catch {
    throw new Error(t('appTableInvalidRange'))
  }
}

function rowIsFree(worksheet: UniverWorksheet, row: number, area: TableRegion): boolean {
  if (row < 0) return false
  const width = area.endColumn - area.startColumn + 1
  const values = worksheet.getRange(row, area.startColumn, 1, width).getValues()[0] ?? []
  return values.every((value) => value === null || value === undefined || value === '')
}

async function createTable(
  ctx: TableDesignContext,
  ref: string,
  hasHeaders: boolean,
): Promise<string> {
  const target = activeTarget(ctx)
  if (!target) return t('appTablesNeedFile')
  const { runtime, state, worksheet, sheetId } = target
  const requested = parseA1(ref)
  await ensureRowsLoaded(ctx, target, requested)
  const room = hasHeaders
    ? { area: requested, shiftDown: false }
    : planHeaderRoom(worksheet, requested)
  for (const other of tablesOnSheet(state, sheetId)) {
    if (areasIntersect(room.area, refArea(other))) {
      throw new Error(t('appTableOverlapsSession', { name: refName(other) }))
    }
  }
  if (room.shiftDown) shiftDataDown(worksheet, requested)
  const area = room.area
  applyAiTableAdd(runtime, state, {
    op: 'add_table',
    sheetId,
    range: areaToA1(area),
    style: DEFAULT_TABLE_STYLE,
    bandedRows: true,
  })
  worksheet.getRange(area.startRow, area.startColumn, 1, 1).activate()
  // Echo first so the shell sees the table in the same render as the signal.
  ctx.refreshSelection()
  ctx.onTableCreated?.()
  return t('appTableCreated')
}

/// Excel inserts a header row above the data; without structural edits
/// (they cannot save together with a new table) the free row above is
/// used, else the data moves down into the free row below. Planned before
/// any cell moves so a later overlap failure leaves the sheet untouched.
function planHeaderRoom(
  worksheet: UniverWorksheet,
  area: TableRegion,
): { area: TableRegion; shiftDown: boolean } {
  if (rowIsFree(worksheet, area.startRow - 1, area)) {
    return { area: { ...area, startRow: area.startRow - 1 }, shiftDown: false }
  }
  if (!rowIsFree(worksheet, area.endRow + 1, area)) throw new Error(t('appTableNoRoomForHeader'))
  return { area: { ...area, endRow: area.endRow + 1 }, shiftDown: true }
}

/// A plain copy cannot re-point formulas, so blocks holding any are refused
/// (the user leaves a free row instead).
function shiftDataDown(worksheet: UniverWorksheet, area: TableRegion): void {
  const rows = area.endRow - area.startRow + 1
  const width = area.endColumn - area.startColumn + 1
  const source = worksheet.getRange(area.startRow, area.startColumn, rows, width)
  const formulas = (source as unknown as { getFormulas?: () => string[][] }).getFormulas?.() ?? []
  if (formulas.some((row) => row.some((formula) => formula.length > 0))) {
    throw new Error(t('appTableNoRoomForHeader'))
  }
  const cells = (source as unknown as { getCellDatas(): unknown[][] }).getCellDatas()
  worksheet.getRange(area.startRow + 1, area.startColumn, rows, width).setValues(cells as never)
  worksheet.getRange(area.startRow, area.startColumn, 1, width).clearContent()
}

function assertNameAvailable(state: LazyWorkbookState, name: string, self: TableRef): void {
  const lower = name.toLowerCase()
  for (const add of state.editJournal.tableAdds) {
    if (self.kind === 'session' && self.add === add) continue
    if (add.name.toLowerCase() === lower) throw new Error(t('appTableNameUsed', { name }))
  }
  for (const sheet of state.file.sheets) {
    sheet.tables.forEach((table, index) => {
      if (self.kind === 'file' && self.sheetId === sheet.id && self.index === index) return
      const original = originalFileTableName(state, sheet.id, index, table)
      if (tableEditFor(state.editJournal, sheet.id, original)?.remove) return
      if (table.name?.toLowerCase() === lower) throw new Error(t('appTableNameInFile', { name }))
    })
  }
  for (const entry of planFileTableRegistrations(state.file.sheets)) {
    if (self.kind === 'file' && entry.sheetId === self.sheetId && entry.tableIndex === self.index) {
      continue
    }
    if (entry.tableName.toLowerCase() === lower) throw new Error(t('appTableNameInFile', { name }))
  }
}

function renameTable(ctx: TableDesignContext, rawName: string): string | void {
  const selected = selectedRef(ctx)
  if (!selected) return
  const name = rawName.trim()
  const { ref, state, worksheet, runtime } = selected
  if (name === refName(ref)) return
  if (!isValidTableName(name)) throw new Error(t('appTableInvalidName'))
  assertNameAvailable(state, name, ref)
  if (ref.kind === 'session') {
    const id = sessionTableId(runtime, ref.add.name)
    updateTableAdd(state.editJournal, ref.sheetId, ref.add.name, { name })
    if (id) void tableApi(worksheet).setTableName(id, name)
    return
  }
  replaceFileTable(ref, { name })
  commitFileTable(state, ref)
  recordTableEdit(state.editJournal, { sheetId: ref.sheetId, tableName: ref.originalName, name })
  if (ref.table.headerRowCount > 0) {
    void tableApi(worksheet).setTableName(fileTableId(ref.sheetId, ref.index), name)
  }
}

/// File table entries are replaced in place so the lazy loader, the AI
/// readers and the structured-reference registry all see the change.
function replaceFileTable(
  ref: Extract<TableRef, { kind: 'file' }>,
  patch: Partial<FileTable>,
): void {
  const next = { ...ref.table, ...patch } as FileTable
  for (const key of Object.keys(patch) as (keyof FileTable)[]) {
    if (patch[key] === undefined) delete (next as Record<string, unknown>)[key]
  }
  ;(ref as { table: FileTable }).table = next
}

function commitFileTable(state: LazyWorkbookState, ref: Extract<TableRef, { kind: 'file' }>): void {
  const sheet = state.file.sheets.find((candidate) => candidate.id === ref.sheetId)
  if (sheet) (sheet.tables as FileTable[])[ref.index] = ref.table
}

function repaintSheet(ctx: TableDesignContext, target: ActiveTarget): void {
  const { state, runtime, worksheet, sheetId } = target
  state.loadedRanges.delete(sheetId)
  state.frozenStripKeys.delete(sheetId)
  void loadVisibleRange(runtime, ctx.lazyWorkbookRef, worksheet, ctx.setMessage)
}

function paletteFor(state: LazyWorkbookState, style: string): Partial<FileTable> {
  const cleared: Record<string, undefined> = {}
  for (const key of TABLE_PALETTE_KEYS) cleared[key] = undefined
  return { ...cleared, styleName: style, ...builtinTablePalette(style, state.file.themeColors) }
}

function restyleTable(ctx: TableDesignContext, style: string): string | void {
  const selected = selectedRef(ctx)
  if (!selected) return
  if (!isBuiltinTableStyle(style)) throw new Error(t('appTableInvalidStyle'))
  const { ref, state, worksheet, runtime } = selected
  if (ref.kind === 'session') {
    updateTableAdd(state.editJournal, ref.sheetId, ref.add.name, { style })
    const updated = state.editJournal.tableAdds.find((add) => add.name === ref.add.name)
    const id = sessionTableId(runtime, ref.add.name)
    if (updated && id) applySessionTableTheme(worksheet, id, updated, state.file.themeColors)
    return
  }
  replaceFileTable(ref, paletteFor(state, style))
  commitFileTable(state, ref)
  recordTableEdit(state.editJournal, { sheetId: ref.sheetId, tableName: ref.originalName, style })
  repaintSheet(ctx, selected)
}

async function toggleOption(
  ctx: TableDesignContext,
  key: TableOptionKey,
  on: boolean,
): Promise<string | void> {
  const selected = selectedRef(ctx)
  if (!selected) return
  const { ref } = selected
  if (refOptions(ref)[key] === on) return
  switch (key) {
    case 'totalsRow':
      await ensureRowsLoaded(ctx, selected, refArea(ref))
      return toggleTotalsRow(ctx, selected, on)
    case 'headerRow':
      await ensureRowsLoaded(ctx, selected, refArea(ref))
      return toggleHeaderRow(ctx, selected, on)
    default:
      return toggleStyleFlag(ctx, selected, key, on)
  }
}

function toggleStyleFlag(
  ctx: TableDesignContext,
  selected: ActiveTarget & { ref: TableRef },
  key: Exclude<TableOptionKey, 'totalsRow' | 'headerRow'>,
  on: boolean,
): void {
  const { ref, state, worksheet, runtime } = selected
  if (ref.kind === 'session') {
    updateTableAdd(state.editJournal, ref.sheetId, ref.add.name, { [key]: on })
    const updated = state.editJournal.tableAdds.find((add) => add.name === ref.add.name)
    const id = sessionTableId(runtime, ref.add.name)
    if (updated && id) applySessionTableTheme(worksheet, id, updated, state.file.themeColors)
    return
  }
  recordTableEdit(state.editJournal, {
    sheetId: ref.sheetId,
    tableName: ref.originalName,
    [key]: on,
  })
  if (key === 'bandedRows' || key === 'bandedColumns') {
    const patch: Partial<FileTable> =
      key === 'bandedRows' ? { showRowStripes: on } : { showColumnStripes: on }
    if (key === 'bandedColumns' && on && !ref.table.columnStripeFill && ref.table.stripeFill) {
      patch.columnStripeFill = ref.table.stripeFill
    }
    replaceFileTable(ref, patch)
    commitFileTable(state, ref)
    repaintSheet(ctx, selected)
  }
}

function toggleTotalsRow(
  ctx: TableDesignContext,
  selected: ActiveTarget & { ref: TableRef },
  on: boolean,
): string | void {
  const { ref, worksheet } = selected
  const area = refArea(ref)
  const width = area.endColumn - area.startColumn + 1
  if (on) {
    const row = area.endRow + 1
    if (!rowIsFree(worksheet, row, area)) throw new Error(t('appTableTotalsRowBlocked'))
    const options = refOptions(ref)
    const dataStart = area.startRow + (options.headerRow ? 1 : 0)
    const dataRows = area.endRow - dataStart + 1
    const lastColumn = area.endColumn
    const values =
      dataRows > 0
        ? worksheet
            .getRange(dataStart, lastColumn, dataRows, 1)
            .getValues()
            .map((cells) => cells[0])
        : []
    const fn = totalsFunctionFor(values)
    if (width > 1) {
      worksheet.getRange(row, area.startColumn, 1, 1).setValue(t('appTableTotalLabel'))
    }
    if (dataRows > 0) {
      const columnRef = areaToA1({
        startRow: dataStart,
        endRow: area.endRow,
        startColumn: lastColumn,
        endColumn: lastColumn,
      })
      worksheet.getRange(row, lastColumn, 1, 1).setValue({ f: `=SUBTOTAL(${fn},${columnRef})` })
    }
    worksheet.getRange(row, area.startColumn, 1, width).setFontWeight('bold')
    applyAreaChange(ctx, selected, { ...area, endRow: row }, { totalsRow: true })
    return
  }
  const totalsRow = area.endRow
  const strip = worksheet.getRange(totalsRow, area.startColumn, 1, width)
  strip.clearContent()
  strip.setFontWeight(null)
  applyAreaChange(ctx, selected, { ...area, endRow: area.endRow - 1 }, { totalsRow: false })
}

function toggleHeaderRow(
  ctx: TableDesignContext,
  selected: ActiveTarget & { ref: TableRef },
  on: boolean,
): string | void {
  const { ref, state, worksheet, runtime } = selected
  const area = refArea(ref)
  const width = area.endColumn - area.startColumn + 1
  const columnNames = ref.kind === 'session' ? ref.add.columnNames : (ref.table.columns ?? [])
  if (on) {
    const row = area.startRow - 1
    if (!rowIsFree(worksheet, row, area)) throw new Error(t('appTableHeaderRowBlocked'))
    const names = uniqueColumnNames(
      Array.from({ length: width }, (_, index) => columnNames[index] ?? ''),
    )
    worksheet.getRange(row, area.startColumn, 1, width).setValues([names])
    applyAreaChange(
      ctx,
      selected,
      { ...area, startRow: row },
      { headerRow: true, columnNames: names },
    )
    const next: TableRegion = { ...area, startRow: row }
    const hasTotals = refOptions(ref).totalsRow
    if (ref.kind === 'session') {
      const id = `ai-table-${Date.now().toString(36)}`
      const updated = state.editJournal.tableAdds.find((add) => add.name === ref.add.name)
      void Promise.resolve(
        tableApi(worksheet).addTable(ref.add.name, dataRange(next, hasTotals), id),
      )
        .then(() => {
          if (updated) applySessionTableTheme(worksheet, id, updated, state.file.themeColors)
        })
        .catch(() => undefined)
    } else {
      const id = fileTableId(ref.sheetId, ref.index)
      void Promise.resolve(
        tableApi(worksheet).addTable(refName(ref), dataRange(next, hasTotals), id, {
          columns: names.map((displayName, index) => ({ id: `${id}-col-${index}`, displayName })),
        }),
      )
        .then(() => tableApi(worksheet).addTableTheme(id, { name: `plain-${id}` }))
        .catch(() => undefined)
    }
    return
  }
  const headerCells = worksheet.getRange(area.startRow, area.startColumn, 1, width)
  const liveNames = uniqueColumnNames(headerCells.getValues()[0] ?? [])
  headerCells.clearContent()
  const id =
    ref.kind === 'session'
      ? sessionTableId(runtime, ref.add.name)
      : fileTableId(ref.sheetId, ref.index)
  if (id)
    void tableApi(worksheet)
      .removeTable(id)
      .catch(() => undefined)
  applyAreaChange(
    ctx,
    selected,
    { ...area, startRow: area.startRow + 1 },
    { headerRow: false, columnNames: liveNames },
  )
}

/// Univer's table range: the data block plus header, never the totals band
/// (so `Table[Column]` sums the data only).
function dataRange(area: TableRegion, totalsRow: boolean): TableRegion {
  return totalsRow ? { ...area, endRow: area.endRow - 1 } : area
}

function applyAreaChange(
  ctx: TableDesignContext,
  selected: ActiveTarget & { ref: TableRef },
  area: TableRegion,
  patch: { totalsRow?: boolean; headerRow?: boolean; columnNames?: string[] },
): void {
  const { ref, state, worksheet, runtime } = selected
  if (ref.kind === 'session') {
    updateTableAdd(state.editJournal, ref.sheetId, ref.add.name, { area, ...patch })
    const updated = state.editJournal.tableAdds.find((add) => add.name === ref.add.name)
    const id = sessionTableId(runtime, ref.add.name)
    if (id && updated && patch.headerRow === undefined) {
      void tableApi(worksheet)
        .setTableRange(id, dataRange(area, optionsOfTableAdd(updated).totalsRow))
        .catch(() => undefined)
    }
    return
  }
  const fileTablePatch: Partial<FileTable> = { range: area }
  if (patch.totalsRow !== undefined) fileTablePatch.totalsRowCount = patch.totalsRow ? 1 : 0
  if (patch.headerRow !== undefined) fileTablePatch.headerRowCount = patch.headerRow ? 1 : 0
  if (patch.columnNames) fileTablePatch.columns = patch.columnNames
  replaceFileTable(ref, fileTablePatch)
  commitFileTable(state, ref)
  recordTableEdit(state.editJournal, {
    sheetId: ref.sheetId,
    tableName: ref.originalName,
    area,
    ...(patch.totalsRow === undefined ? {} : { totalsRow: patch.totalsRow }),
    ...(patch.headerRow === undefined ? {} : { headerRow: patch.headerRow }),
    ...(patch.columnNames === undefined ? {} : { columnNames: patch.columnNames }),
  })
  if (patch.headerRow === undefined && ref.table.headerRowCount > 0) {
    void tableApi(worksheet)
      .setTableRange(
        fileTableId(ref.sheetId, ref.index),
        dataRange(area, (ref.table.totalsRowCount ?? 0) > 0),
      )
      .catch(() => undefined)
  }
  repaintSheet(ctx, selected)
}

async function resizeTable(ctx: TableDesignContext, rawRef: string): Promise<string | void> {
  const selected = selectedRef(ctx)
  if (!selected) return
  const { ref, state, worksheet } = selected
  const current = refArea(ref)
  const next = parseA1(rawRef)
  if (next.endColumn - next.startColumn + 1 > 1_000) throw new Error(t('appTableTooWide'))
  const options = refOptions(ref)
  await ensureRowsLoaded(ctx, selected, {
    startRow: Math.min(current.startRow, next.startRow),
    endRow: Math.max(current.endRow, next.endRow),
    startColumn: Math.min(current.startColumn, next.startColumn),
    endColumn: Math.max(current.endColumn, next.endColumn),
  })
  const problem = resizeProblem(current, next, options)
  if (problem === 'header-row-moved') throw new Error(t('appTableResizeHeaderRow'))
  if (problem === 'no-overlap') throw new Error(t('appTableResizeOverlap'))
  if (problem === 'no-data-rows') throw new Error(t('appTableNeedsRows'))
  for (const other of tablesOnSheet(state, ref.sheetId)) {
    const same =
      other.kind === ref.kind &&
      (other.kind === 'session'
        ? other.add === (ref as { add?: WorkbookTableAdd }).add
        : other.index === (ref as { index?: number }).index)
    if (!same && areasIntersect(next, refArea(other))) {
      throw new Error(t('appTableOverlapsSession', { name: refName(other) }))
    }
  }
  if (
    current.startRow === next.startRow &&
    current.endRow === next.endRow &&
    current.startColumn === next.startColumn &&
    current.endColumn === next.endColumn
  ) {
    return
  }
  // Every refusal happens before the first cell write.
  if (options.totalsRow) assertTotalsRowCanMove(worksheet, current, next)
  const width = next.endColumn - next.startColumn + 1
  let columnNames: string[]
  if (options.headerRow) {
    const headers =
      worksheet.getRange(next.startRow, next.startColumn, 1, width).getValues()[0] ?? []
    columnNames = uniqueColumnNames(headers)
    columnNames.forEach((name, index) => {
      if (name !== String(headers[index] ?? '').trim()) {
        worksheet.getRange(next.startRow, next.startColumn + index, 1, 1).setValue(name)
      }
    })
  } else {
    // No header cells to read: the stored names follow their sheet column,
    // new columns take ColumnN.
    const previous = ref.kind === 'session' ? ref.add.columnNames : (ref.table.columns ?? [])
    columnNames = uniqueColumnNames(
      Array.from({ length: width }, (_, index) => {
        const previousIndex = next.startColumn + index - current.startColumn
        return previousIndex >= 0 ? (previous[previousIndex] ?? '') : ''
      }),
    )
  }
  if (options.totalsRow) {
    moveTotalsRow(worksheet, current, next)
    refreshTotalsFormulas(worksheet, next, options.headerRow)
  }
  if (ref.kind === 'session') {
    const totalsRow = optionsOfTableAdd(ref.add).totalsRow
    updateTableAdd(state.editJournal, ref.sheetId, ref.add.name, { area: next, columnNames })
    const id = sessionTableId(selected.runtime, ref.add.name)
    if (id) {
      void tableApi(worksheet)
        .setTableRange(id, dataRange(next, totalsRow))
        .catch(() => undefined)
    }
    return t('appTableResized')
  }
  applyAreaChange(ctx, selected, next, { columnNames })
  return t('appTableResized')
}

/// The totals band stays the table's last row. When that row is free the
/// totals cells follow their sheet column there. A shrink onto existing
/// cells follows Excel: they simply become the totals band. A grow over
/// occupied rows would leave the old SUBTOTAL cells inside the data body,
/// so it is refused.
function assertTotalsRowCanMove(
  worksheet: UniverWorksheet,
  current: TableRegion,
  next: TableRegion,
): void {
  if (next.endRow <= current.endRow) return
  if (!rowIsFree(worksheet, next.endRow, { ...next, startRow: next.endRow })) {
    throw new Error(t('appTableTotalsRowBlocked'))
  }
}

function moveTotalsRow(worksheet: UniverWorksheet, current: TableRegion, next: TableRegion): void {
  if (current.endRow === next.endRow) return
  const destination: TableRegion = { ...next, startRow: next.endRow }
  if (!rowIsFree(worksheet, next.endRow, destination)) return
  const oldWidth = current.endColumn - current.startColumn + 1
  const source = worksheet.getRange(current.endRow, current.startColumn, 1, oldWidth)
  const cells = (source as unknown as { getCellDatas(): unknown[][] }).getCellDatas()[0] ?? []
  source.clearContent()
  source.setFontWeight(null)
  const width = next.endColumn - next.startColumn + 1
  const moved = Array.from({ length: width }, (_, index) => {
    const previousIndex = next.startColumn + index - current.startColumn
    return previousIndex >= 0 ? (cells[previousIndex] ?? null) : null
  })
  const target = worksheet.getRange(next.endRow, next.startColumn, 1, width)
  target.setValues([moved] as never)
  target.setFontWeight('bold')
}

/// A resize moves the data block under the totals row: every SUBTOTAL the
/// totals row carries is re-pointed at its column's new data range.
function refreshTotalsFormulas(
  worksheet: UniverWorksheet,
  area: TableRegion,
  headerRow: boolean,
): void {
  const totalsRow = area.endRow
  const dataStart = area.startRow + (headerRow ? 1 : 0)
  const dataEnd = totalsRow - 1
  if (dataEnd < dataStart) return
  for (let column = area.startColumn; column <= area.endColumn; column += 1) {
    const cell = worksheet.getRange(totalsRow, column, 1, 1)
    const formula = (cell as unknown as { getFormula?: () => string }).getFormula?.() ?? ''
    const match = /^=SUBTOTAL\((\d+),/.exec(formula)
    if (!match) continue
    const columnRef = areaToA1({
      startRow: dataStart,
      endRow: dataEnd,
      startColumn: column,
      endColumn: column,
    })
    cell.setValue({ f: `=SUBTOTAL(${match[1]},${columnRef})` })
  }
}

function convertToRange(ctx: TableDesignContext): string | void {
  const selected = selectedRef(ctx)
  if (!selected) return
  const { ref, state, worksheet, runtime } = selected
  if (ref.kind === 'session') {
    const id = sessionTableId(runtime, ref.add.name)
    removeTableAdd(state.editJournal, ref.sheetId, ref.add.name)
    if (id)
      void tableApi(worksheet)
        .removeTable(id)
        .catch(() => undefined)
    return t('appTableConverted')
  }
  if (ref.originalName.length === 0) throw new Error(t('appTableFileUnnamed'))
  recordTableEdit(state.editJournal, {
    sheetId: ref.sheetId,
    tableName: ref.originalName,
    remove: true,
  })
  // Tombstone: no header (keeps it out of the formula registry), no style
  // (the loader paints nothing); the journal's remove flag hides it from
  // the Table Design lookup.
  const hadHeader = ref.table.headerRowCount > 0
  const cleared: Record<string, undefined> = {}
  for (const key of TABLE_PALETTE_KEYS) cleared[key] = undefined
  replaceFileTable(ref, {
    ...cleared,
    styleName: undefined,
    headerRowCount: 0,
    showRowStripes: false,
    showColumnStripes: false,
  })
  commitFileTable(state, ref)
  if (hadHeader) {
    void tableApi(worksheet)
      .removeTable(fileTableId(ref.sheetId, ref.index))
      .catch(() => undefined)
  }
  repaintSheet(ctx, selected)
  return t('appTableConverted')
}
