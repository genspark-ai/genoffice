/**
 * Pivot-table and slicer actions (create dialog, PivotTable Fields pane,
 * refresh, slicer panels). Extracted from App.tsx; the App component passes a
 * PivotActionContext built fresh per call so refs and state never go stale.
 */
import { columnLabel, parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  applyPivotSlicer,
  growPivotDefinition,
  recomputePivotData,
} from '@genoffice/xlsx-gateway/domain/pivot-engine'
import {
  timelineDomainOf,
  timelineSelection,
  type MonthKey,
} from '@genoffice/xlsx-gateway/domain/pivot-timeline'
import {
  defaultDataFieldCaption,
  isPivotAggregation,
} from '@genoffice/xlsx-gateway/domain/pivot-value-modes'
import type { WorkbookFile, WorkbookPivotAdd, WorkbookPivotDefinition } from '../shared/desktop-api'
import { journalSize, recordPivotCacheRefresh, recordPivotRefreshUpdate } from './edit-journal'
import { t } from './i18n/locale'
import { isCalendarDatePattern } from './numfmt-fix'
import type { PivotCreateRequest } from './PivotDialog'
import {
  defaultPivotModel,
  modelToConfig,
  validatePivotModel,
  type OoXmlPivotConfig,
  type PivotField,
  type PivotLayoutModel,
  type PivotValueSpec,
} from './pivot-field-model'
import type { SlicerMember, SlicerUiState } from './SlicerPanel'
import type { TimelineUiState } from './TimelinePanel'
import { resolvePivotSource } from './pivot-source'
import type { LazyWorkbookState, UniverRuntime, UniverWorksheet } from './univer-state'
import {
  applyAiPivotAdd,
  applyGrownPivotOutput,
  pivotConfigToOpParts,
  pivotPageRows,
  readPivotSourceGrid,
  type PivotRelayoutTarget,
} from './workbook-ops'

/// Editing an existing pivot from the Fields pane: context locked when the
/// pane opens, refreshed after every apply.
export interface PivotEditContext {
  readonly key: string
  readonly sourceSheetId: string
  readonly targetSheetId: string
  readonly sourceRange: string
  /// Top-left of the whole block (report-filter rows included).
  readonly targetCell: string
  readonly fields: PivotField[]
  relayout: PivotRelayoutTarget
}

/// What the Fields pane edits: the field list and the current layout.
export interface PivotEditSeed {
  readonly key: string
  readonly fields: readonly PivotField[]
  readonly sourceRange: string
  readonly targetCell: string
  readonly model: PivotLayoutModel
}

/// Non-null while the "Insert Slicer" field picker is open.
export interface SlicerPickerState {
  sheetId: string
  pivotPath: string
  fields: readonly { field: number; name: string }[]
}

/// Non-null while the "Insert Timeline" field picker is open.
export interface TimelinePickerState {
  sheetId: string
  pivotPath: string
  fields: readonly { field: number; name: string }[]
}

/** The App refs/state the pivot actions need; built fresh per call. */
export interface PivotActionContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  pivotEditContextRef: { current: PivotEditContext | null }
  slicers: readonly SlicerUiState[]
  slicerPicker: SlicerPickerState | null
  setSlicers: (update: (current: readonly SlicerUiState[]) => readonly SlicerUiState[]) => void
  setSlicerPicker: (value: SlicerPickerState | null) => void
  timelines: readonly TimelineUiState[]
  timelinePicker: TimelinePickerState | null
  setTimelines: (
    update: (current: readonly TimelineUiState[]) => readonly TimelineUiState[],
  ) => void
  setTimelinePicker: (value: TimelinePickerState | null) => void
  setMessage: (message: string) => void
  setPendingEdits: (count: number) => void
}

type FilePivotMeta = WorkbookFile['sheets'][number]['pivotTables'][number]

/// The pivot under the selection: one from the file (possibly with a pending
/// layout edit journaled this session) or one created this session.
export type PivotHit =
  | {
      readonly kind: 'file'
      readonly sheetId: string
      readonly pivot: FilePivotMeta
      readonly definition: WorkbookPivotDefinition | undefined
      readonly relayout: WorkbookPivotAdd | undefined
    }
  | { readonly kind: 'session'; readonly sheetId: string; readonly add: WorkbookPivotAdd }

function pivotSourceValueFields(definition: WorkbookPivotDefinition): string[] {
  return definition.dataFields.flatMap(({ field }) => {
    const source = definition.fields[field]
    return source && source.formula === undefined ? [source.name] : []
  })
}

function activeSelection(ctx: PivotActionContext): {
  sheetId: string
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
} | null {
  const runtime = ctx.univerRef.current
  if (!runtime) return null
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!worksheet || !range) return null
  return {
    sheetId: worksheet.getSheetId(),
    startRow: range.getRow(),
    startColumn: range.getColumn(),
    endRow: range.getRow() + range.getHeight() - 1,
    endColumn: range.getColumn() + range.getWidth() - 1,
  }
}

function hitsBlock(
  selection: NonNullable<ReturnType<typeof activeSelection>>,
  body: { startRow: number; startColumn: number; endRow: number; endColumn: number },
  pageRows: number,
): boolean {
  return (
    selection.startRow <= body.endRow &&
    selection.endRow >= body.startRow - pageRows &&
    selection.startColumn <= body.endColumn &&
    selection.endColumn >= body.startColumn
  )
}

export function findPivotAtSelection(ctx: PivotActionContext): PivotHit | null {
  const state = ctx.lazyWorkbookRef.current
  const selection = activeSelection(ctx)
  if (!state || !selection) return null
  const sheetMeta = state.file.sheets.find((sheet) => sheet.id === selection.sheetId)
  if (!sheetMeta) return null
  for (const pivot of sheetMeta.pivotTables) {
    const definition = state.pivotDefinitions.get(pivot.path)
    const relayout =
      pivot.cachePath === null
        ? undefined
        : state.editJournal.pivotRefreshUpdates.get(pivot.cachePath)?.relayout
    const pageCount = relayout?.pageFieldIndices?.length ?? definition?.pageFields.length ?? 0
    if (hitsBlock(selection, parseRange(pivot.outputRef), pivotPageRows(pageCount))) {
      return { kind: 'file', sheetId: selection.sheetId, pivot, definition, relayout }
    }
  }
  for (const add of state.editJournal.pivotAdds) {
    if (add.sheetId !== selection.sheetId) continue
    if (hitsBlock(selection, add.location, pivotPageRows(add.pageFieldIndices?.length ?? 0))) {
      return { kind: 'session', sheetId: selection.sheetId, add }
    }
  }
  return null
}

/// Stable identity of the pivot under the selection (drives the Fields pane).
export function pivotSelectionKey(ctx: PivotActionContext): string | null {
  const hit = findPivotAtSelection(ctx)
  if (!hit) return null
  return hit.kind === 'file' ? hit.pivot.path : `session:${hit.add.name}`
}

/// Slicers and timelines bind to a file pivot with a loaded definition; a
/// pivot created this session qualifies after save-and-reopen. When it returns
/// null the reason was already shown.
export function filePivotAtSelection(ctx: PivotActionContext): {
  sheetId: string
  pivot: FilePivotMeta
  definition: WorkbookPivotDefinition
} | null {
  const hit = findPivotAtSelection(ctx)
  if (!hit) {
    ctx.setMessage(t('appCursorNotInPivot'))
    return null
  }
  if (hit.kind === 'session' || !hit.definition) {
    ctx.setMessage(t('appPivotDefNotLoadedSave'))
    return null
  }
  return { sheetId: hit.sheetId, pivot: hit.pivot, definition: hit.definition }
}

/// Recomputes every pivot table on the sheet from fresh source values and
/// writes the data area back (journaled like manual edits). The layout is
/// the file's own — drift fails closed inside the engine.
export function refreshPivotTables(ctx: PivotActionContext, sheetId: string): number {
  const state = ctx.lazyWorkbookRef.current
  const runtime = ctx.univerRef.current
  if (!state || !runtime) throw new Error(t('appOpenXlsxFirst'))
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const sheetMeta = state.file.sheets.find((sheet) => sheet.id === sheetId)
  if (!workbook || !sheetMeta) throw new Error(`Unknown sheet: ${sheetId}`)
  if (sheetMeta.pivotTables.length === 0) {
    throw new Error(t('appSheetNoPivot'))
  }
  if (!state.formulaMode || !state.flags.preloadComplete) {
    throw new Error(t('appPivotNeedsFullLoad'))
  }
  const target = workbook.getSheetBySheetId(sheetId)
  if (!target) throw new Error(`Unknown sheet: ${sheetId}`)
  let refreshed = 0
  for (const pivot of sheetMeta.pivotTables) {
    const definition = state.pivotDefinitions.get(pivot.path)
    if (!definition) {
      throw new Error(t('appPivotDefNotLoaded'))
    }
    const sourceSheet = workbook
      .getSheets()
      .find((sheet) => sheet.getSheetName() === definition.sourceSheet)
    if (!sourceSheet) {
      throw new Error(t('appPivotSourceSheetMissing', { name: definition.sourceSheet }))
    }
    const sourceValues = readPivotSourceGrid(
      sourceSheet.getRange(definition.sourceRef),
      ctx.lazyWorkbookRef.current?.file.date1904 === true,
      pivotSourceValueFields(definition),
    )
    // (3) Automatic layout growth: when source data has new categories outside
    // the cache, first fold the new members into the layout (in memory), then
    // recompute; without new categories, take the existing data-area-only
    // write-back path.
    const growth = growPivotDefinition(definition, sourceValues)
    const grown = growth.definition
    const { data } = recomputePivotData(grown, sourceValues)
    const bounds = parseRange(definition.outputRef)
    if (!growth.grown) {
      const startRow = bounds.startRow + definition.firstDataRow
      const startColumn = bounds.startColumn + definition.firstDataCol
      if (
        startRow + data.length - 1 !== bounds.endRow ||
        startColumn + (data[0]?.length ?? 0) - 1 !== bounds.endColumn
      ) {
        throw new Error(t('appPivotLayoutMismatch'))
      }
      const dataRange =
        `${columnLabel(startColumn)}${startRow + 1}` +
        `:${columnLabel(bounds.endColumn)}${bounds.endRow + 1}`
      target
        .getRange(dataRange)
        .setValues(data.map((line) => line.map((value) => ({ v: value, f: null, si: null }))))
    } else {
      const newOutputRef = applyGrownPivotOutput(target, grown, data, bounds)
      if (pivot.cachePath !== null) {
        // On save, the pivotTableDefinition's location ref is widened to the
        // new area; OOXML write-back of rowItems/colItems and sharedItems is
        // deferred — refreshOnLoad makes Excel rebuild them on open (see the
        // pivot-engine TODO).
        recordPivotRefreshUpdate(state.editJournal, pivot.cachePath, sheetId, newOutputRef)
      }
      // The grown definition (with the new outputRef) takes over subsequent
      // refreshes and edit protection.
      state.pivotDefinitions.set(pivot.path, {
        ...grown,
        outputRef: newOutputRef,
      } as WorkbookPivotDefinition)
      const newBounds = parseRange(newOutputRef)
      const staleRange = sheetMeta.pivotRanges.find(
        (range) =>
          range.startRow === bounds.startRow &&
          range.startColumn === bounds.startColumn &&
          range.endRow === bounds.endRow &&
          range.endColumn === bounds.endColumn,
      )
      if (staleRange) {
        staleRange.endRow = newBounds.endRow
        staleRange.endColumn = newBounds.endColumn
      }
      pivot.outputRef = newOutputRef
    }
    if (pivot.cachePath !== null) {
      recordPivotCacheRefresh(state.editJournal, pivot.cachePath)
    }
    refreshed += 1
  }
  ctx.setPendingEdits(journalSize(state.editJournal))
  return refreshed
}

/// Excel's "numeric field": every non-blank cell is a number and the column is
/// not date-formatted (dates group on Rows, not in Values).
function columnIsNumeric(
  worksheet: UniverWorksheet,
  startRow: number,
  column: number,
  rows: number,
): boolean {
  if (rows <= 0) return false
  const range = worksheet.getRange(startRow, column, rows, 1)
  const values = range.getRawValues() as unknown[][]
  const patterns = range.getNumberFormats()
  let seen = false
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index]?.[0]
    if (value === null || value === undefined || value === '') continue
    if (typeof value !== 'number') return false
    if (isCalendarDatePattern(patterns[index]?.[0] ?? '')) return false
    seen = true
  }
  return seen
}

/// Source headers of a range (the active selection by default), with Excel's
/// numeric flag per field.
export function pivotFieldOptions(
  ctx: PivotActionContext,
  sourceRange?: string,
  worksheet?: UniverWorksheet,
): PivotField[] {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const sheet = worksheet ?? workbook?.getActiveSheet()
  const range = sourceRange ? sheet?.getRange(sourceRange) : workbook?.getActiveRange()
  if (!sheet || !range || range.getHeight() < 2) return []
  const headerRow =
    sheet.getRange(range.getRow(), range.getColumn(), 1, range.getWidth()).getValues()[0] ?? []
  const start = range.getColumn()
  const dataRows = range.getHeight() - 1
  return headerRow.slice(0, 200).map((header, offset) => ({
    label:
      header === null || header === undefined || header === ''
        ? t('appColumnLabel', { col: columnLabel(start + offset) })
        : String(header),
    colIndex: start + offset,
    numeric: columnIsNumeric(sheet, range.getRow() + 1, start + offset, dataRows),
  }))
}

/// Create PivotTable: Excel's default placement (first text field on Rows,
/// first numeric field on Values) bakes the initial grid; the Fields pane then
/// takes over. Returns an error message; null = success (the new pivot is
/// selected so the pane can pick it up).
export function handleCreatePivot(
  ctx: PivotActionContext,
  request: PivotCreateRequest,
): string | null {
  const runtime = ctx.univerRef.current
  if (!runtime) return t('appWorkbookNotReady')
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!workbook || !worksheet) return t('appOpenWorkbookFirst')
  const state = ctx.lazyWorkbookRef.current
  if (!state) return t('appOpenXlsxFirst')
  const sheetId = worksheet.getSheetId()
  const fields = pivotFieldOptions(ctx, request.sourceRange)
  const model = defaultPivotModel(fields)
  if (validatePivotModel(model) !== null) return t('dlgPivotErrNeedValue')
  const config = modelToConfig(model, request.sourceRange, request.targetCell)
  const parts = pivotConfigToOpParts(config, fields)
  if (typeof parts === 'string') return parts
  try {
    const result = applyAiPivotAdd(runtime, state, {
      op: 'add_pivot',
      sheetId,
      sourceRange: request.sourceRange,
      targetCell: request.targetCell || 'A1',
      ...parts,
    })
    worksheet.getRange(result.location.startRow, result.location.startColumn, 1, 1).activate()
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.setMessage(t('appPivotCreated', { cell: request.targetCell || 'A1' }))
    return null
  } catch (e) {
    return e instanceof Error ? e.message : t('appPivotCreateFailed')
  }
}

function topLeftOf(area: { startRow: number; startColumn: number }, rowsAbove: number): string {
  return `${columnLabel(area.startColumn)}${Math.max(0, area.startRow - rowsAbove) + 1}`
}

/// Layout model of a journaled pivot (session add or pending relayout).
function modelOfPivotAdd(add: WorkbookPivotAdd): PivotLayoutModel {
  const columns =
    add.columnFieldIndices ?? (add.columnFieldIndex === undefined ? [] : [add.columnFieldIndex])
  const values: PivotValueSpec[] = add.values.map((value, index) => {
    const filter = add.filters?.find((entry) => entry.kind === 'value' && entry.dataField === index)
    return {
      fieldIndex: value.fieldIndex,
      agg: value.agg,
      ...(value.showDataAs !== undefined ? { showDataAs: value.showDataAs } : {}),
      ...(value.name !== undefined ? { name: value.name } : {}),
      ...(value.formula !== undefined ? { formula: value.formula, calcName: value.calcName } : {}),
      ...(filter !== undefined && filter.kind === 'value'
        ? {
            filter: {
              op: filter.op,
              ...(filter.count !== undefined ? { count: filter.count } : {}),
              ...(filter.from !== undefined ? { from: filter.from } : {}),
              ...(filter.to !== undefined ? { to: filter.to } : {}),
            },
          }
        : {}),
    }
  })
  const groupings: Record<number, PivotLayoutModel['groupings'][number]> = {}
  for (const grouping of add.groupings ?? []) {
    groupings[grouping.fieldIndex] =
      grouping.kind === 'date'
        ? { kind: 'date', dateUnit: grouping.dateUnit }
        : { kind: 'range', rangeStep: grouping.rangeStep }
  }
  const labelFilters: Record<number, PivotLayoutModel['labelFilters'][number]> = {}
  for (const filter of add.filters ?? []) {
    if (filter.kind === 'label') labelFilters[filter.field] = { op: filter.op, value: filter.value }
  }
  return {
    filters: (add.pageFieldIndices ?? []).map((fieldIndex, page) => {
      const item = add.pageItems?.[page] ?? null
      return {
        fieldIndex,
        item: item === null ? null : (add.pageLevelItems?.[page]?.[item] ?? null),
      }
    }),
    columns,
    rows: [...add.rowFieldIndices],
    values,
    groupings,
    labelFilters,
  }
}

/// Layout model of a file pivot from its parsed definition. Returns a message
/// key when the definition is outside what the pane can edit.
function modelOfDefinition(
  definition: WorkbookPivotDefinition,
): { model: PivotLayoutModel; sourceFieldCount: number } | string {
  if (definition.unsupported.length > 0) {
    return t('appPivotEditUnsupported', { reasons: definition.unsupported.join('; ') })
  }
  if (definition.rowFields.some((field) => field < 0)) return t('appPivotEditValuesOnRows')
  if (definition.dataFields.some((dataField) => !isPivotAggregation(dataField.subtotal))) {
    return t('appPivotEditAggUnsupported')
  }
  const sourceFieldCount = definition.fields.filter((field) => field.formula === undefined).length
  const memberLabel = (field: number, item: number | null): string | null => {
    if (item === null) return null
    const x = definition.fieldItems[field]?.[item]?.x
    const shared =
      x === null || x === undefined ? undefined : definition.fields[field]?.sharedItems[x]
    return shared === undefined || shared === null ? '' : String(shared)
  }
  const values: PivotValueSpec[] = definition.dataFields.map((dataField, index) => {
    const source = definition.fields[dataField.field]
    const agg = dataField.subtotal as PivotValueSpec['agg']
    const isCalc = source?.formula !== undefined
    const defaultName = defaultDataFieldCaption(agg, source?.name ?? '')
    const filter = definition.filters.find(
      (entry) => entry.kind === 'value' && entry.dataField === index,
    )
    return {
      fieldIndex: isCalc ? -1 : dataField.field,
      agg,
      ...(dataField.showDataAs !== undefined ? { showDataAs: dataField.showDataAs } : {}),
      ...(dataField.name !== defaultName && !isCalc ? { name: dataField.name } : {}),
      ...(isCalc ? { formula: source.formula, calcName: source.name } : {}),
      ...(filter !== undefined && filter.kind === 'value'
        ? {
            filter: {
              op: filter.op,
              ...(filter.count !== undefined ? { count: filter.count } : {}),
              ...(filter.from !== undefined ? { from: filter.from } : {}),
              ...(filter.to !== undefined ? { to: filter.to } : {}),
            },
          }
        : {}),
    }
  })
  // The pane's value filter always targets the level-1 row field.
  if (
    definition.filters.some(
      (filter) => filter.kind === 'value' && filter.field !== definition.rowFields[0],
    )
  ) {
    return t('appPivotEditHasFeatures')
  }
  const groupings: Record<number, PivotLayoutModel['groupings'][number]> = {}
  definition.fields.forEach((field, index) => {
    if (!field.grouping) return
    groupings[index] =
      field.grouping.kind === 'date'
        ? { kind: 'date', dateUnit: field.grouping.dateUnit }
        : { kind: 'range', rangeStep: field.grouping.rangeStep }
  })
  const labelFilters: Record<number, PivotLayoutModel['labelFilters'][number]> = {}
  for (const filter of definition.filters) {
    if (filter.kind === 'label') labelFilters[filter.field] = { op: filter.op, value: filter.value }
  }
  return {
    sourceFieldCount,
    model: {
      filters: definition.pageFields.map((page) => ({
        fieldIndex: page.field,
        item: memberLabel(page.field, page.item),
      })),
      columns: definition.colFields.filter((field) => field >= 0),
      rows: [...definition.rowFields],
      values,
      groupings,
      labelFilters,
    },
  }
}

/// Seed for the Fields pane from the pivot under the cursor; locks the edit
/// context used by every apply. When it returns null, the reason was already
/// shown via setMessage.
export function pivotEditInitial(ctx: PivotActionContext): PivotEditSeed | null {
  const state = ctx.lazyWorkbookRef.current
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  if (!state || !workbook) {
    ctx.setMessage(t('appOpenXlsxFirst'))
    return null
  }
  const hit = findPivotAtSelection(ctx)
  if (!hit) {
    ctx.setMessage(t('appPutCursorInPivot'))
    return null
  }
  const key = hit.kind === 'file' ? hit.pivot.path : `session:${hit.add.name}`
  const add = hit.kind === 'session' ? hit.add : hit.relayout
  let sourceSheetId: string
  let sourceRange: string
  let model: PivotLayoutModel
  let bodyBounds: PivotRelayoutTarget['oldBounds']
  let target: PivotRelayoutTarget['target']
  let fieldNames: string[]
  if (add !== undefined) {
    sourceSheetId = add.sourceSheetId
    sourceRange =
      `${columnLabel(add.sourceArea.startColumn)}${add.sourceArea.startRow + 1}` +
      `:${columnLabel(add.sourceArea.endColumn)}${add.sourceArea.endRow + 1}`
    model = modelOfPivotAdd(add)
    bodyBounds = hit.kind === 'file' ? parseRange(hit.pivot.outputRef) : { ...add.location }
    target =
      hit.kind === 'file'
        ? { pivotPath: hit.pivot.path, cachePath: hit.pivot.cachePath! }
        : { sessionName: add.name }
    fieldNames = [...add.fieldNames]
  } else {
    if (hit.kind !== 'file') return null
    if (hit.pivot.cachePath === null) {
      ctx.setMessage(t('appPivotNoCacheDef'))
      return null
    }
    const definition = hit.definition
    if (!definition) {
      ctx.setMessage(t('appPivotDefNotLoadedSave'))
      return null
    }
    const parsed = modelOfDefinition(definition)
    if (typeof parsed === 'string') {
      ctx.setMessage(parsed)
      return null
    }
    const sourceSheet = state.file.sheets.find((sheet) => sheet.name === definition.sourceSheet)
    if (!sourceSheet || !workbook.getSheetBySheetId(sourceSheet.id)) {
      ctx.setMessage(t('appPivotSourceSheetNotFound', { name: definition.sourceSheet }))
      return null
    }
    sourceSheetId = sourceSheet.id
    sourceRange = definition.sourceRef
    model = parsed.model
    bodyBounds = parseRange(hit.pivot.outputRef)
    target = { pivotPath: hit.pivot.path, cachePath: hit.pivot.cachePath }
    fieldNames = definition.fields.slice(0, parsed.sourceFieldCount).map((field) => field.name)
  }
  const sourceSheet = workbook.getSheetBySheetId(sourceSheetId)
  if (!sourceSheet) {
    ctx.setMessage(t('appPivotSourceSheetNotFound', { name: sourceSheetId }))
    return null
  }
  const sampled = pivotFieldOptions(ctx, sourceRange, sourceSheet)
  const sourceBounds = parseRange(sourceRange)
  const fields: PivotField[] = fieldNames.map((label, index) => ({
    label,
    colIndex: sourceBounds.startColumn + index,
    numeric: sampled[index]?.numeric ?? false,
  }))
  const oldPageRows = Math.min(bodyBounds.startRow, pivotPageRows(model.filters.length))
  const targetCell = topLeftOf(bodyBounds, oldPageRows)
  ctx.pivotEditContextRef.current = {
    key,
    sourceSheetId,
    targetSheetId: hit.sheetId,
    sourceRange,
    targetCell,
    fields,
    relayout: { target, oldBounds: { ...bodyBounds }, oldPageRows },
  }
  return { key, fields, sourceRange, targetCell, model }
}

/// Distinct labels of one source field (report-filter item picker).
export function pivotFieldMembers(ctx: PivotActionContext, fieldIndex: number): string[] {
  const context = ctx.pivotEditContextRef.current
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const sourceSheet = context ? workbook?.getSheetBySheetId(context.sourceSheetId) : undefined
  if (!context || !sourceSheet) return []
  const grid = readPivotSourceGrid(
    sourceSheet.getRange(context.sourceRange),
    ctx.lazyWorkbookRef.current?.file.date1904 === true,
  )
  const members: string[] = []
  for (const row of grid.slice(1)) {
    const label = String(row[fieldIndex] ?? '')
    if (!members.includes(label)) members.push(label)
  }
  return members
}

/// Applies a layout from the Fields pane to the pivot locked in the edit
/// context; the context's bounds move with the result so the next apply
/// clears the right area.
export function handleEditPivotApply(
  ctx: PivotActionContext,
  config: OoXmlPivotConfig,
): string | null {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  const context = ctx.pivotEditContextRef.current
  if (!runtime || !state || !context) return t('appWorkbookNotReady')
  const parts = pivotConfigToOpParts(config, context.fields)
  if (typeof parts === 'string') return parts
  try {
    const result = applyAiPivotAdd(
      runtime,
      state,
      {
        op: 'add_pivot',
        sheetId: context.sourceSheetId,
        ...(context.targetSheetId === context.sourceSheetId
          ? {}
          : { targetSheetId: context.targetSheetId }),
        sourceRange: context.sourceRange,
        targetCell: context.targetCell,
        ...parts,
      },
      context.relayout,
    )
    context.relayout = {
      target: context.relayout.target,
      oldBounds: { ...result.location },
      oldPageRows: result.pageRows,
    }
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.setMessage(t('appPivotLayoutUpdated'))
    return null
  } catch (error) {
    return error instanceof Error ? error.message : t('appPivotEditFailed')
  }
}

export function getSourceRange(ctx: PivotActionContext): string {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!range || !worksheet) return ''
  const source = resolvePivotSource(
    worksheet,
    {
      startRow: range.getRow(),
      startColumn: range.getColumn(),
      endRow: range.getRow() + range.getHeight() - 1,
      endColumn: range.getColumn() + range.getWidth() - 1,
    },
    ctx.lazyWorkbookRef.current,
  )
  const start = `${columnLabel(source.startColumn)}${source.startRow + 1}`
  const end = `${columnLabel(source.endColumn)}${source.endRow + 1}`
  return `${start}:${end}`
}

export function isSelectionInPivot(ctx: PivotActionContext): boolean {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  if (!runtime || !state) return false
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!worksheet || !range) return false
  const sheetId = worksheet.getSheetId()
  const sheetMeta = state.file.sheets.find((sheet) => sheet.id === sheetId)
  if (!sheetMeta) return false
  const selRow = range.getRow()
  const selCol = range.getColumn()
  const selEndRow = selRow + range.getHeight() - 1
  const selEndCol = selCol + range.getWidth() - 1
  return sheetMeta.pivotRanges.some(
    (pivot) =>
      selRow <= pivot.endRow &&
      selEndRow >= pivot.startRow &&
      selCol <= pivot.endColumn &&
      selEndCol >= pivot.startColumn,
  )
}

/// Data › Refresh All: every pivot table on every sheet, one journaled pass.
export function handleRefreshAllPivots(ctx: PivotActionContext): string | null {
  const state = ctx.lazyWorkbookRef.current
  if (!ctx.univerRef.current || !state) return t('appOpenXlsxFirst')
  const pivotSheets = state.file.sheets.filter((sheet) => sheet.pivotTables.length > 0)
  if (pivotSheets.length === 0) return t('appWorkbookNoPivot')
  let count = 0
  try {
    for (const sheet of pivotSheets) {
      count += refreshPivotTables(ctx, sheet.id)
    }
  } catch (e) {
    return e instanceof Error ? e.message : t('appRefreshFailed')
  }
  ctx.setMessage(t('appPivotsRefreshed', { count }))
  return null
}

export function handleRefreshPivot(ctx: PivotActionContext): string | null {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  if (!runtime || !state) return t('appOpenXlsxFirst')
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!worksheet) return t('appNoActiveSheet')
  const sheetId = worksheet.getSheetId()
  const sheetMeta = state.file.sheets.find((sheet) => sheet.id === sheetId)
  if (!sheetMeta || sheetMeta.pivotTables.length === 0) {
    return t('appCurrentSheetNoPivot')
  }
  try {
    const count = refreshPivotTables(ctx, sheetId)
    ctx.setMessage(t('appPivotsRefreshed', { count }))
    return null
  } catch (e) {
    return e instanceof Error ? e.message : t('appRefreshFailed')
  }
}

/// Slicers and timelines both apply through the pivot's per-field hidden-items
/// state, so a field can carry at most one filter panel: a second one on the
/// same field would silently overwrite (and on removal, wholesale clear) the
/// first one's selection. Both pickers exclude fields already taken.
function fieldsWithFilterPanel(ctx: PivotActionContext, pivotPath: string): Set<number> {
  const taken = new Set<number>()
  for (const slicer of ctx.slicers) {
    if (slicer.pivotPath === pivotPath) taken.add(slicer.field)
  }
  for (const timeline of ctx.timelines) {
    if (timeline.pivotPath === pivotPath) taken.add(timeline.field)
  }
  return taken
}

/// Insert slicer: when the cursor hits a pivot, open the field picker
/// (row/column/report-filter dimension fields).
export function handleOpenSlicerPicker(ctx: PivotActionContext): void {
  const state = ctx.lazyWorkbookRef.current
  if (!state) {
    ctx.setMessage(t('appSlicerNeedsFile'))
    return
  }
  const found = filePivotAtSelection(ctx)
  if (!found) return
  const definition = found.definition
  const taken = fieldsWithFilterPanel(ctx, found.pivot.path)
  const seen = new Set<number>()
  const fields: { field: number; name: string }[] = []
  for (const field of [
    ...definition.rowFields,
    ...definition.colFields,
    ...definition.pageFields.map((page) => page.field),
  ]) {
    if (field < 0 || seen.has(field)) continue
    seen.add(field)
    fields.push({
      field,
      name: definition.fields[field]?.name ?? t('appFieldN', { n: field + 1 }),
    })
  }
  if (fields.length === 0) {
    ctx.setMessage(t('appPivotNoSlicerFields'))
    return
  }
  const available = fields.filter((entry) => !taken.has(entry.field))
  if (available.length === 0) {
    ctx.setMessage(t('appFieldFilterTaken'))
    return
  }
  ctx.setSlicerPicker({ sheetId: found.sheetId, pivotPath: found.pivot.path, fields: available })
}

/// Creates the slicer panel after field selection: members come from the pivot
/// cache's sharedItems; initial selection = currently visible (unhidden)
/// members. Returns an error message; null = success.
export function handleCreateSlicer(ctx: PivotActionContext, field: number): string | null {
  const state = ctx.lazyWorkbookRef.current
  const picker = ctx.slicerPicker
  if (!state || !picker) return t('appSlicerPivotStale')
  const definition = state.pivotDefinitions.get(picker.pivotPath)
  if (!definition) return t('appPivotDefNotLoaded')
  const members: SlicerMember[] = []
  const selected: number[] = []
  definition.fieldItems[field]?.forEach((item, member) => {
    if (item.x === null) return
    const shared = definition.fields[field]?.sharedItems[item.x]
    members.push({
      member,
      label:
        shared === null || shared === undefined || shared === '' ? t('appBlank') : String(shared),
    })
    if (!item.hidden) selected.push(member)
  })
  if (members.length === 0) return t('appFieldNoMembers')
  const slicer: SlicerUiState = {
    id: `slicer-${Date.now().toString(36)}-${ctx.slicers.length + 1}`,
    sheetId: picker.sheetId,
    pivotPath: picker.pivotPath,
    field,
    fieldName: definition.fields[field]?.name ?? t('appFieldN', { n: field + 1 }),
    members,
    selected,
  }
  ctx.setSlicers((current) => [...current, slicer])
  ctx.setMessage(t('appSlicerCreated', { name: slicer.fieldName }))
  return null
}

/// Applies the slicer's selection to the pivot: unselected members become
/// hidden entries, the layout is rebuilt and recomputed, and the output area is
/// rewritten wholesale via the refresh "growth" path (the layout may shrink or
/// recover). Returns an error message; null = success (incl. no-change no-ops).
export function applySlicerSelection(
  ctx: PivotActionContext,
  // Both slicers and timelines apply through here (same hidden-items model).
  slicer: { readonly sheetId: string; readonly pivotPath: string; readonly field: number },
  selectedMembers: readonly number[] | null,
): string | null {
  const state = ctx.lazyWorkbookRef.current
  const runtime = ctx.univerRef.current
  if (!state || !runtime) return t('appOpenXlsxFirst')
  if (!state.formulaMode || !state.flags.preloadComplete) {
    return t('appSlicerNeedsFullLoad')
  }
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const sheetMeta = state.file.sheets.find((sheet) => sheet.id === slicer.sheetId)
  const pivot = sheetMeta?.pivotTables.find((entry) => entry.path === slicer.pivotPath)
  if (!workbook || !sheetMeta || !pivot) return t('appSlicerPivotMissing')
  const target = workbook.getSheetBySheetId(slicer.sheetId)
  if (!target) return t('appSlicerSheetMissing')
  const definition = state.pivotDefinitions.get(pivot.path)
  if (!definition) return t('appPivotDefNotLoaded')
  const sourceSheet = workbook
    .getSheets()
    .find((sheet) => sheet.getSheetName() === definition.sourceSheet)
  if (!sourceSheet) return t('appPivotSourceSheetMissing', { name: definition.sourceSheet })
  try {
    const sourceValues = readPivotSourceGrid(
      sourceSheet.getRange(definition.sourceRef),
      ctx.lazyWorkbookRef.current?.file.date1904 === true,
      pivotSourceValueFields(definition),
    )
    const next = applyPivotSlicer(definition, sourceValues, slicer.field, selectedMembers)
    if (next === definition) return null
    const { data } = recomputePivotData(next, sourceValues)
    const bounds = parseRange(definition.outputRef)
    const newOutputRef = applyGrownPivotOutput(target, next, data, bounds)
    if (pivot.cachePath !== null) {
      // On save, the location ref is widened/shrunk to the new area; OOXML
      // write-back of the hidden entries themselves relies on refreshOnLoad
      // letting Excel rebuild (same strategy as the refresh growth path).
      recordPivotRefreshUpdate(state.editJournal, pivot.cachePath, slicer.sheetId, newOutputRef)
      recordPivotCacheRefresh(state.editJournal, pivot.cachePath)
    }
    state.pivotDefinitions.set(pivot.path, {
      ...next,
      outputRef: newOutputRef,
    } as WorkbookPivotDefinition)
    const newBounds = parseRange(newOutputRef)
    const staleRange = sheetMeta.pivotRanges.find(
      (range) =>
        range.startRow === bounds.startRow &&
        range.startColumn === bounds.startColumn &&
        range.endRow === bounds.endRow &&
        range.endColumn === bounds.endColumn,
    )
    if (staleRange) {
      staleRange.endRow = newBounds.endRow
      staleRange.endColumn = newBounds.endColumn
    }
    pivot.outputRef = newOutputRef
    ctx.setPendingEdits(journalSize(state.editJournal))
    return null
  } catch (error) {
    return error instanceof Error ? error.message : t('appSlicerFilterFailed')
  }
}

export function handleSlicerToggle(
  ctx: PivotActionContext,
  slicerId: string,
  member: number,
): void {
  const slicer = ctx.slicers.find((entry) => entry.id === slicerId)
  if (!slicer) return
  const next = slicer.selected.includes(member)
    ? slicer.selected.filter((entry) => entry !== member)
    : [...slicer.selected, member].sort((a, b) => a - b)
  if (next.length === 0) {
    ctx.setMessage(t('appSlicerKeepOne'))
    return
  }
  const failure = applySlicerSelection(ctx, slicer, next)
  if (failure !== null) {
    ctx.setMessage(failure)
    return
  }
  ctx.setSlicers((current) =>
    current.map((entry) => (entry.id === slicerId ? { ...entry, selected: next } : entry)),
  )
  ctx.setMessage(t('appSlicerApplied', { name: slicer.fieldName }))
}

export function handleSlicerSelectAll(ctx: PivotActionContext, slicerId: string): void {
  const slicer = ctx.slicers.find((entry) => entry.id === slicerId)
  if (!slicer) return
  const failure = applySlicerSelection(ctx, slicer, null)
  if (failure !== null) {
    ctx.setMessage(failure)
    return
  }
  ctx.setSlicers((current) =>
    current.map((entry) =>
      entry.id === slicerId
        ? { ...entry, selected: entry.members.map((item) => item.member) }
        : entry,
    ),
  )
  ctx.setMessage(t('appSlicerCleared', { name: slicer.fieldName }))
}

export function handleRemoveSlicer(ctx: PivotActionContext, slicerId: string): void {
  const slicer = ctx.slicers.find((entry) => entry.id === slicerId)
  if (!slicer) return
  const failure = applySlicerSelection(ctx, slicer, null)
  if (failure !== null) {
    ctx.setMessage(failure)
    return
  }
  ctx.setSlicers((current) => current.filter((entry) => entry.id !== slicerId))
  ctx.setMessage(t('appSlicerRemoved', { name: slicer.fieldName }))
}

/// Insert timeline: like the slicer picker, but only ungrouped fields whose
/// non-blank cache items all parse as dates qualify.
export function handleOpenTimelinePicker(ctx: PivotActionContext): void {
  const state = ctx.lazyWorkbookRef.current
  if (!state) {
    ctx.setMessage(t('appSlicerNeedsFile'))
    return
  }
  const found = filePivotAtSelection(ctx)
  if (!found) return
  const definition = found.definition
  const taken = fieldsWithFilterPanel(ctx, found.pivot.path)
  const seen = new Set<number>()
  const fields: { field: number; name: string }[] = []
  for (const field of [
    ...definition.rowFields,
    ...definition.colFields,
    ...definition.pageFields.map((page) => page.field),
  ]) {
    if (field < 0 || seen.has(field)) continue
    seen.add(field)
    const meta = definition.fields[field]
    if (!meta || meta.grouping !== undefined) continue
    if (timelineDomainOf(meta.sharedItems, definition.fieldItems[field] ?? []) === null) continue
    fields.push({ field, name: meta.name || t('appFieldN', { n: field + 1 }) })
  }
  if (fields.length === 0) {
    ctx.setMessage(t('appTimelineNoDateFields'))
    return
  }
  const available = fields.filter((entry) => !taken.has(entry.field))
  if (available.length === 0) {
    ctx.setMessage(t('appFieldFilterTaken'))
    return
  }
  ctx.setTimelinePicker({ sheetId: found.sheetId, pivotPath: found.pivot.path, fields: available })
}

/// Creates the timeline panel after field selection. Returns an error message;
/// null = success.
export function handleCreateTimeline(ctx: PivotActionContext, field: number): string | null {
  const state = ctx.lazyWorkbookRef.current
  const picker = ctx.timelinePicker
  if (!state || !picker) return t('appSlicerPivotStale')
  const definition = state.pivotDefinitions.get(picker.pivotPath)
  if (!definition) return t('appPivotDefNotLoaded')
  const meta = definition.fields[field]
  const domain = timelineDomainOf(meta?.sharedItems ?? [], definition.fieldItems[field] ?? [])
  if (!domain) return t('appTimelineNoDateFields')
  const timeline: TimelineUiState = {
    id: `timeline-${Date.now().toString(36)}-${ctx.timelines.length + 1}`,
    sheetId: picker.sheetId,
    pivotPath: picker.pivotPath,
    field,
    fieldName: meta?.name || t('appFieldN', { n: field + 1 }),
    members: domain.members,
    minMonth: domain.minMonth,
    maxMonth: domain.maxMonth,
    range: null,
  }
  ctx.setTimelines((current) => [...current, timeline])
  ctx.setMessage(t('appTimelineCreated', { name: timeline.fieldName }))
  return null
}

/// Applies a month range (null = clear) through the slicer mechanism.
export function handleTimelineRange(
  ctx: PivotActionContext,
  timelineId: string,
  range: { readonly start: MonthKey; readonly end: MonthKey } | null,
): void {
  const timeline = ctx.timelines.find((entry) => entry.id === timelineId)
  if (!timeline) return
  const selected = timelineSelection(
    { members: timeline.members, minMonth: timeline.minMonth, maxMonth: timeline.maxMonth },
    range,
  )
  if (selected !== null && selected.length === 0) {
    ctx.setMessage(t('appTimelineEmptyRange'))
    return
  }
  const failure = applySlicerSelection(ctx, timeline, selected)
  if (failure !== null) {
    ctx.setMessage(failure)
    return
  }
  ctx.setTimelines((current) =>
    current.map((entry) => (entry.id === timelineId ? { ...entry, range } : entry)),
  )
  ctx.setMessage(
    range === null
      ? t('appTimelineCleared', { name: timeline.fieldName })
      : t('appTimelineApplied', { name: timeline.fieldName }),
  )
}

export function handleRemoveTimeline(ctx: PivotActionContext, timelineId: string): void {
  const timeline = ctx.timelines.find((entry) => entry.id === timelineId)
  if (!timeline) return
  const failure = applySlicerSelection(ctx, timeline, null)
  if (failure !== null) {
    ctx.setMessage(failure)
    return
  }
  ctx.setTimelines((current) => current.filter((entry) => entry.id !== timelineId))
  ctx.setMessage(t('appTimelineRemoved', { name: timeline.fieldName }))
}
