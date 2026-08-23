/**
 * Pivot-table and slicer actions (create/edit/refresh dialogs, slicer panels).
 * Extracted from App.tsx; the App component passes a PivotActionContext built
 * fresh per call so refs and state never go stale.
 */
import { columnLabel, parseRange } from '../domain/cell-address'
import { applyPivotSlicer, growPivotDefinition, recomputePivotData } from '../domain/pivot-engine'
import { timelineDomainOf, timelineSelection, type MonthKey } from '../domain/pivot-timeline'
import type { WorkbookFile, WorkbookPivotDefinition } from '../shared/desktop-api'
import { journalSize, recordPivotCacheRefresh, recordPivotRefreshUpdate } from './edit-journal'
import { t } from './i18n/locale'
import type { OoXmlPivotConfig, PivotEditSeed, PivotField } from './PivotDialog'
import type { SlicerMember, SlicerUiState } from './SlicerPanel'
import type { TimelineUiState } from './TimelinePanel'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'
import { applyAiPivotAdd, applyGrownPivotOutput, pivotConfigToOpParts } from './workbook-ops'

/// A3 editing of an existing pivot: context locked when the dialog opens, used
/// on Apply.
export interface PivotEditContext {
  sourceSheetId: string
  targetSheetId: string
  sourceRange: string
  pivotPath: string
  cachePath: string
  oldBounds: { startRow: number; startColumn: number; endRow: number; endColumn: number }
  fields: PivotField[]
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

export function findPivotAtSelection(ctx: PivotActionContext): {
  sheetId: string
  pivot: WorkbookFile['sheets'][number]['pivotTables'][number]
  definition: WorkbookPivotDefinition | undefined
} | null {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  if (!runtime || !state) return null
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!worksheet || !range) return null
  const sheetId = worksheet.getSheetId()
  const sheetMeta = state.file.sheets.find((sheet) => sheet.id === sheetId)
  if (!sheetMeta) return null
  const selRow = range.getRow()
  const selCol = range.getColumn()
  const selEndRow = selRow + range.getHeight() - 1
  const selEndCol = selCol + range.getWidth() - 1
  for (const pivot of sheetMeta.pivotTables) {
    const bounds = parseRange(pivot.outputRef)
    if (
      selRow <= bounds.endRow &&
      selEndRow >= bounds.startRow &&
      selCol <= bounds.endColumn &&
      selEndCol >= bounds.startColumn
    ) {
      return { sheetId, pivot, definition: state.pivotDefinitions.get(pivot.path) }
    }
  }
  return null
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
    const sourceValues = sourceSheet.getRange(definition.sourceRef).getValues() as (
      string | number | boolean | null | undefined
    )[][]
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

export function pivotFieldOptions(ctx: PivotActionContext): { label: string; colIndex: number }[] {
  const range = ctx.univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveRange()
  if (!range || range.getHeight() < 2) return []
  const headerRow = range.getValues()[0] ?? []
  const start = range.getColumn()
  return headerRow.slice(0, 26).map((header, offset) => ({
    label:
      header === null || header === undefined || header === ''
        ? t('appColumnLabel', { col: columnLabel(start + offset) })
        : String(header),
    colIndex: start + offset,
  }))
}

export function handleCreatePivot(
  ctx: PivotActionContext,
  config: OoXmlPivotConfig,
): string | null {
  const runtime = ctx.univerRef.current
  if (!runtime) return t('appWorkbookNotReady')
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!workbook || !worksheet) return t('appOpenWorkbookFirst')
  const state = ctx.lazyWorkbookRef.current
  if (!state) return t('appOpenXlsxFirst')
  const sheetId = worksheet.getSheetId()
  const parts = pivotConfigToOpParts(config, pivotFieldOptions(ctx))
  if (typeof parts === 'string') return parts
  try {
    applyAiPivotAdd(runtime, state, {
      op: 'add_pivot',
      sheetId,
      sourceRange: config.sourceRange,
      targetCell: config.targetCell || 'A1',
      ...parts,
    })
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.setMessage(t('appPivotCreated', { cell: config.targetCell || 'A1' }))
    return null
  } catch (e) {
    return e instanceof Error ? e.message : t('appPivotCreateFailed')
  }
}

/// A3 editing of an existing pivot: pivot under the cursor → dialog prefill
/// context. When it returns null, the reason was already shown via setMessage.
/// The context is also stored in a ref for Apply.
export function pivotEditInitial(ctx: PivotActionContext): PivotEditSeed | null {
  const state = ctx.lazyWorkbookRef.current
  if (!state) {
    ctx.setMessage(t('appOpenXlsxFirst'))
    return null
  }
  const found = findPivotAtSelection(ctx)
  if (!found) {
    ctx.setMessage(t('appPutCursorInPivot'))
    return null
  }
  const { sheetId, pivot, definition } = found
  if (pivot.cachePath === null) {
    ctx.setMessage(t('appPivotNoCacheDef'))
    return null
  }
  if (!definition) {
    ctx.setMessage(t('appPivotDefNotLoadedSave'))
    return null
  }
  if (definition.unsupported.length > 0) {
    ctx.setMessage(t('appPivotEditUnsupported', { reasons: definition.unsupported.join('; ') }))
    return null
  }
  // MVP boundary: existing grouping/filters/report filters/calculated fields
  // cannot be prefilled into the dialog, and a layout change would drop them —
  // fail closed.
  if (
    definition.fields.some(
      (field) => field.grouping !== undefined || field.formula !== undefined,
    ) ||
    definition.filters.length > 0 ||
    definition.pageFields.length > 0
  ) {
    ctx.setMessage(t('appPivotEditHasFeatures'))
    return null
  }
  if (definition.rowFields.some((field) => field < 0)) {
    ctx.setMessage(t('appPivotEditValuesOnRows'))
    return null
  }
  const supportedAggs = new Set(['sum', 'count', 'average', 'max', 'min'])
  if (definition.dataFields.some((dataField) => !supportedAggs.has(dataField.subtotal))) {
    ctx.setMessage(t('appPivotEditAggUnsupported'))
    return null
  }
  const sourceSheet = state.file.sheets.find((sheet) => sheet.name === definition.sourceSheet)
  if (
    !sourceSheet ||
    !ctx.univerRef.current?.univerAPI.getActiveWorkbook()?.getSheetBySheetId(sourceSheet.id)
  ) {
    ctx.setMessage(t('appPivotSourceSheetNotFound', { name: definition.sourceSheet }))
    return null
  }
  const sourceBounds = parseRange(definition.sourceRef)
  const bounds = parseRange(pivot.outputRef)
  // No calculated fields (guaranteed above), so cache-field indices map
  // one-to-one to dialog field indices.
  const fields: PivotField[] = definition.fields.map((field, index) => ({
    label: field.name,
    colIndex: sourceBounds.startColumn + index,
  }))
  ctx.pivotEditContextRef.current = {
    sourceSheetId: sourceSheet.id,
    targetSheetId: sheetId,
    sourceRange: definition.sourceRef,
    pivotPath: pivot.path,
    cachePath: pivot.cachePath,
    oldBounds: {
      startRow: bounds.startRow,
      startColumn: bounds.startColumn,
      endRow: bounds.endRow,
      endColumn: bounds.endColumn,
    },
    fields,
  }
  return {
    fields,
    sourceRange: definition.sourceRef,
    initial: {
      rowFieldIndices: definition.rowFields.filter((field) => field >= 0),
      colFieldIndices: definition.colFields.filter((field) => field >= 0),
      values: definition.dataFields.map((dataField) => ({
        fieldIndex: dataField.field,
        agg: dataField.subtotal as 'sum' | 'count' | 'average' | 'max' | 'min',
        ...(dataField.showDataAs !== undefined ? { showDataAs: dataField.showDataAs } : {}),
      })),
      targetCell: `${columnLabel(bounds.startColumn)}${bounds.startRow + 1}`,
    },
  }
}

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
    applyAiPivotAdd(
      runtime,
      state,
      {
        op: 'add_pivot',
        sheetId: context.sourceSheetId,
        ...(context.targetSheetId === context.sourceSheetId
          ? {}
          : { targetSheetId: context.targetSheetId }),
        sourceRange: context.sourceRange,
        targetCell: config.targetCell,
        ...parts,
      },
      {
        pivotPath: context.pivotPath,
        cachePath: context.cachePath,
        oldBounds: context.oldBounds,
      },
    )
    ctx.setPendingEdits(journalSize(state.editJournal))
    ctx.setMessage(t('appPivotLayoutUpdated'))
    return null
  } catch (error) {
    return error instanceof Error ? error.message : t('appPivotEditFailed')
  }
}

export function getSourceRange(ctx: PivotActionContext): string {
  const range = ctx.univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveRange()
  if (!range) return ''
  const start = `${columnLabel(range.getColumn())}${range.getRow() + 1}`
  const end = `${columnLabel(range.getColumn() + range.getWidth() - 1)}${range.getRow() + range.getHeight() - 1 + 1}`
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
  const found = findPivotAtSelection(ctx)
  if (!found) {
    ctx.setMessage(t('appCursorNotInPivot'))
    return
  }
  if (!found.definition) {
    ctx.setMessage(t('appPivotDefNotLoaded'))
    return
  }
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
    const sourceValues = sourceSheet.getRange(definition.sourceRef).getValues() as (
      string | number | boolean | null | undefined
    )[][]
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
  const found = findPivotAtSelection(ctx)
  if (!found) {
    ctx.setMessage(t('appCursorNotInPivot'))
    return
  }
  if (!found.definition) {
    ctx.setMessage(t('appPivotDefNotLoaded'))
    return
  }
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
