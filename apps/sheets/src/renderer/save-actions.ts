/**
 * The ⌘S save flow: serializes the edit journal and Univer-owned states
 * (filters, CF, DV, notes, defined names) into one saveWorkbookEdits call,
 * with the two-phase split when structural changes entangle additions.
 * Extracted from App.tsx; the App component passes a SaveContext built fresh
 * per call so refs and state never go stale.
 */
import type { WorkbookFile, WorkbookFilterState } from '../shared/desktop-api'
import {
  isSheetRemoved,
  toSaveChartEdits,
  toSaveEdits,
  toSaveHyperlinkEdits,
  toSavePageSetupStates,
  toSavePivotAdds,
  toSaveSheetOps,
  toSaveSparklineAdds,
  toSaveStructuralOps,
  toSaveTableAdds,
  toSaveVisualAdds,
  toSaveVisualEdits,
} from './edit-journal'
import { t } from './i18n/locale'
import { showToast } from './toast-bus'
import {
  collectCfStates,
  collectDefinedNamesState,
  collectDvStates,
  collectFilterStates,
  collectNoteStates,
} from './univer-sync'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/** The App refs/state the save flow needs; built fresh per call. */
export interface SaveContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  setMessage: (message: string) => void
  openLazyWorkbook: (opened: WorkbookFile) => void
}

/**
 * mode 'recovery': assemble the very same payload but hand it to the
 * crash-recovery writer instead of the save pipeline — no dialogs, no status
 * messages, no session swap, the opened file untouched.
 */
export async function handleSave(
  ctx: SaveContext,
  mode: 'save' | 'save-as' | 'recovery',
  quiet = false,
): Promise<void> {
  const state = ctx.lazyWorkbookRef.current
  if (!state) {
    if (mode !== 'recovery') ctx.setMessage(t('appDemoNoSave'))
    return
  }
  const edits = toSaveEdits(state.editJournal)
  const structuralOps = toSaveStructuralOps(state.editJournal)
  const chartEdits = toSaveChartEdits(state.editJournal)
  const visualEdits = toSaveVisualEdits(state.editJournal)
  const visualAdditions = toSaveVisualAdds(state.editJournal)
  const tableAdditions = toSaveTableAdds(state.editJournal)
  const pivotAdditions = toSavePivotAdds(state.editJournal)
  const sheetOps = toSaveSheetOps(state.editJournal)
  const hyperlinkEdits = toSaveHyperlinkEdits(state.editJournal)
  let filterStates: WorkbookFilterState[]
  try {
    filterStates = collectFilterStates(ctx.univerRef.current, state)
  } catch (error: unknown) {
    const failed = error instanceof Error ? error.message : t('appFilterSnapshotFailed')
    ctx.setMessage(failed)
    if (mode !== 'recovery' && !quiet) showToast(failed, 'error')
    return
  }
  const cfStates = collectCfStates(ctx.univerRef.current, state)
  const dvStates = collectDvStates(ctx.univerRef.current, state)
  const sheetProtections = [...state.editJournal.sheetProtection]
    .filter(([sheetId]) => !isSheetRemoved(state.editJournal, sheetId))
    .map(([sheetId, isProtected]) => ({ sheetId, protected: isProtected }))
  const pageSetupStates = toSavePageSetupStates(state.editJournal)
  const noteStates = collectNoteStates(ctx.univerRef.current, state)
  const pivotCacheRefreshPaths = [...state.editJournal.pivotCacheRefresh]
  // Output-area expansion after layout growth (location ref write-back); the
  // count folds into cacheRefresh.
  const pivotRefreshUpdates = [...state.editJournal.pivotRefreshUpdates].map(
    ([cachePath, update]) => ({
      cachePath,
      sheetId: update.sheetId,
      newOutputRef: update.newOutputRef,
      ...(update.relayout === undefined ? {} : { relayout: update.relayout }),
    }),
  )
  const definedNamesState = collectDefinedNamesState(ctx.univerRef.current, state)
  const workbookProtectionState =
    state.editJournal.workbookProtection.desired === null
      ? null
      : { lockStructure: state.editJournal.workbookProtection.desired }
  const protectedRangeStates = [...state.editJournal.protectedRangesDirty]
    .filter((sheetId) => !isSheetRemoved(state.editJournal, sheetId))
    .map((sheetId) => ({
      sheetId,
      ranges: (state.sheetProtectedRanges.get(sheetId) ?? []).map(({ name, sqref }) => ({
        name,
        sqref,
      })),
    }))
  const themeState =
    state.editJournal.theme.colors === undefined && state.editJournal.theme.fonts === undefined
      ? null
      : {
          ...(state.editJournal.theme.colors === undefined
            ? {}
            : { colors: state.editJournal.theme.colors }),
          ...(state.editJournal.theme.fonts === undefined
            ? {}
            : { fonts: state.editJournal.theme.fonts }),
        }
  // Recalculated formula results: the engine's values are on screen but
  // deliberately kept out of the journal (they must not become literals). Send them
  // separately so the save refreshes each formula cell's cached <v>, keeping its <f>.
  const formulaValues = [...(state.recalc?.overlay ?? [])].flatMap(([sheetId, cells]) =>
    isSheetRemoved(state.editJournal, sheetId)
      ? []
      : [...cells].flatMap(([key, cell]) => {
          if (cell.v === undefined) return []
          const [row, column] = key.split(':').map(Number)
          if (row === undefined || column === undefined) return []
          return [{ sheetId, row, column, value: cell.v }]
        }),
  )
  // The gateway fails closed when these additions ride with structural or
  // sheet changes (their coordinates entangle). Instead of bouncing the
  // user, hold them back and save in two sequential phases: structure
  // first, then the additions against the reopened session. Pre-existing
  // sheet ids are stable across saves (`sheet-<sheetId attr>`), so the
  // held ops stay addressable — ops on sheets created this session are
  // the exception and keep the explicit error.
  const hasShifts = structuralOps.length > 0 || sheetOps.length > 0
  const heldPivots = hasShifts ? pivotAdditions : []
  const heldTables = structuralOps.length > 0 ? tableAdditions : []
  const heldNames = hasShifts ? definedNamesState : null
  const splitSave = heldPivots.length > 0 || heldTables.length > 0 || heldNames !== null
  if (splitSave) {
    const addedSheetIds = state.editJournal.sheets.added
    const strandedHeld = [
      ...heldPivots.flatMap((pivot) => [pivot.sheetId, pivot.sourceSheetId]),
      ...heldTables.map((table) => table.sheetId),
    ].some((sheetId) => addedSheetIds.has(sheetId))
    if (strandedHeld) {
      if (mode !== 'recovery') {
        ctx.setMessage(t('appSaveHeldStranded'))
        if (!quiet) showToast(t('appSaveHeldStranded'), 'error')
      }
      return
    }
  }
  const total =
    edits.length +
    structuralOps.length +
    chartEdits.length +
    sheetOps.length +
    filterStates.length +
    hyperlinkEdits.length +
    cfStates.length +
    dvStates.length +
    pageSetupStates.length +
    noteStates.length +
    pivotCacheRefreshPaths.length +
    sheetProtections.length +
    (definedNamesState === null ? 0 : 1) +
    (workbookProtectionState === null ? 0 : 1) +
    (themeState === null ? 0 : 1) +
    protectedRangeStates.length +
    visualAdditions.length +
    visualEdits.length +
    tableAdditions.length +
    pivotAdditions.length
  // A restored crash-recovery session carries its changes in the workbook
  // bytes themselves, not the journal: a plain Save with nothing pending must
  // still write back to the original file (and clear the recovery copy).
  const restoreWriteBack = mode === 'save' && state.file.restoredFromRecovery === true
  if (total === 0 && mode !== 'save-as' && !restoreWriteBack) {
    if (mode !== 'recovery') ctx.setMessage(t('appNoEditsToSave'))
    return
  }
  // Sheet ops rebuild workbook.xml's tab list, so the save needs the final
  // on-screen order.
  const sheetOrder =
    sheetOps.length === 0
      ? []
      : (ctx.univerRef.current?.univerAPI
          .getActiveWorkbook()
          ?.getSheets()
          .map((sheet) => sheet.getSheetId()) ?? [])
  if (sheetOps.length > 0 && sheetOrder.length === 0) {
    if (mode !== 'recovery') {
      ctx.setMessage(t('appSheetOrderReadFailed'))
      if (!quiet) showToast(t('appSheetOrderReadFailed'), 'error')
    }
    return
  }
  const payload = {
    sessionId: state.file.sessionId,
    mode: mode === 'recovery' ? ('save' as const) : mode,
    edits,
    structuralOps,
    chartEdits,
    visualEdits,
    visualAdditions,
    tableAdditions,
    pivotAdditions,
    sheetOps,
    sheetOrder,
    filterStates,
    hyperlinkEdits,
    cfStates,
    dvStates,
    pageSetupStates,
    noteStates,
    pivotCacheRefreshPaths,
    pivotRefreshUpdates,
    sheetProtections,
    sparklineAdditions: toSaveSparklineAdds(state.editJournal),
    formulaValues,
    definedNamesState,
    themeState,
    workbookProtectionState,
    protectedRangeStates,
  }
  if (mode === 'recovery') {
    // Best-effort; a failure only means this tick's copy is skipped
    await window.desktopApi.writeWorkbookRecovery(payload).catch(() => ({ ok: false }))
    return
  }
  try {
    ctx.setMessage(t('appSavingEdits', { count: total }))
    const result = await window.desktopApi.saveWorkbookEdits({
      sessionId: state.file.sessionId,
      mode,
      ...(restoreWriteBack ? { restoreWriteBack: true } : {}),
      edits,
      structuralOps,
      chartEdits,
      visualEdits,
      visualAdditions,
      tableAdditions: splitSave && heldTables.length > 0 ? [] : tableAdditions,
      pivotAdditions: splitSave && heldPivots.length > 0 ? [] : pivotAdditions,
      sheetOps,
      sheetOrder,
      filterStates,
      hyperlinkEdits,
      cfStates,
      dvStates,
      pageSetupStates,
      noteStates,
      pivotCacheRefreshPaths,
      pivotRefreshUpdates,
      sheetProtections,
      sparklineAdditions: toSaveSparklineAdds(state.editJournal),
      formulaValues,
      definedNamesState: splitSave ? null : definedNamesState,
      themeState,
      workbookProtectionState,
      protectedRangeStates,
    })
    if (ctx.lazyWorkbookRef.current !== state) return
    if (result.canceled) {
      ctx.setMessage(t('appSaveCanceled'))
      return
    }
    if (!splitSave) {
      ctx.openLazyWorkbook(result.file)
      const saved = t('appSaved', {
        touched: result.touchedEntries.length,
        total: result.file.entryCount,
      })
      ctx.setMessage(saved)
      if (!quiet) showToast(saved)
      return
    }
    try {
      const second = await window.desktopApi.saveWorkbookEdits({
        sessionId: result.file.sessionId,
        mode: 'save',
        edits: [],
        structuralOps: [],
        chartEdits: [],
        visualEdits: [],
        visualAdditions: [],
        tableAdditions: heldTables,
        pivotAdditions: heldPivots,
        sheetOps: [],
        sheetOrder: [],
        filterStates: [],
        hyperlinkEdits: [],
        cfStates: [],
        dvStates: [],
        pageSetupStates: [],
        noteStates: [],
        pivotCacheRefreshPaths: [],
        pivotRefreshUpdates: [],
        sheetProtections: [],
        sparklineAdditions: [],
        // The first phase already refreshed the cached values
        formulaValues: [],
        definedNamesState: heldNames,
        themeState: null,
        workbookProtectionState: null,
        protectedRangeStates: [],
      })
      if (ctx.lazyWorkbookRef.current !== state) return
      if (second.canceled) {
        ctx.openLazyWorkbook(result.file)
        ctx.setMessage(t('appSaveSecondCanceled'))
        return
      }
      ctx.openLazyWorkbook(second.file)
      const saved = t('appSavedTwoPhase')
      ctx.setMessage(saved)
      if (!quiet) showToast(saved)
    } catch (error: unknown) {
      if (ctx.lazyWorkbookRef.current !== state) return
      ctx.openLazyWorkbook(result.file)
      const failed = t('appSaveSecondFailed', {
        reason: error instanceof Error ? error.message : String(error),
      })
      ctx.setMessage(failed)
      if (!quiet) showToast(failed, 'error')
    }
  } catch (error: unknown) {
    const message = stripIpcErrorWrapper(error instanceof Error ? error.message : '')
    const failed = localizeSaveError(message) ?? (message || t('appSaveFailed'))
    ctx.setMessage(failed)
    if (!quiet) showToast(failed, 'error')
  }
}

/// Errors thrown by main-process handlers arrive wrapped as
/// "Error invoking remote method 'workbook:save': Error: <message>" —
/// only <message> is meant for the user.
export function stripIpcErrorWrapper(message: string): string {
  return message.replace(/^Error invoking remote method '[^']+': (?:Error: )?/, '')
}

/// Gateway save errors users can trigger through normal editing, keyed by a
/// stable substring; unmatched messages surface verbatim (English) as before.
const SAVE_ERROR_PATTERNS = [
  ['cannot shift here', 'appStructuralShiftBlocked'],
  ['extended (x14) data validation', 'appSaveErrX14Dv'],
  ['Multi-select list rules', 'appSaveErrMultiSelectList'],
  ['extended conditional formatting (x14)', 'appSaveErrX14Cf'],
  ['data-bar extension format (x14)', 'appSaveErrX14Cf'],
  ['A new pivot cannot be saved together with sheet management', 'appSaveErrPivotWithSheetOps'],
  ['A new pivot cannot be saved together with row/column', 'appSaveErrPivotWithRowCol'],
  ['A new table cannot be saved together with row/column', 'appSaveErrTableWithRowCol'],
  ['Defined-name edits cannot be saved together', 'appSaveErrNamesWithStructural'],
  // Only the mid-save race from the gateway: the pre-save disk-changed error
  // arrives from the main process already localized (and advises Save As,
  // which stays usable), so it must pass through untouched.
  ['The workbook changed on disk while saving', 'appSaveErrChangedOnDisk'],
  ['style edits cannot be saved', 'appSaveErrStylesheetLimited'],
  ['Saving would change the workbook package structure', 'appSaveErrPackageGuard'],
  ['charts support', 'appSaveErrChartUnsupported'],
  ['Converting a', 'appSaveErrChartUnsupported'],
  ['Replacing series on a', 'appSaveErrChartUnsupported'],
  ['overlaps the moved rows', 'appSaveErrMoveOverlap'],
  ['Moving the header row of table', 'appSaveErrMoveOverlap'],
  ['Moving the totals row of table', 'appSaveErrMoveOverlap'],
] as const

export function localizeSaveError(message: string): string | null {
  const hit = SAVE_ERROR_PATTERNS.find(([pattern]) => message.includes(pattern))
  return hit ? t(hit[1]) : null
}
