/**
 * Formula view — Excel's Show Formulas (⌘` / Ctrl+`), persisted per sheet as
 * sheetView/@showFormulas.
 *
 * Univer's RENDER_RAW_FORMULA_KEY only reaches the rich-text/rotated cell
 * path in this version, so plain cells never change: the actual value swap
 * happens in a CELL_CONTENT interceptor below. The context key is still
 * toggled for its side effect — every sheet skeleton resets its cache and
 * repaints when it flips.
 */
import {
  CellValueType,
  DOCS_FORMULA_BAR_EDITOR_UNIT_ID_KEY,
  ICommandService,
  IContextService,
  IUniverInstanceService,
  UniverInstanceType,
  type DocumentDataModel,
} from '@univerjs/core'
import { CoverContentCommand } from '@univerjs/docs-ui'
import { RENDER_RAW_FORMULA_KEY } from '@univerjs/preset-sheets-core'
import { BEFORE_CELL_EDIT, INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'
import { IEditorBridgeService } from '@univerjs/sheets-ui'

import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/// Formula-view sheets for blank/demo workbooks (no LazyWorkbookState).
const demoFormulaViewSheets = new Set<string>()

export function formulaViewSheets(state: LazyWorkbookState | null): Set<string> {
  return state ? state.showFormulaSheets : demoFormulaViewSheets
}

/// Sync the (global) raw-formula render key to the sheet's per-sheet flag;
/// flipping it makes every skeleton drop its text cache and repaint.
export function applyShowFormulasView(
  runtime: UniverRuntime,
  state: LazyWorkbookState | null,
  sheetId: string,
): void {
  const contextService = runtime.univer.__getInjector().get(IContextService)
  const next = formulaViewSheets(state).has(sheetId)
  if (Boolean(contextService.getContextValue(RENDER_RAW_FORMULA_KEY)) !== next) {
    contextService.setContextValue(RENDER_RAW_FORMULA_KEY, next)
  }
}

/// Formula text for a cell whose model carries no `f`: the harvested
/// per-sheet index, unless the coordinates are unreliable (structural
/// edits shifted them) or the user overwrote the cell this session.
export function indexedFormulaText(
  state: LazyWorkbookState | null,
  sheetId: string,
  row: number,
  column: number,
): string | undefined {
  if (!state) return undefined
  if ((state.editJournal.structuralOps.get(sheetId)?.length ?? 0) > 0) return undefined
  if (state.editJournal.cells.get(sheetId)?.has(`${row}:${column}`)) return undefined
  return state.formulaText.get(sheetId)?.get(`${row}:${column}`)
}

export function installFormulaViewInterceptor(
  runtime: UniverRuntime,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const interceptorService = injector.get(SheetInterceptorService)
  return interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    // Above NUMFMT (10): a formula cell in formula view shows its formula
    // text, not the formatted value, so the chain stops here.
    priority: 9999,
    handler: (cell, location, next) => {
      const state = lazyWorkbookRef.current
      if (!formulaViewSheets(state).has(location.subUnitId)) {
        return next(cell)
      }
      const formula =
        location.rawData?.f ??
        indexedFormulaText(state, location.subUnitId, location.row, location.col)
      if (typeof formula !== 'string' || formula === '') return next(cell)
      return { ...(cell ?? {}), v: formula, t: CellValueType.STRING, p: null }
    },
  })
}

/// Streamed workbooks whose closure gave up have no `f` in the model, so the
/// formula bar shows only values. Inject the harvested formula text
/// into the view model: the grid ignores `f` (displayRawFormula=false) and
/// the engine reads the raw cell matrix, so this is display-only.
export function installFormulaTextInterceptor(
  runtime: UniverRuntime,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const interceptorService = injector.get(SheetInterceptorService)
  const cellContent = interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    priority: 9998,
    handler: (cell, location, next) => {
      if (location.rawData?.f) return next(cell)
      const formula = indexedFormulaText(
        lazyWorkbookRef.current,
        location.subUnitId,
        location.row,
        location.col,
      )
      if (!formula) return next(cell)
      return next({ ...(cell ?? {}), f: formula })
    },
  })
  // The cell EDITOR composes its content through the BEFORE_CELL_EDIT write
  // chain, not CELL_CONTENT — without this hook a streamed formula cell
  // opens for editing as its cached value.
  const beforeEdit = interceptorService.writeCellInterceptor.intercept(BEFORE_CELL_EDIT, {
    handler: (value, context, next) => {
      if (value?.f) return next(value)
      const formula = indexedFormulaText(
        lazyWorkbookRef.current,
        context.subUnitId,
        context.row,
        context.col,
      )
      if (!formula) return next(value)
      return next({ ...(value ?? {}), f: formula })
    },
  })
  // The formula-bar PREVIEW is a separate doc unit that Univer re-syncs from
  // edit-cell-state emissions, and a later value-built emission overwrites a
  // formula-built one for the same cell. Settle it after each burst: when the
  // selected cell has harvested formula text, cover the bar's document with
  // it (the same command Univer's own fx button uses; selection untouched).
  const bridge = injector.get(IEditorBridgeService)
  const commandService = injector.get(ICommandService)
  const univerInstanceService = injector.get(IUniverInstanceService)
  let generation = 0
  const barSync = bridge.currentEditCellState$.subscribe((state) => {
    // Every emission — including value cells and null states — invalidates
    // any pending cover, or a stale timeout rewrites the bar with the
    // previous cell's formula after Univer synced the new selection.
    const token = ++generation
    const cellState = state as {
      sheetId?: string
      row?: number
      column?: number
    } | null
    if (!cellState || cellState.row === undefined || cellState.column === undefined) return
    const sheetId =
      cellState.sheetId ?? runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    if (!sheetId) return
    const formula = indexedFormulaText(
      lazyWorkbookRef.current,
      sheetId,
      cellState.row,
      cellState.column,
    )
    if (!formula) return
    setTimeout(() => {
      if (token !== generation) return
      // Never stomp live typing.
      if (bridge.isVisible().visible) return
      // The bridge's latest state must still be this cell.
      const latest = bridge.getEditCellState() as { row?: number; column?: number } | null
      if (!latest || latest.row !== cellState.row || latest.column !== cellState.column) return
      const doc = univerInstanceService.getUnit<DocumentDataModel>(
        DOCS_FORMULA_BAR_EDITOR_UNIT_ID_KEY,
        UniverInstanceType.UNIVER_DOC,
      )
      const stream = doc?.getBody()?.dataStream ?? ''
      if (stream.replace(/\r\n$/, '') === formula) return
      commandService.syncExecuteCommand(CoverContentCommand.id, {
        unitId: DOCS_FORMULA_BAR_EDITOR_UNIT_ID_KEY,
        body: { dataStream: formula },
        segmentId: '',
      })
    }, 0)
  })
  return {
    dispose() {
      cellContent.dispose()
      beforeEdit()
      barSync.unsubscribe()
    },
  }
}
