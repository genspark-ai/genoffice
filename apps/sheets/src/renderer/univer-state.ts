/**
 * Shared Univer runtime types and lazy-workbook streaming state.
 *
 * Used by both the App component (App.tsx) and the module-level sync
 * helpers (univer-sync.ts).
 */
import { BorderType, LocalUndoRedoService, type IRange } from '@univerjs/core'
import { SheetInterceptorService } from '@univerjs/sheets'

import type {
  WorkbookFile,
  WorkbookPagePrintSettings,
  WorkbookPivotDefinition,
} from '../shared/desktop-api'
import type { createUniver } from './create-univer'
import type { EditJournal } from './edit-journal'
import { netAxisDelta } from './view-transform'

export type UniverRuntime = ReturnType<typeof createUniver>
export type ActiveWorkbook = NonNullable<
  ReturnType<UniverRuntime['univerAPI']['getActiveWorkbook']>
>
export type UniverWorksheet = NonNullable<ReturnType<ActiveWorkbook['getActiveSheet']>>

export interface LazyWorkbookState {
  readonly file: WorkbookFile
  readonly generation: number
  readonly loadedRanges: Map<string, IRange>
  readonly loadingKeys: Map<string, string>
  readonly retryTimers: Map<string, ReturnType<typeof setTimeout>>
  readonly appliedMerges: Map<string, Set<string>>
  readonly appliedRowKeys: Map<string, Set<string>>
  /// Per-sheet union of IStyleData keys carried by <row s= customFormat> and
  /// <col style=> defaults. Univer composes row/col styles into every cell
  /// per-property, but an OOXML cell xf is complete: styled cells null these
  /// keys out so nothing bleeds through (Excel semantics).
  readonly rowColStyleKeys: Map<string, Set<string>>
  readonly appliedCfSheets: Set<string>
  readonly appliedFilterSheets: Set<string>
  readonly appliedDvSheets: Set<string>
  /// Sheets whose last range read predated the sidecar's indexingComplete:
  /// conditional formatting / filters / validations were not yet available,
  /// so an already-loaded range must not satisfy the next load request.
  readonly decorationsPendingSheets: Set<string>
  /// File-side worksheet protection, known once a sheet finishes indexing.
  readonly sheetProtections: Map<string, { protected: boolean; hasPassword: boolean }>
  /// File-side manual page breaks (0-based index of the row/column after the
  /// break, file coordinates), known once a sheet finishes indexing.
  readonly sheetPageBreaks: Map<string, { rowBreaks: number[]; colBreaks: number[] }>
  /// File-side saved print settings (pageSetup / margins / headerFooter),
  /// known once a sheet finishes indexing.
  readonly sheetFilePageSetups: Map<string, WorkbookPagePrintSettings>
  /// File-side allow-edit ranges, known once a sheet finishes indexing.
  readonly sheetProtectedRanges: Map<
    string,
    { name: string; sqref: string; hasPassword: boolean }[]
  >
  /// Defined names the Univer engine rejected at install — preserved verbatim
  /// by the declarative defined-names save.
  readonly uninstalledDefinedNames: Set<string>
  readonly hyperlinkTargets: Map<string, Map<string, string>>
  readonly frozenStripKeys: Map<string, string>
  /// Where each sheet's filter button came from: the worksheet's own
  /// autoFilter (saveable) or a table part (whose filter lives in the table
  /// XML — editing it is blocked).
  readonly filterOrigins: Map<string, { origin: 'worksheet' | 'table'; range: IRange }>
  /// Sheets whose view shows formulas instead of values
  /// (sheetView/@showFormulas): seeded from the file, flipped by the
  /// Formulas-tab toggle, applied
  /// to the global raw-formula render key on sheet activation.
  readonly showFormulaSheets: Set<string>
  /// Small workbooks are fully loaded with formulas handed to Univer's engine
  /// for live recalculation; large ones stream cached values only.
  readonly formulaMode: boolean
  readonly editJournal: EditJournal
  readonly flags: { preloadComplete: boolean }
  /// Closure mode: on streamed workbooks whose formula dependency closure is
  /// small, the closure cells are installed once and pinned (re-applied after
  /// viewport eviction) so the engine recalculates them live.
  readonly closure: {
    status: 'idle' | 'pending' | 'active' | 'unavailable'
    readonly pinned: Map<string, Map<string, PinnedClosureCell>>
  }
  /// Formula text per sheet ('row:col', file coordinates) harvested from
  /// readWorkbookFormulas, so the formula bar can show formulas even when the
  /// closure gave up and the engine never sees them. Display-only.
  readonly formulaText: Map<string, Map<string, string>>
  /// File-cached formula results per sheet ('row:col', screen coordinates),
  /// shown in place of the engine's result when its recalculation errors
  /// (unsupported function, unresolved name). Display-only.
  readonly cachedFormulaValues: Map<string, Map<string, string | number | boolean>>
  /// Parsed pivot definitions keyed by part path, loaded eagerly at open so
  /// pivot refresh stays synchronous.
  readonly pivotDefinitions: Map<string, WorkbookPivotDefinition>
  /// File-declared hidden rows per sheet, recorded as row properties stream
  /// in — lets the viewport loader budget its window by VISIBLE rows.
  readonly hiddenFileRows: Map<string, Set<number>>
  /// Row-property stream coverage per sheet, contiguous from row 0. Beyond
  /// this row the hidden set is incomplete, so consumers ranking by VISIBLE
  /// order (filtered-table stripes) must fall back to physical parity.
  readonly hiddenRowsCoveredThrough: Map<string, number>
  /// Known row/column outline levels per sheet (file reads + this session's
  /// group edits). Rows outside loaded ranges default to level 0.
  readonly outline: Map<
    string,
    {
      readonly rows: Map<number, { level: number; collapsed: boolean }>
      readonly cols: Map<number, { level: number; collapsed: boolean }>
    }
  >
  /// IronCalc fallback when closure mode is unavailable: file formula-cell
  /// keys per sheet, engine values overlaid on viewport patches, and a
  /// session kill switch after the engine rejects the workbook repeatedly.
  readonly recalc: {
    timer: ReturnType<typeof setTimeout> | null
    generation: number
    /// consecutive engine failures; a success resets it
    failures: number
    /// a sheet's formula list came back truncated (>100k formulas): a cold
    /// IronCalc import of such a workbook grinds for minutes and gigabytes,
    /// so the engine fallback is off for the session — cached values stand
    engineOverBudget: boolean
    readonly formulaCells: Map<string, ReadonlySet<number>>
    readonly overlay: Map<string, Map<string, PinnedClosureCell>>
    /// per-sheet: viewport row the last SUCCESSFUL overlay window was
    /// anchored at and whether it covered every formula band; a partial
    /// window re-anchors when the user scrolls far from it (alpha ledger
    /// r141). Written only after the sidecar run succeeds — early writes
    /// latched stale flags on failure (bugbot).
    readonly follow: Map<string, { anchorRow: number; complete: boolean }>
    /// re-anchor throttle: no new run while one is in flight, and at most
    /// one every few seconds — each run reads thousands of sidecar cells
    running: boolean
    lastRunAt: number
  }
}

export interface PinnedClosureCell {
  readonly f?: string
  readonly v?: string | number | boolean | null
}

/// Data extent in screen coordinates: the file extent shifted by this
/// session's structural row/column ops. Null when the sheet is unknown.
export function lazySheetScreenExtent(
  state: LazyWorkbookState,
  sheetId: string,
): { rows: number; columns: number } | null {
  const sheet = state.file.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheet) return null
  const ops = state.editJournal.structuralOps.get(sheetId) ?? []
  return {
    rows: Math.max(sheet.rowCount + netAxisDelta(ops, 'row'), 0),
    columns: Math.max(sheet.columnCount + netAxisDelta(ops, 'column'), 0),
  }
}

/// Budget for closure mode: formula cells plus every precedent they read.
export const CLOSURE_MAX_CELLS = 50_000

/// Streaming re-installs viewport cells through the same mutation user edits
/// produce; this flag keeps programmatic patches out of the edit journal AND
/// out of the undo stack (installJournalSuppressionUndoFilter patches the
/// Univer undo service to drop entries while it is active) — otherwise a
/// freshly opened workbook already "has undo", and undoing would strip loaded
/// file content/layout instead of user edits.
/// Shared mutable state between App.tsx and univer-sync.ts.
export const journalSuppression = { active: false }

/// An AI batch whose applied payload exceeds this many cells keeps no undo
/// entry: the stack retains the full mutation matrices both ways (five
/// 200k-cell copies held ~336MB), and entries accumulate across proposals.
/// The apply path surfaces a "too large to undo" notice instead.
export const AI_UNDO_CELL_BUDGET = 100_000

/// Raised around an AI plan apply. Batching merges each command's small item
/// into the stack-top entry, so the budget must be tracked cumulatively over
/// the activation; `dropped` reports the batch entry was discarded.
export const aiBulkUndoGate = { active: false, dropped: false, cells: 0, pushed: 0 }

export interface UndoRedoItemLike {
  readonly unitID: string
  readonly redoMutations?: readonly { params?: unknown }[]
}

/// Bounded: stops counting past `cap` (the budget check needs no exact total).
export function undoPayloadCells(item: UndoRedoItemLike, cap: number): number {
  let cells = 0
  for (const mutation of item.redoMutations ?? []) {
    const cellValue = (mutation.params as { cellValue?: Record<string, object> } | undefined)
      ?.cellValue
    if (!cellValue || typeof cellValue !== 'object') continue
    for (const rowKey in cellValue) {
      const row = cellValue[rowKey]
      if (!row || typeof row !== 'object') continue
      cells += Object.keys(row).length
      if (cells > cap) return cells
    }
  }
  return cells
}

let undoFilterInstalled = false

/// Drops undo-stack entries pushed while journalSuppression is active.
/// This must patch LocalUndoRedoService.prototype: the DI injector hands out
/// a lazy redi proxy, so assigning a wrapper onto the resolved instance only
/// shadows the proxy — Univer-internal command handlers resolve the real
/// instance and would keep calling the unwrapped method (which is exactly how
/// file opens used to leak load-time set-range-values entries into undo).
export function installJournalSuppressionUndoFilter(): void {
  if (undoFilterInstalled) return
  undoFilterInstalled = true
  const proto = LocalUndoRedoService.prototype as unknown as {
    pushUndoRedo(item: UndoRedoItemLike): void
  }
  const originalPush = proto.pushUndoRedo
  proto.pushUndoRedo = function (this: unknown, item: UndoRedoItemLike) {
    if (journalSuppression.active) return
    if (aiBulkUndoGate.active) {
      if (aiBulkUndoGate.dropped) return
      aiBulkUndoGate.cells += undoPayloadCells(item, AI_UNDO_CELL_BUDGET + 1)
      if (aiBulkUndoGate.cells > AI_UNDO_CELL_BUDGET) {
        aiBulkUndoGate.dropped = true
        // Chunks already merged into the stack-top batch entry must go too —
        // keeping them would make undo revert only part of the operation.
        const service = this as {
          _getUndoStack?: (unitId: string) => unknown[] | undefined
          _getRedoStack?: (unitId: string) => unknown[] | undefined
          _updateStatus?: () => void
        }
        if (aiBulkUndoGate.pushed > 0) {
          const stack = service._getUndoStack?.(item.unitID)
          if (Array.isArray(stack) && stack.length > 0) stack.pop()
        }
        // The original push clears redo on entry; a batch dropped on its
        // first item must not leave a stale redo replayable over the change.
        const redoStack = service._getRedoStack?.(item.unitID)
        if (Array.isArray(redoStack)) redoStack.length = 0
        service._updateStatus?.()
        return
      }
      aiBulkUndoGate.pushed += 1
    }
    originalPush.call(this, item)
  }
}

/// Excel never re-measures row heights when opening a file: a row shows its
/// stored ht (or the sheet default) and wrapped/tall content is clipped.
/// Univer's AutoHeightController re-measures every auto (ia≠0) row touched by
/// SetRangeValues-style commands, so streaming file content into the grid
/// ballooned wrapped rows. While this flag is up, the auto-height interceptor
/// yields nothing; rows keep ia=1 so later USER edits still auto-fit.
export const loadAutoHeightSuppression = { active: false }

let autoHeightGateInstalled = false

/// Same prototype-patch shape as the undo filter above: command handlers
/// resolve the real service, so wrapping the resolved instance is not enough.
export function installLoadAutoHeightGate(): void {
  if (autoHeightGateInstalled) return
  autoHeightGateInstalled = true
  type AutoHeightMutations = {
    preUndos: unknown[]
    undos: unknown[]
    preRedos: unknown[]
    redos: unknown[]
  }
  const proto = SheetInterceptorService.prototype as unknown as {
    generateMutationsOfAutoHeight(ctx: unknown): AutoHeightMutations
  }
  const original = proto.generateMutationsOfAutoHeight
  proto.generateMutationsOfAutoHeight = function (this: unknown, ctx: unknown) {
    if (loadAutoHeightSuppression.active) {
      return { preUndos: [], undos: [], preRedos: [], redos: [] }
    }
    return original.call(this, ctx)
  }
}

export const BORDER_COMMAND_TYPES: Record<string, BorderType> = {
  all: BorderType.ALL,
  outer: BorderType.OUTSIDE,
  'thick-outer': BorderType.OUTSIDE,
  top: BorderType.TOP,
  bottom: BorderType.BOTTOM,
  left: BorderType.LEFT,
  right: BorderType.RIGHT,
  none: BorderType.NONE,
}
