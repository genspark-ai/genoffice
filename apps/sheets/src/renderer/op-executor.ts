/**
 * Applies a planned batch of workbook operations to the live Univer workbook:
 * prechecks, one undo item per batch, per-op facade calls and journal
 * records. Extracted from App.tsx so ribbon actions can run the same ops the
 * AI proposes; every function receives its App-scope context explicitly.
 */
import { IUndoRedoService } from '@univerjs/core'
import type { CellBounds } from '../domain/chart-visual'
import {
  columnIndex,
  columnLabel,
  formatAddress,
  parseAddress,
  parseRange,
} from '../domain/cell-address'
import { offsetFormulaRefs } from '../domain/formula-shift'
import { computeSortedRowOrder } from '../domain/sort-range'
import {
  copyTargetBounds,
  filteredCopySourceRows,
  matchableCellText,
  replaceOccurrences,
  type WorkbookOperation,
} from '../domain/workbook-dsl'
import type { ApplyOutcome, ChangePlan, StructuralChange } from '../domain/workbook.types'
import { pushBulkFillUndo } from './bulk-fill-undo'
import {
  journalCellContentAt,
  plainCellValue,
  journalSize,
  recordBulkConstantFill,
  recordPageSetup,
  recordSheetProtection,
  recordSparklineAdd,
  removeBulkConstantFill,
  removeTableAdd,
  restoreJournalCells,
  type PageSetupJournalState,
} from './edit-journal'
import { indexedFormulaText } from './formula-view'
import { buildLazyChangePlan } from './lazy-plan'
import { t } from './i18n/locale'
import type { PivotActionContext } from './pivot-actions'
import { refreshPivotTables } from './pivot-actions'
import {
  carryCopyFormulasPlan,
  collectStreamedFormulaPrecedents,
  lazyGateFailure,
  streamedRefSheetList,
  structuralDeleteFormulaError,
  type StreamedRefSheet,
} from './plan-operations'
import { aiBulkUndoGate, journalSuppression } from './univer-state'
import type {
  ActiveWorkbook,
  LazyWorkbookState,
  UniverRuntime,
  UniverWorksheet,
} from './univer-state'
import {
  absRangeRef,
  applyAiConditionalFormat,
  applyAiDataValidation,
  applyAiHyperlink,
  applyFilterCriteria,
  applyFormatPatchToRange,
  applyJournalOverlay,
  applyRangeInLoadedChunks,
  lazyWorkbookCellReader,
  loadVisibleRange,
  measureImage,
  pinStreamedPrecedents,
  protectSheetGuard,
  pushVisualUndo,
  queueSparklineInstall,
  readCopySourceDirect,
  sniffImageMime,
  workbookStructureLocked,
} from './univer-sync'
import {
  applyAiShapeEdit,
  buildAiChartEdit,
  insertAiChartVisual,
  insertAiImageVisual,
  insertAiShapeVisual,
  type VisualActionContext,
} from './visual-actions'
import {
  applyAiPivotAdd,
  applyAiTableAdd,
  applyAiTableColumnAdd,
  applyAiTableColumnDelete,
  applyAiTableRowAdd,
  applyAiTableRowDelete,
} from './workbook-ops'
import type { ChartEditData, ShapeEditChanges } from './WorkbookVisuals'

export type PlannedOp = StructuralChange['op']

export interface LoadedImage {
  dataUrl: string
  mediaType: string
  width: number
  height: number
}

/** App-scope refs/state the executors need; built fresh per apply by App. */
export interface OpExecutorContext {
  runtime: UniverRuntime
  workbook: ActiveWorkbook
  state: LazyWorkbookState
  lazyWorkbookRef: { current: LazyWorkbookState | null }
  sparklineDisposablesRef: { current: { dispose(): void }[] }
  sparklineTimerRef: { current: ReturnType<typeof setTimeout> | null }
  shapeEditRef: { current: (visualId: string, changes: ShapeEditChanges) => void }
  chartEditRef: { current: (chartPath: string, edit: ChartEditData) => void }
  setMessage: (message: string) => void
  setPendingEdits: (count: number) => void
  visualContext: () => VisualActionContext
  pivotContext: () => PivotActionContext
  queueChartDataSync: (sheetId: string, bounds: CellBounds) => void
}

/** Per-batch scratch shared by the op executors. */
interface OpRun {
  readonly ctx: OpExecutorContext
  readonly workbook: ActiveWorkbook
  readonly sheetById: (id: string) => UniverWorksheet
  readonly plannedOps: readonly PlannedOp[]
  /// Previewed after-values of the batch's single-cell changes, keyed
  /// "sheetId:row:column" — a fill source may be one of them.
  readonly plannedCellContents: ReadonlyMap<
    string,
    { value: string | number | boolean | null; formula: string | null }
  >
  readonly imageData: ReadonlyMap<string, LoadedImage>
  readonly notices: string[]
  /// Called only AFTER a mutation actually committed (end of each op, plus the
  /// intermediate commit points of multi-step handlers): a throw before any
  /// call really is "unchanged".
  readonly markApplied: () => void
  readonly userFacing: boolean
}

export interface ApplyPlanOptions {
  /** Runs after the async prechecks and right before the first mutation; a
   * non-null outcome aborts with that outcome untouched. */
  verifyBeforeApply?: () => ApplyOutcome | null
  /** Status-bar text on success; null keeps the status bar as is. */
  successMessage?: string | null
  /** Ribbon/dialog caller: gate refusals show their localized string instead
   * of the model-facing reason. */
  userFacing?: boolean
}

export function richCellText(cell: unknown): string | null {
  const stream = (cell as { p?: { body?: { dataStream?: unknown } } } | null | undefined)?.p?.body
    ?.dataStream
  return typeof stream === 'string' ? stream.replace(/\r\n$/, '').replace(/\r/g, '\n') : null
}

const SHEET_LIFECYCLE_OPS = new Set([
  'add_sheet',
  'delete_sheet',
  'duplicate_sheet',
  'move_sheet',
  'set_sheet_hidden',
  'rename_sheet',
])

/**
 * Image bytes load BEFORE any drift check and the (synchronous) mutation
 * loop, so a slow disk read can never interleave with edits.
 */
export async function prefetchOpImages(
  ops: readonly PlannedOp[],
): Promise<Map<string, LoadedImage>> {
  const imageData = new Map<string, LoadedImage>()
  for (const op of ops) {
    if (op.op !== 'add_image' || imageData.has(op.path)) continue
    let dataUrl: string
    let mediaType: string
    // file:// = a BYOK-generated image in the local store (fetchImage resolves it)
    if (/^(https?|file):\/\//i.test(op.path)) {
      const fetched = await window.desktopApi.fetchImage(op.path)
      if (!fetched) throw new Error(t('appCannotReadImage'))
      // Trust the bytes, not the Content-Type header the handler echoed
      const sniffed = sniffImageMime(fetched.base64)
      if (!sniffed) {
        throw new Error('The downloaded image is not PNG/JPEG/GIF — pick another image URL.')
      }
      dataUrl = `data:${sniffed};base64,${fetched.base64}`
      mediaType = sniffed
    } else {
      const image = await window.desktopApi.readLocalImage({ path: op.path })
      dataUrl = `data:${image.mediaType};base64,${image.base64}`
      mediaType = image.mediaType
    }
    const size = await measureImage(dataUrl)
    imageData.set(op.path, { dataUrl, mediaType, ...size })
  }
  return imageData
}

/**
 * Names the formula a batch's first row/column-shifting delete would strand,
 * so the caller rejects before ANY op runs (the save-time guard would otherwise abort at
 * ⌘S). Only the FIRST shifting op is checked: pre-batch formula texts are
 * exact for it, while any later delete sees coordinates an earlier shift
 * already moved (a stale check would false-reject); those keep the save-time
 * guard as the backstop. Table row/column deletes are whole-sheet deletes
 * underneath, so they translate to the same span check.
 */
export async function precheckStructuralDeletes(
  state: LazyWorkbookState,
  workbook: ActiveWorkbook,
  ops: readonly PlannedOp[],
): Promise<string | null> {
  const deleteSpanOf = (
    op: PlannedOp,
  ):
    | { op: 'delete_rows'; sheetId: string; row: number; count: number }
    | { op: 'delete_cols'; sheetId: string; column: string; count: number }
    | 'shifts'
    | null => {
    if (op.op === 'delete_rows' || op.op === 'delete_cols') return op
    if (op.op === 'delete_table_row' || op.op === 'delete_table_column') {
      const entry = state.editJournal.tableAdds.find(
        (table) =>
          table.sheetId === op.sheetId && table.name.toLowerCase() === op.tableName.toLowerCase(),
      )
      if (!entry) return 'shifts'
      // mirrors applyAiTableRowDelete/applyAiTableColumnDelete addressing
      return op.op === 'delete_table_row'
        ? {
            op: 'delete_rows',
            sheetId: op.sheetId,
            row: entry.area.startRow + op.row + 1,
            count: op.count ?? 1,
          }
        : {
            op: 'delete_cols',
            sheetId: op.sheetId,
            column: columnLabel(entry.area.startColumn + op.column - 1),
            count: op.count ?? 1,
          }
    }
    if (
      op.op === 'insert_rows' ||
      op.op === 'insert_cols' ||
      op.op === 'add_table_row' ||
      op.op === 'add_table_column'
    ) {
      return 'shifts'
    }
    return null
  }
  for (const op of ops) {
    const translated = deleteSpanOf(op)
    if (translated === null) continue
    if (translated !== 'shifts') {
      const spanError = await structuralDeleteFormulaError(state, workbook, translated)
      if (spanError) return spanError
    }
    break
  }
  return null
}

/**
 * All commands of one batch merge into a single undo item (⌘Z / [Undo] rolls
 * back the whole batch in one step). Disposing pushes the batched item
 * through the undo gate, so the success path must settle before reading
 * `aiBulkUndoGate.dropped`.
 */
export function beginUndoBatch(runtime: UniverRuntime): { settle(): void } {
  const batchUnitId = runtime.univerAPI.getActiveWorkbook()?.getId()
  const undoBatching = batchUnitId
    ? runtime.univer.__getInjector().get(IUndoRedoService).__tempBatchingUndoRedo(batchUnitId)
    : null
  aiBulkUndoGate.active = true
  aiBulkUndoGate.dropped = false
  aiBulkUndoGate.cells = 0
  aiBulkUndoGate.pushed = 0
  let settled = false
  return {
    settle() {
      if (settled) return
      settled = true
      try {
        undoBatching?.dispose()
      } finally {
        aiBulkUndoGate.active = false
      }
    },
  }
}

/** Wraps ops the UI built (ribbon, dialogs) into the plan shape the AI path applies. */
export function planFromOps(
  ops: readonly WorkbookOperation[],
  workbook: ActiveWorkbook,
  summary: string,
): ChangePlan {
  return buildLazyChangePlan(
    {
      dslVersion: 1,
      transactionId: `ui-${Date.now().toString(36)}`,
      baseRevision: 0,
      summary,
      operations: [...ops],
    },
    lazyWorkbookCellReader(workbook),
    (id) => {
      const sheet = workbook.getSheetBySheetId(id)
      if (!sheet) throw new Error(`Unknown sheet: ${id}`)
      return sheet.getSheetName()
    },
  )
}

/** Excel-style default name for an inserted sheet: the lowest free SheetN. */
export function nextSheetName(taken: readonly string[]): string {
  const lower = new Set(taken.map((name) => name.toLowerCase()))
  for (let n = 1; ; n += 1) {
    const candidate = `Sheet${n}`
    if (!lower.has(candidate.toLowerCase())) return candidate
  }
}

/**
 * Applies a plan: prechecks, then every operation inside one undo batch —
 * range/structural ops first (so inserts establish the final coordinate
 * space), then cell, format and rename changes.
 */
let applyQueue: Promise<void> = Promise.resolve()

export function applyChangePlan(
  plan: ChangePlan,
  ctx: OpExecutorContext,
  options: ApplyPlanOptions = {},
): Promise<ApplyOutcome> {
  // One apply at a time: the undo batch and aiBulkUndoGate are not reentrant,
  // and the async prechecks yield — a ribbon click landing during an AI apply
  // (or a second click during a delete precheck) must wait its turn.
  const run = applyQueue.then(() => applyChangePlanNow(plan, ctx, options))
  applyQueue = run.then(
    () => undefined,
    () => undefined,
  )
  return run
}

async function applyChangePlanNow(
  plan: ChangePlan,
  ctx: OpExecutorContext,
  options: ApplyPlanOptions,
): Promise<ApplyOutcome> {
  const { runtime, workbook, state, lazyWorkbookRef, setMessage, setPendingEdits } = ctx
  const plannedOps = plan.structuralChanges.map((structural) => structural.op)
  const notices: string[] = []
  let imageData: Map<string, LoadedImage>
  try {
    imageData = await prefetchOpImages(plannedOps)
    const spanError = await precheckStructuralDeletes(state, workbook, plannedOps)
    if (spanError) throw new Error(options.userFacing ? t('appDeleteSpanFormulas') : spanError)
  } catch (error: unknown) {
    const reason = error instanceof Error ? error.message : t('appCannotReadImage')
    setMessage(reason)
    return { ok: false, reason }
  }
  const stale = options.verifyBeforeApply?.()
  if (stale) return stale
  const undoBatch = beginUndoBatch(runtime)
  let anyApplied = false
  try {
    // Structural and layout changes go through the same facade commands as
    // the ribbon, so BeforeCommandExecute gating and the edit journal apply.
    // Every operation routes to its own sheetId — the active sheet at
    // propose time has no special role.
    const sheetById = (id: string): UniverWorksheet => {
      const found = workbook.getSheetBySheetId(id)
      if (!found) throw new Error(`Unknown sheet: ${id}`)
      return found
    }
    // Range operations execute before cellChanges so structural inserts
    // establish the final coordinate space first. A fill source can itself
    // be one of those pending cell changes (for example CT2="north"
    // followed by fill CT2:CT88588), so make the previewed after-value
    // available before it has landed in Univer or the journal.
    const plannedCellContents = new Map<
      string,
      { value: string | number | boolean | null; formula: string | null }
    >()
    for (const change of plan.cellChanges) {
      const cell = parseRange(change.address)
      if (cell.startRow !== cell.endRow || cell.startColumn !== cell.endColumn) continue
      plannedCellContents.set(`${change.sheetId}:${cell.startRow}:${cell.startColumn}`, {
        value: change.after.value,
        formula: change.after.formula ?? null,
      })
    }
    const run: OpRun = {
      ctx,
      workbook,
      sheetById,
      plannedOps,
      plannedCellContents,
      imageData,
      notices,
      markApplied: () => {
        anyApplied = true
      },
      userFacing: options.userFacing ?? false,
    }
    // Checked for the whole batch BEFORE anything applies — a mid-loop
    // throw would leave earlier ops committed.
    if (workbookStructureLocked(state) && plannedOps.some((op) => SHEET_LIFECYCLE_OPS.has(op.op))) {
      throw new Error(t('appWorkbookStructureLocked'))
    }
    for (const op of plannedOps) await executeOp(op, run)
    setPendingEdits(journalSize(state.editJournal))
    // Streaming mode: formulas in this batch may reference file cells the
    // engine does not hold — load & pin them first so the formulas compute
    // correctly and survive eviction (propose checked the session budget).
    if (!state.formulaMode) {
      const streamedNeeds = new Map<string, Set<number>>()
      let refSheets: StreamedRefSheet[] | undefined
      for (const change of plan.cellChanges) {
        if (!change.after.formula) continue
        refSheets ??= streamedRefSheetList(state, workbook)
        collectStreamedFormulaPrecedents(
          change.after.formula,
          change.sheetId,
          refSheets,
          streamedNeeds,
          state.closure.pinned,
        )
      }
      if (
        streamedNeeds.size > 0 &&
        !(await pinStreamedPrecedents(runtime, lazyWorkbookRef, streamedNeeds))
      ) {
        throw new Error(
          'The cells referenced by the formulas could not be loaded from the streamed workbook — ' +
            'the formulas were not written; retry in a moment.',
        )
      }
    }
    for (const change of plan.cellChanges) {
      const range = sheetById(change.sheetId).getRange(change.address)
      if (change.after.formula) range.setFormula(change.after.formula)
      else if (change.after.value === null) range.clearContent()
      else {
        // Explicit f/si null mirrors the cell editor: overwriting a formula
        // cell with a value must clear the formula (in Univer and journal).
        // A rich-text target also needs p cleared, or setValues merges and
        // the old document keeps rendering over the new value.
        const wasRich = range.getCellDatas()[0]?.[0]?.p != null
        range.setValues([
          wasRich
            ? [{ v: change.after.value, f: null, si: null, p: null }]
            : [{ v: change.after.value, f: null, si: null }],
        ])
      }
      run.markApplied()
    }
    // Same facade setters as the ribbon, so the edit journal records them
    // (indent included — it lands as a pd patch in set-range-values).
    for (const formatChange of plan.formatChanges) {
      applyFormatPatchToRange(
        sheetById(formatChange.sheetId).getRange(formatChange.range),
        formatChange.format,
      )
      run.markApplied()
    }
    for (const rename of plan.sheetRenames) {
      sheetById(rename.sheetId).setName(rename.after)
      run.markApplied()
    }
    if (options.successMessage !== null) {
      setMessage(options.successMessage ?? t('appAppliedJournaled'))
    }
    undoBatch.settle()
    if (aiBulkUndoGate.dropped) {
      notices.push(
        'this change is too large for the undo history — the [Undo] button and ⌘Z will not revert it',
      )
    }
    return notices.length > 0 ? { ok: true, notices } : { ok: true }
  } catch (error: unknown) {
    const reason = error instanceof Error ? error.message : t('appApplyTxFailed')
    setMessage(reason)
    // Settle here, not just in finally: the return value is computed first,
    // and a budget-dropped batch must not advertise the ⌘Z rollback.
    undoBatch.settle()
    return anyApplied
      ? { ok: false, reason, partiallyApplied: true, undoDropped: aiBulkUndoGate.dropped }
      : { ok: false, reason }
  } finally {
    undoBatch.settle()
  }
}

/** One structural / layout / range-level op against the live workbook. */
async function executeOp(op: PlannedOp, run: OpRun): Promise<void> {
  const { ctx, workbook, sheetById, plannedCellContents, imageData, notices } = run
  const { runtime, state, lazyWorkbookRef, setMessage, setPendingEdits } = ctx
  // BeforeCommandExecute gates cancel commands silently; re-check them
  // here so a gated op fails loud instead of reporting success.
  const gate = lazyGateFailure(state, op)
  if (gate) throw new Error(run.userFacing && gate.messageKey ? t(gate.messageKey) : gate.reason)
  if (op.op === 'insert_rows') sheetById(op.sheetId).insertRowsBefore(op.row - 1, op.count)
  else if (op.op === 'delete_rows') sheetById(op.sheetId).deleteRows(op.row - 1, op.count)
  else if (op.op === 'insert_cols')
    sheetById(op.sheetId).insertColumnsBefore(columnIndex(op.column), op.count)
  else if (op.op === 'delete_cols')
    sheetById(op.sheetId).deleteColumns(columnIndex(op.column), op.count)
  else if (op.op === 'add_sheet') {
    // A UI caller picks its SheetN before the queue turn; re-derive against
    // the live list so two quick clicks cannot collide.
    const taken = workbook.getSheets().map((sheet) => sheet.getSheetName())
    const name =
      run.userFacing && taken.some((existing) => existing.toLowerCase() === op.name.toLowerCase())
        ? nextSheetName(taken)
        : op.name
    workbook.insertSheet(
      name,
      op.rows !== undefined || op.columns !== undefined
        ? {
            sheet: {
              ...(op.rows !== undefined ? { rowCount: op.rows } : {}),
              ...(op.columns !== undefined ? { columnCount: op.columns } : {}),
            },
          }
        : undefined,
    )
  } else if (op.op === 'delete_sheet') workbook.deleteSheet(op.sheetId)
  else if (op.op === 'merge_cells') sheetById(op.sheetId).getRange(op.range).merge()
  else if (op.op === 'unmerge_cells') sheetById(op.sheetId).getRange(op.range).breakApart()
  else if (op.op === 'set_row_height') {
    sheetById(op.sheetId).setRowHeights(
      op.row - 1,
      op.count,
      Math.round((op.heightPoints * 96) / 72),
    )
  } else if (op.op === 'set_col_width') {
    sheetById(op.sheetId).setColumnWidths(columnIndex(op.column), op.count, Math.round(op.widthPx))
  } else if (op.op === 'set_rows_hidden') {
    if (op.hidden) sheetById(op.sheetId).hideRows(op.row - 1, op.count)
    else sheetById(op.sheetId).showRows(op.row - 1, op.count)
  } else if (op.op === 'set_cols_hidden') {
    if (op.hidden) sheetById(op.sheetId).hideColumns(columnIndex(op.column), op.count)
    else sheetById(op.sheetId).showColumns(columnIndex(op.column), op.count)
  } else if (op.op === 'duplicate_sheet') {
    if (!workbook) throw new Error(t('appNoWorkbookOpen'))
    const copy = workbook.duplicateSheet(sheetById(op.sheetId))
    run.markApplied()
    if (op.name) copy.setName(op.name)
  } else if (op.op === 'set_sheet_hidden') {
    if (op.hidden) sheetById(op.sheetId).hideSheet()
    else sheetById(op.sheetId).showSheet()
  } else if (op.op === 'move_sheet') {
    if (!workbook) throw new Error(t('appNoWorkbookOpen'))
    workbook.moveSheet(sheetById(op.sheetId), op.position - 1)
  } else if (op.op === 'add_sparkline') {
    const target = workbook?.getSheetBySheetId(op.sheetId)
    if (!target) throw new Error(`Unknown sheet: ${op.sheetId}`)
    const bounds = parseRange(op.dataRange)
    const rows = Math.min(bounds.endRow - bounds.startRow + 1, 200)
    const base =
      op.targetCell === undefined
        ? { row: bounds.startRow, column: bounds.endColumn + 1 }
        : parseAddress(op.targetCell)
    const sheetName = target.getSheetName()
    const cells = Array.from({ length: rows }, (_, offset) => ({
      cell: `${columnLabel(base.column)}${base.row + offset + 1}`,
      sourceRef: absRangeRef(
        sheetName,
        `${columnLabel(bounds.startColumn)}${bounds.startRow + offset + 1}` +
          `:${columnLabel(bounds.endColumn)}${bounds.startRow + offset + 1}`,
      ),
    }))
    recordSparklineAdd(state.editJournal, {
      id: `sparkline-${Date.now().toString(36)}-${state.editJournal.sparklineAdds.length + 1}`,
      sheetId: op.sheetId,
      type: op.type,
      ...(op.color === undefined ? {} : { color: op.color }),
      cells,
    })
    setPendingEdits(journalSize(state.editJournal))
    queueSparklineInstall(
      runtime,
      lazyWorkbookRef,
      ctx.sparklineDisposablesRef,
      ctx.sparklineTimerRef,
    )
  } else if (op.op === 'delete_visual') {
    const visual = [...state.file.visuals, ...state.editJournal.visualAdds].find(
      (candidate) => candidate.id === op.visualId || candidate.chartPath === op.visualId,
    )
    if (!visual) throw new Error(`Unknown visual: ${op.visualId}`)
    ctx.shapeEditRef.current(visual.id, { remove: true })
  } else if (op.op === 'delete_table') {
    if (!removeTableAdd(state.editJournal, op.sheetId, op.tableName)) {
      throw new Error(t('appTableNotDeletable', { name: op.tableName }))
    }
    setPendingEdits(journalSize(state.editJournal))
  } else if (op.op === 'add_chart') {
    await insertAiChartVisual(ctx.visualContext(), runtime, state, op)
  } else if (op.op === 'add_shape') {
    insertAiShapeVisual(ctx.visualContext(), runtime, state, op)
  } else if (op.op === 'edit_shape') {
    applyAiShapeEdit(ctx.visualContext(), runtime, state, op)
  } else if (op.op === 'add_image') {
    const image = imageData.get(op.path)
    if (!image) throw new Error(t('appImageNotLoaded', { path: op.path }))
    insertAiImageVisual(ctx.visualContext(), runtime, state, op, image)
  } else if (op.op === 'add_table') {
    applyAiTableAdd(runtime, state, op)
  } else if (op.op === 'add_table_row') {
    applyAiTableRowAdd(runtime, state, op)
  } else if (op.op === 'add_table_column') {
    applyAiTableColumnAdd(runtime, state, op)
  } else if (op.op === 'delete_table_row') {
    applyAiTableRowDelete(runtime, state, op)
  } else if (op.op === 'delete_table_column') {
    applyAiTableColumnDelete(runtime, state, op)
  } else if (op.op === 'add_pivot') {
    applyAiPivotAdd(runtime, state, op)
  } else if (op.op === 'set_hyperlink') {
    applyAiHyperlink(state, sheetById(op.sheetId), op)
  } else if (op.op === 'protect_sheet') {
    const guard = protectSheetGuard(state, op.sheetId, op.protected)
    if (guard) throw new Error(guard)
    const original = state.sheetProtections.get(op.sheetId)?.protected ?? false
    recordSheetProtection(state.editJournal, op.sheetId, op.protected, original)
  } else if (op.op === 'set_filter') {
    const target = sheetById(op.sheetId)
    const existing = target.getFilter()
    if (existing) {
      existing.remove()
      // Only count a verified removal — the remove command can be
      // cancelled by an edit gate.
      if (target.getFilter()) throw new Error(t('appAutoFilterRemoveFailed'))
      run.markApplied()
    }
    if (!target.getRange(op.range).createFilter()) {
      throw new Error(t('appAutoFilterCreateFailed'))
    }
  } else if (op.op === 'clear_filter') {
    const target = sheetById(op.sheetId)
    target.getFilter()?.remove()
    if (target.getFilter()) {
      throw new Error(t('appAutoFilterRemoveFailed'))
    }
  } else if (op.op === 'set_filter_criteria') {
    applyFilterCriteria(
      sheetById(op.sheetId),
      op.column,
      op.values === null ? null : { values: op.values },
    )
  } else if (op.op === 'add_conditional_format') {
    applyAiConditionalFormat(sheetById(op.sheetId), op)
  } else if (op.op === 'clear_conditional_formats') {
    const target = sheetById(op.sheetId)
    for (const rule of target.getConditionalFormattingRules()) {
      if (rule.cfId) {
        target.deleteConditionalFormattingRule(rule.cfId)
        run.markApplied()
      }
    }
  } else if (op.op === 'set_data_validation') {
    applyAiDataValidation(runtime, sheetById(op.sheetId), op)
  } else if (op.op === 'add_defined_name') {
    if (!workbook) throw new Error(t('appNoWorkbookOpen'))
    workbook.insertDefinedName(op.name, op.ref)
  } else if (op.op === 'delete_defined_name') {
    if (!workbook) throw new Error(t('appNoWorkbookOpen'))
    workbook.deleteDefinedName(op.name)
  } else if (op.op === 'set_page_setup') {
    sheetById(op.sheetId)
    const prior = state.editJournal.pageSetup.get(op.sheetId) ?? {}
    const patch: PageSetupJournalState = {}
    if (op.orientation !== undefined) patch.orientation = op.orientation
    if (op.paperSize !== undefined) patch.paperSize = op.paperSize
    if (op.margins !== undefined) patch.margins = op.margins
    if (op.printGridlines !== undefined) patch.printGridlines = op.printGridlines
    if (op.printHeadings !== undefined) patch.printHeadings = op.printHeadings
    if (op.printArea !== undefined) patch.printArea = op.printArea
    // Scale and fit-to-page are exclusive; whichever the op sets wins,
    // and a fit on one axis keeps the other axis' prior value.
    if (op.scale !== undefined) {
      patch.scale = op.scale
      patch.fitToPage = false
    } else if (op.fitToWidth !== undefined || op.fitToHeight !== undefined) {
      patch.fitToWidth = op.fitToWidth ?? prior.fitToWidth ?? 0
      patch.fitToHeight = op.fitToHeight ?? prior.fitToHeight ?? 0
      patch.fitToPage = patch.fitToWidth > 0 || patch.fitToHeight > 0
    }
    recordPageSetup(state.editJournal, op.sheetId, patch)
  } else if (op.op === 'set_freeze') {
    // Journaled by the set-frozen mutation listener.
    const target = sheetById(op.sheetId)
    if (op.rows === 0 && op.columns === 0) {
      target.cancelFreeze()
    } else {
      target.setFreeze({
        startRow: op.rows > 0 ? op.rows : -1,
        startColumn: op.columns > 0 ? op.columns : -1,
        xSplit: op.columns,
        ySplit: op.rows,
      })
    }
  } else if (op.op === 'refresh_pivot') {
    refreshPivotTables(ctx.pivotContext(), op.sheetId)
  } else if (op.op === 'clear_range') {
    // Range-level clear (>2000 cells). On streamed workbooks each chunk
    // is loaded first so the clear lands on real cells and the edit
    // journal records it; preloaded workbooks clear in one command.
    const targetSheet = sheetById(op.sheetId)
    await applyRangeInLoadedChunks(
      runtime,
      lazyWorkbookRef,
      targetSheet,
      parseRange(op.range),
      (chunk) => {
        targetSheet
          .getRange(
            chunk.startRow,
            chunk.startColumn,
            chunk.endRow - chunk.startRow + 1,
            chunk.endColumn - chunk.startColumn + 1,
          )
          .clearContent()
      },
      setMessage,
      // Clearing writes no formulas; neighbor columns are irrelevant.
      { neighborColumns: false },
    )
  } else if (op.op === 'fill_range') {
    // Fill/copy: tile the source block across the target with bulk
    // setValues (chunked on streamed workbooks). Relative formula
    // references shift per copy, exactly like Excel's fill handle;
    // validation (geometry, source loaded, cost) ran at propose time.
    const targetSheet = sheetById(op.sheetId)
    const sourceSheet = sheetById(op.sourceSheetId ?? op.sheetId)
    const src = parseRange(op.source)
    const dst = parseRange(op.target)
    const sourceRows = src.endRow - src.startRow + 1
    const sourceColumns = src.endColumn - src.startColumn + 1
    type FillSourceCell = {
      value: string | number | boolean | null
      formula: string | null
      t: number | null
      s: Record<string, unknown> | null
    }
    // Source contents are captured up front: chunk loading evicts the
    // current window, so later reads from the grid could see blanks.
    const lazyState = lazyWorkbookRef.current
    const sourceSheetId = op.sourceSheetId ?? op.sheetId
    // Raw model values and resolved styles, like copy_range: the view
    // model's getValue() returns display text for dates/formatted
    // numbers, and style-pool id strings save without their format.
    const fillStyles = runtime.univerAPI.getActiveWorkbook()?.getWorkbook().getStyles()
    const sourceCells: FillSourceCell[][] = []
    for (let row = 0; row < sourceRows; row += 1) {
      const rowCells: FillSourceCell[] = []
      for (let column = 0; column < sourceColumns; column += 1) {
        const sourceRow = src.startRow + row
        const sourceColumn = src.startColumn + column
        const cell = sourceSheet.getRange(formatAddress(sourceRow, sourceColumn))
        const plannedCell = plannedCellContents.get(`${sourceSheetId}:${sourceRow}:${sourceColumn}`)
        const journalCell = lazyState
          ? journalCellContentAt(lazyState.editJournal, sourceSheetId, sourceRow, sourceColumn)
          : { found: false as const }
        const raw = cell.getCellDatas()[0]?.[0]
        const style = raw?.s ? (fillStyles?.getStyleByCell(raw) ?? null) : null
        rowCells.push({
          value:
            plannedCell !== undefined
              ? plannedCell.value
              : journalCell.found
                ? journalCell.value
                : ((plainCellValue(raw?.v, raw?.t) ??
                    cell.getValue() ??
                    null) as FillSourceCell['value']),
          formula:
            plannedCell !== undefined
              ? plannedCell.formula
              : journalCell.found
                ? journalCell.formula
                : cell.getFormula() || null,
          // Overlays supply value/formula (grid may lag the journal),
          // but the resident grid cell still holds the latest applied
          // style — session-formatted sources must not go out bare.
          t: raw?.t ?? null,
          s: style ? (style as Record<string, unknown>) : null,
        })
      }
      sourceCells.push(rowCells)
    }
    type FillMatrixCell = {
      v: string | number | boolean | null
      f: string | null
      si: null
      t?: number
      s?: Record<string, unknown>
    }
    // Constant fills skip neighbor-column loads entirely: a target in a
    // freshly inserted column is then journal-owned and applies without
    // waiting for background indexing. Formula fills still need real
    // neighbor values to compute against.
    const fillWritesFormulas = sourceCells.some((row) => row.some((cell) => cell.formula !== null))
    // The journal fast path records a bare value — a styled source
    // (e.g. a date) would save without its number format, showing raw
    // serials. Those fills take the chunked path, which carries t/s.
    const fillCarriesStyle = sourceCells[0]?.[0]?.s != null
    if (
      lazyState &&
      !fillWritesFormulas &&
      !fillCarriesStyle &&
      sourceRows === 1 &&
      sourceColumns === 1
    ) {
      const { fill, purgedCells } = recordBulkConstantFill(lazyState.editJournal, {
        sheetId: op.sheetId,
        startRow: dst.startRow,
        endRow: dst.endRow,
        startColumn: dst.startColumn,
        endColumn: dst.endColumn,
        value: sourceCells[0]?.[0]?.value ?? null,
      })
      const targetIsInsertedThisSession =
        run.plannedOps.some((plannedOp) => {
          if (plannedOp.op !== 'insert_cols' || plannedOp.sheetId !== op.sheetId) return false
          const index = columnIndex(plannedOp.column)
          return dst.startColumn >= index && dst.endColumn < index + plannedOp.count
        }) ||
        (lazyState.editJournal.structuralOps.get(op.sheetId) ?? []).some(
          (structuralOp) =>
            structuralOp.kind === 'insert-cols' &&
            dst.startColumn >= structuralOp.index &&
            dst.endColumn < structuralOp.index + structuralOp.count,
        )
      const refresh = (direction: 'apply' | 'undo' = 'apply') => {
        const current = lazyWorkbookRef.current
        if (!current) return
        setPendingEdits(journalSize(current.editJournal))
        const active = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
        if (active?.getSheetId() !== op.sheetId) return
        if (targetIsInsertedThisSession) {
          // The inserted column has no underlying file cells to restore.
          // Update only the resident intersection; evicting the whole
          // viewport made unrelated A:J data flash blank while a CT fill
          // reloaded asynchronously.
          const loaded = current.loadedRanges.get(op.sheetId)
          if (!loaded) return
          const startRow = Math.max(loaded.startRow, fill.startRow)
          const endRow = Math.min(loaded.endRow, fill.endRow)
          const startColumn = Math.max(loaded.startColumn, fill.startColumn)
          const endColumn = Math.min(loaded.endColumn, fill.endColumn)
          if (startRow > endRow || startColumn > endColumn) return
          const value = direction === 'undo' ? null : fill.value
          const matrix = Array.from({ length: endRow - startRow + 1 }, () =>
            Array.from({ length: endColumn - startColumn + 1 }, () => ({
              v: value,
              f: null,
              si: null,
            })),
          )
          journalSuppression.active = true
          try {
            active
              .getRange(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1)
              .setValues(matrix)
            // Undo may have reinstated purged pre-fill entries (a
            // clear, a copy — possibly formulas or rich text); replay
            // them with the same overlay full reloads use.
            if (direction === 'undo') {
              applyJournalOverlay(active, current.editJournal, {
                startRow,
                endRow,
                startColumn,
                endColumn,
              })
            }
          } finally {
            journalSuppression.active = false
          }
          return
        }
        current.loadedRanges.delete(op.sheetId)
        void loadVisibleRange(runtime, lazyWorkbookRef, active, setMessage)
      }
      if (targetIsInsertedThisSession) {
        pushBulkFillUndo(
          runtime,
          fill,
          purgedCells,
          ({ fill: carriedFill, purgedCells: carriedPurged, direction }) => {
            const current = lazyWorkbookRef.current
            if (!current) return false
            if (direction === 'undo') {
              removeBulkConstantFill(current.editJournal, carriedFill)
              restoreJournalCells(current.editJournal, carriedFill.sheetId, carriedPurged)
            } else {
              recordBulkConstantFill(current.editJournal, carriedFill)
            }
            refresh(direction === 'undo' ? 'undo' : 'apply')
            return true
          },
        )
      } else {
        // Existing columns need their prior values to undo. Keep the
        // in-session closure path; undo-carry intentionally truncates it
        // at Save rather than pretending the old values are recoverable.
        pushVisualUndo(runtime, {
          undo: () => {
            removeBulkConstantFill(lazyState.editJournal, fill)
            restoreJournalCells(lazyState.editJournal, fill.sheetId, purgedCells)
            refresh('undo')
          },
          redo: () => {
            recordBulkConstantFill(lazyState.editJournal, fill)
            refresh()
          },
        })
      }
      ctx.queueChartDataSync(op.sheetId, dst)
      refresh()
      return
    }
    await applyRangeInLoadedChunks(
      runtime,
      lazyWorkbookRef,
      targetSheet,
      dst,
      (chunk) => {
        const matrix: FillMatrixCell[][] = []
        for (let row = chunk.startRow; row <= chunk.endRow; row += 1) {
          const matrixRow: FillMatrixCell[] = []
          for (let column = chunk.startColumn; column <= chunk.endColumn; column += 1) {
            const sourceRowOffset = (row - dst.startRow) % sourceRows
            const sourceColumnOffset = (column - dst.startColumn) % sourceColumns
            const source = sourceCells[sourceRowOffset]?.[sourceColumnOffset]
            if (source?.formula) {
              const rowDelta = row - (src.startRow + sourceRowOffset)
              const columnDelta = column - (src.startColumn + sourceColumnOffset)
              matrixRow.push({
                v: null,
                f: offsetFormulaRefs(source.formula, rowDelta, columnDelta),
                si: null,
                ...(source.s ? { s: source.s } : {}),
              })
            } else {
              matrixRow.push({
                v: source?.value ?? null,
                f: null,
                si: null,
                ...(source?.t != null ? { t: source.t } : {}),
                ...(source?.s ? { s: source.s } : {}),
              })
            }
          }
          matrix.push(matrixRow)
        }
        targetSheet
          .getRange(
            chunk.startRow,
            chunk.startColumn,
            chunk.endRow - chunk.startRow + 1,
            chunk.endColumn - chunk.startColumn + 1,
          )
          .setValues(matrix)
      },
      setMessage,
      { neighborColumns: fillWritesFormulas },
    )
  } else if (op.op === 'copy_range') {
    // Copy one block once: the source is read chunk by chunk (loading
    // streamed regions first, so cached file values are real), then the
    // target is written the same way. Relative formula references shift
    // by the block offset, exactly like Excel paste; geometry/overlap
    // were validated at propose time, so read-then-write is safe.
    const targetSheet = sheetById(op.sheetId)
    const sourceSheet = sheetById(op.sourceSheetId ?? op.sheetId)
    const src = parseRange(op.source)
    const dst = copyTargetBounds(op)
    const rowDelta = dst.startRow - src.startRow
    const columnDelta = dst.startColumn - src.startColumn
    // v/t/s come from the raw cell model: getValues() reads the view
    // model, where the numfmt interceptor has already replaced dates
    // and formatted numbers with their display strings — copying that
    // turns a date serial into text. The display string is still kept,
    // but only for filter matching (its long-standing semantics).
    type CopyCell = {
      v: string | number | boolean | null
      t: number | null
      s: Record<string, unknown> | null
      display: string | number | boolean | null
      f: string | null
    }
    // Styles resolve to objects up front: the edit journal ignores
    // style-pool id strings, so ids would save without their number
    // format.
    const workbookStyles = runtime.univerAPI.getActiveWorkbook()?.getWorkbook().getStyles()
    const sourceCells: CopyCell[][] = []
    // Streamed unfiltered copies read the sidecar payload directly —
    // installing every source chunk into the grid and reading it back
    // dominated bulk-copy time and transient memory. Filter matching
    // needs the grid's display text, so filtered copies keep the grid
    // path; so does full-load mode (already resident, no load cost).
    const directSource =
      op.filterColumn === undefined &&
      lazyWorkbookRef.current &&
      !lazyWorkbookRef.current.formulaMode
        ? await readCopySourceDirect(
            lazyWorkbookRef,
            op.sourceSheetId ?? op.sheetId,
            src,
            setMessage,
          )
        : null
    if (directSource) {
      for (let row = 0; row < directSource.length; row += 1) {
        const rowIn = directSource[row]
        if (!rowIn) continue
        sourceCells[row] = rowIn.map((cell) => ({
          v: cell.v,
          t: cell.t,
          s: cell.s,
          display: cell.v,
          f: cell.f,
        }))
      }
    }
    if (!directSource)
      await applyRangeInLoadedChunks(
        runtime,
        lazyWorkbookRef,
        sourceSheet,
        src,
        (chunk) => {
          const chunkRange = sourceSheet.getRange(
            chunk.startRow,
            chunk.startColumn,
            chunk.endRow - chunk.startRow + 1,
            chunk.endColumn - chunk.startColumn + 1,
          )
          const values = chunkRange.getValues() as (string | number | boolean | null)[][]
          const rawCells = chunkRange.getCellDatas()
          const formulas = chunkRange.getFormulas()
          for (let row = chunk.startRow; row <= chunk.endRow; row += 1) {
            const rowCells: CopyCell[] = []
            for (let column = chunk.startColumn; column <= chunk.endColumn; column += 1) {
              const display = values[row - chunk.startRow]?.[column - chunk.startColumn] ?? null
              const raw = rawCells[row - chunk.startRow]?.[column - chunk.startColumn]
              const style = raw?.s ? (workbookStyles?.getStyleByCell(raw) ?? null) : null
              rowCells.push({
                // rich-text cells have no plain v anywhere (display
                // included); degrade to their dataStream text
                v: (raw?.v as CopyCell['v']) ?? richCellText(raw) ?? display,
                t: raw?.t ?? null,
                s: style ? (style as Record<string, unknown>) : null,
                display,
                f: formulas[row - chunk.startRow]?.[column - chunk.startColumn] || null,
              })
            }
            sourceCells[row - src.startRow] = rowCells
          }
        },
        setMessage,
        // Reading source values/formula text never needs neighbor columns.
        { neighborColumns: false },
      )
    // A filtered copy keeps only the matching source rows, compacted at
    // the target, as static values (per-row formula reference shifts
    // would be irregular under compaction, and extraction wants data).
    const filtered = op.filterColumn !== undefined
    const sourceRows = filteredCopySourceRows(op, (row, column) => {
      const cell = sourceCells[row - src.startRow]?.[column - src.startColumn]
      return matchableCellText(cell?.display ?? null)
    })
    if (sourceRows === null) {
      throw new Error(
        `copy_range filter matched no source rows — no ${op.filterColumn} cell in ${op.source} equals ` +
          `${(op.filterValues ?? []).join(', ')} (matching is trimmed and case-insensitive). ` +
          'Read the column to check the actual values.',
      )
    }
    const writeBounds = filtered
      ? { ...dst, endRow: dst.startRow + sourceRows.length - 1 }
      : { ...dst }
    if (filtered && writeBounds.endRow >= targetSheet.getMaxRows()) {
      throw new Error(
        `copy_range filter matched ${sourceRows.length} rows, but the target sheet grid has only ` +
          `${targetSheet.getMaxRows()} rows (the write would reach row ${writeBounds.endRow + 1}) — ` +
          'create the target sheet with enough rows (add_sheet rows) or insert_rows first, then retry.',
      )
    }
    // Streamed (cached-value) mode never installs `f` into the grid; the
    // harvested formula index recovers the source formulas so an
    // unfiltered copy carries them live — their referenced file cells
    // are loaded & pinned into the engine like any streamed formula
    // write. Over the shared session pin budget (or an over-expensive
    // formula, or a failed pin read) the copy falls back to the source
    // cells' current values and says so in the tool result.
    const carriedFormulas = new Map<string, string>()
    if (!filtered) {
      const lazy = lazyWorkbookRef.current
      if (lazy && !lazy.formulaMode) {
        const copySourceSheetId = op.sourceSheetId ?? op.sheetId
        for (let row = src.startRow; row <= src.endRow; row += 1) {
          for (let column = src.startColumn; column <= src.endColumn; column += 1) {
            const cell = sourceCells[row - src.startRow]?.[column - src.startColumn]
            if (!cell || cell.f !== null) continue
            // The direct read carries the file's formula text per cell;
            // the grid path recovers it from the harvested index.
            const text = directSource
              ? directSource[row - src.startRow]?.[column - src.startColumn]?.fileFormula
              : indexedFormulaText(lazy, copySourceSheetId, row, column)
            if (text) carriedFormulas.set(`${row}:${column}`, text)
          }
        }
        if (carriedFormulas.size > 0 && workbook) {
          const freezeAll = (reason: string): void => {
            notices.push(
              `copy_range ${op.source}: ${carriedFormulas.size} formula cell(s) were ` +
                `copied as their current values — ${reason}.`,
            )
            carriedFormulas.clear()
          }
          const carryPlan = carryCopyFormulasPlan(
            lazy,
            workbook,
            op.sheetId,
            targetSheet.getSheetName(),
            carriedFormulas,
            rowDelta,
            columnDelta,
          )
          if (!carryPlan.ok) {
            freezeAll(carryPlan.reason)
          } else if (
            carryPlan.needs.size > 0 &&
            !(await pinStreamedPrecedents(runtime, lazyWorkbookRef, carryPlan.needs))
          ) {
            freezeAll(
              'the cells they reference could not be loaded right now; ' +
                'retry the copy in a moment to carry them live',
            )
          }
        }
      }
    }
    // Copied formulas need real neighbor values at the target to
    // compute against; value-only copies load just the target columns.
    const copyWritesFormulas =
      carriedFormulas.size > 0 ||
      (!filtered && sourceCells.some((row) => row?.some((cell) => cell.f !== null)))
    await applyRangeInLoadedChunks(
      runtime,
      lazyWorkbookRef,
      targetSheet,
      writeBounds,
      (chunk) => {
        const matrix: {
          v: string | number | boolean | null
          f: string | null
          si: null
          t?: number
          s?: Record<string, unknown>
        }[][] = []
        for (let row = chunk.startRow; row <= chunk.endRow; row += 1) {
          const matrixRow: (typeof matrix)[number] = []
          const sourceRow = sourceRows[row - dst.startRow]
          for (let column = chunk.startColumn; column <= chunk.endColumn; column += 1) {
            const cell =
              sourceRow === undefined
                ? undefined
                : sourceCells[sourceRow - src.startRow]?.[column - dst.startColumn]
            const formulaText =
              !filtered && cell && sourceRow !== undefined
                ? (cell.f ??
                  carriedFormulas.get(
                    `${sourceRow}:${src.startColumn + (column - dst.startColumn)}`,
                  ) ??
                  null)
                : null
            if (formulaText && cell) {
              matrixRow.push({
                v: null,
                f: offsetFormulaRefs(formulaText, rowDelta, columnDelta),
                si: null,
                ...(cell.s ? { s: cell.s } : {}),
              })
            } else {
              matrixRow.push({
                v: cell?.v ?? null,
                f: null,
                si: null,
                ...(cell?.t != null ? { t: cell.t } : {}),
                ...(cell?.s ? { s: cell.s } : {}),
              })
            }
          }
          matrix.push(matrixRow)
        }
        targetSheet
          .getRange(
            chunk.startRow,
            chunk.startColumn,
            chunk.endRow - chunk.startRow + 1,
            chunk.endColumn - chunk.startColumn + 1,
          )
          .setValues(matrix)
      },
      setMessage,
      { neighborColumns: copyWritesFormulas },
    )
  } else if (op.op === 'convert_to_values') {
    // Freeze formulas into their computed values, chunk by chunk. The
    // write is a sparse object matrix (absolute row/column keys) with
    // only the formula cells, one command per chunk — non-formula
    // cells, including rich text, are never touched.
    const targetSheet = sheetById(op.sheetId)
    await applyRangeInLoadedChunks(
      runtime,
      lazyWorkbookRef,
      targetSheet,
      parseRange(op.range),
      (chunk) => {
        const chunkRange = targetSheet.getRange(
          chunk.startRow,
          chunk.startColumn,
          chunk.endRow - chunk.startRow + 1,
          chunk.endColumn - chunk.startColumn + 1,
        )
        // Raw model values: getValues() reads the view model, where the
        // numfmt interceptor rewrote formatted results into display
        // strings — freezing those would turn dates/percentages into
        // text (the cell keeps its number format, so the raw value
        // still renders identically). Formulas the engine cannot
        // compute still SHOW the file's cached result via a
        // display-only interceptor while the model holds an error —
        // freezing that error would lose the visible value, so those
        // fall back to the display text.
        const rawCells = chunkRange.getCellDatas()
        const displays = chunkRange.getValues() as (string | number | boolean | null)[][]
        const formulas = chunkRange.getFormulas()
        const updates: Record<
          number,
          Record<number, { v: string | number | boolean | null; t?: number; f: null; si: null }>
        > = {}
        let touched = 0
        for (let row = chunk.startRow; row <= chunk.endRow; row += 1) {
          for (let column = chunk.startColumn; column <= chunk.endColumn; column += 1) {
            if (!formulas[row - chunk.startRow]?.[column - chunk.startColumn]) continue
            const raw = rawCells[row - chunk.startRow]?.[column - chunk.startColumn]
            const display = displays[row - chunk.startRow]?.[column - chunk.startColumn] ?? null
            let value = (raw?.v as string | number | boolean | null | undefined) ?? null
            let keepType = true
            const engineError = typeof value === 'string' && value.startsWith('#')
            if ((value === null || engineError) && display !== null && display !== value) {
              value = display
              keepType = false
            }
            ;(updates[row] ??= {})[column] = {
              v: value,
              f: null,
              si: null,
              ...(keepType && raw?.t != null ? { t: raw.t } : {}),
            }
            touched += 1
          }
        }
        if (touched > 0) chunkRange.setValues(updates)
      },
      setMessage,
    )
  } else if (op.op === 'find_replace') {
    // Range-level replace (>MAX_EXPANDED_CELL_OPS cells): scan loaded
    // chunks and rewrite only the matching text cells with a sparse
    // object-matrix write (one command per chunk). Formula cells and
    // non-string values are skipped, matching the per-cell path.
    const targetSheet = sheetById(op.sheetId)
    const matchCase = op.matchCase ?? false
    const needle = matchCase ? op.find : op.find.toLowerCase()
    // Spaces trimmed, line breaks kept — same convention as the find
    // dialog (lazy-find.ts) so chunk replaces agree with per-cell ones.
    const trimSpaces = (s: string): string => s.replace(/^ +/g, '').replace(/ +$/g, '')
    await applyRangeInLoadedChunks(
      runtime,
      lazyWorkbookRef,
      targetSheet,
      parseRange(op.range),
      (chunk) => {
        const chunkRange = targetSheet.getRange(
          chunk.startRow,
          chunk.startColumn,
          chunk.endRow - chunk.startRow + 1,
          chunk.endColumn - chunk.startColumn + 1,
        )
        // Raw model values: the view model shows formatted numbers and
        // dates as strings, so matching on it would overwrite them with
        // replacement text (type corruption). Mirrors the per-cell path.
        // Rich-text cells have no plain v — degrade them to display
        // text, like copy_range, so they still participate.
        const rawCells = chunkRange.getCellDatas()
        const formulas = chunkRange.getFormulas()
        const updates: Record<
          number,
          Record<number, { v: string; f: null; si: null; p?: null }>
        > = {}
        let touched = 0
        for (let row = chunk.startRow; row <= chunk.endRow; row += 1) {
          for (let column = chunk.startColumn; column <= chunk.endColumn; column += 1) {
            if (formulas[row - chunk.startRow]?.[column - chunk.startColumn]) continue
            const raw = rawCells[row - chunk.startRow]?.[column - chunk.startColumn]
            const isRich = raw?.p != null && raw?.v == null
            // rich cells materialize no v anywhere (display included) —
            // their text lives in the p dataStream
            const value = raw?.v ?? (isRich ? richCellText(raw) : null)
            if (typeof value !== 'string') continue
            const haystack = matchCase ? value : value.toLowerCase()
            let next: string | null = null
            if (op.wholeCell) {
              if (trimSpaces(haystack) === needle.trim()) next = op.replace
            } else if (haystack.includes(needle)) {
              next = replaceOccurrences(value, op.find, op.replace, matchCase)
            }
            if (next === null || next === value) continue
            // rich-text matches must clear the document, or setValues
            // merges and the old rich text keeps rendering
            ;(updates[row] ??= {})[column] = isRich
              ? { v: next, f: null, si: null, p: null }
              : { v: next, f: null, si: null }
            touched += 1
          }
        }
        if (touched > 0) chunkRange.setValues(updates)
      },
      setMessage,
      // Replacing text in value cells never involves formulas.
      { neighborColumns: false },
    )
  } else if (op.op === 'sort_range') {
    // Range-level sort (>MAX_EXPANDED_CELL_OPS cells). Pass 1 loads
    // every chunk and copies the raw cell payloads out (formulas reject
    // loud — moving formula text re-targets relative references); the
    // row order comes from the same comparator as the per-cell path;
    // pass 2 reloads each chunk and rewrites the moved rows, carrying
    // each value's own t (forced-text cells must not re-infer as
    // numbers). Values move, formats stay, like the per-cell path — the
    // value-only write keeps the target cell's existing style.
    const targetSheet = sheetById(op.sheetId)
    const bounds = parseRange(op.range)
    const sortWidth = bounds.endColumn - bounds.startColumn + 1
    const sortHeight = bounds.endRow - bounds.startRow + 1
    const headerRows = (op.hasHeader ?? false) ? 1 : 0
    type SortCell = {
      v: string | number | boolean | null
      t: number | null
      rich: boolean
    }
    const sortCells: SortCell[][] = Array.from({ length: sortHeight }, () =>
      Array.from({ length: sortWidth }, (): SortCell => ({ v: null, t: null, rich: false })),
    )
    await applyRangeInLoadedChunks(
      runtime,
      lazyWorkbookRef,
      targetSheet,
      bounds,
      (chunk) => {
        const chunkRange = targetSheet.getRange(
          chunk.startRow,
          chunk.startColumn,
          chunk.endRow - chunk.startRow + 1,
          chunk.endColumn - chunk.startColumn + 1,
        )
        const rawCells = chunkRange.getCellDatas()
        const formulas = chunkRange.getFormulas()
        for (let row = chunk.startRow; row <= chunk.endRow; row += 1) {
          for (let column = chunk.startColumn; column <= chunk.endColumn; column += 1) {
            // Header cells never move, so a formula there is fine —
            // the per-cell path only reads from the first data row too.
            const isHeader = row < bounds.startRow + headerRows
            if (!isHeader && formulas[row - chunk.startRow]?.[column - chunk.startColumn]) {
              throw new Error(
                `The sort range contains a formula at ${formatAddress(row, column)} — ` +
                  'sorting would silently re-target its references. Sort values only, ' +
                  'or convert formulas to values first.',
              )
            }
            const raw = rawCells[row - chunk.startRow]?.[column - chunk.startColumn]
            // rich cells materialize no v anywhere — degrade to their
            // dataStream text, like copy_range / find_replace
            const isRich = raw?.p != null && raw?.v == null
            sortCells[row - bounds.startRow]![column - bounds.startColumn] = {
              v:
                (raw?.v as string | number | boolean | null | undefined) ??
                (isRich ? richCellText(raw) : null) ??
                null,
              t: isRich ? null : (raw?.t ?? null),
              rich: isRich,
            }
          }
        }
      },
      setMessage,
      // Sorting only moves plain values; formulas are rejected above.
      { neighborColumns: false },
    )
    const dataRows = sortCells.slice(headerRows)
    const order = computeSortedRowOrder(
      dataRows.map((row) => row.map((cell) => cell.v)),
      columnIndex(op.byColumn) - bounds.startColumn,
      op.order === 'asc',
    )
    const movedRows = order
      .map((sourceIndex, target) => ({ sourceIndex, target }))
      .filter((entry) => entry.sourceIndex !== entry.target)
    if (movedRows.length > 0) {
      await applyRangeInLoadedChunks(
        runtime,
        lazyWorkbookRef,
        targetSheet,
        bounds,
        (chunk) => {
          const updates: Record<
            number,
            Record<
              number,
              {
                v: string | number | boolean | null
                f: null
                si: null
                t: number | null
                p?: null
              }
            >
          > = {}
          let touched = 0
          for (const { sourceIndex, target } of movedRows) {
            const row = bounds.startRow + headerRows + target
            if (row < chunk.startRow || row > chunk.endRow) continue
            const sourceRow = dataRows[sourceIndex]
            const targetRow = dataRows[target]
            if (!sourceRow || !targetRow) continue
            const startColumn = Math.max(chunk.startColumn, bounds.startColumn)
            const endColumn = Math.min(chunk.endColumn, bounds.endColumn)
            for (let column = startColumn; column <= endColumn; column += 1) {
              const offset = column - bounds.startColumn
              const sourceCell = sourceRow[offset]
              const targetCell = targetRow[offset]
              if (!sourceCell || !targetCell) continue
              ;(updates[row] ??= {})[column] = {
                v: sourceCell.v,
                f: null,
                si: null,
                // Always written: the sparse patch merges into the
                // target, so an omitted t would leave the target's old
                // type on the moved value (a number landing on a
                // forced-text cell would keep t=text).
                t: sourceCell.t,
                // targets that held rich text must clear the document,
                // or setValues merges and the old rich text keeps
                // rendering
                ...(targetCell.rich ? { p: null } : {}),
              }
              touched += 1
            }
          }
          if (touched > 0) {
            targetSheet
              .getRange(
                chunk.startRow,
                chunk.startColumn,
                chunk.endRow - chunk.startRow + 1,
                chunk.endColumn - chunk.startColumn + 1,
              )
              .setValues(updates)
          }
        },
        setMessage,
        { neighborColumns: false },
      )
    }
  } else if (op.op === 'set_note') {
    const target = sheetById(op.sheetId)
    const noteRange = target.getRange(op.address)
    if (op.text === null) {
      noteRange.deleteNote()
    } else {
      const cell = parseAddress(op.address)
      noteRange.createOrUpdateNote({
        id: `note-${op.sheetId}-${cell.row}-${cell.column}`,
        row: cell.row,
        col: cell.column,
        width: 220,
        height: 90,
        note: op.text,
      })
    }
  } else {
    ctx.chartEditRef.current(
      op.chartPath,
      await buildAiChartEdit(ctx.visualContext(), state, workbook, op),
    )
  }
  run.markApplied()
}
