/**
 * Chart/shape edit handlers (the chartEditRef/chartVectorRef/shapeEditRef
 * implementations) and the live chart↔data sync. Extracted from App.tsx;
 * the App component passes a VisualSyncContext built fresh per call so refs
 * and state never go stale.
 */
import { parseRange } from '../domain/cell-address'
import {
  applyChartStateEdit,
  refIntersects,
  splitSheetRef,
  type CellBounds,
  type ChartVisualState,
  type SheetVisual,
} from '../domain/chart-visual'
import type { InMemoryWorkbookAdapter } from '../domain/in-memory-workbook'
import type { WorkbookVisualObject } from '../shared/desktop-api'
import { recordChartEdit, recordVisualEdit, removeVisualAdd, updateVisualAdd } from './edit-journal'
import { t } from './i18n/locale'
import {
  captureVisualJournal,
  pushVisualUndo,
  readChartGridValues,
  readChartRangeVector,
  readDemoChartRangeVector,
  restoreVisualJournal,
} from './univer-sync'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'
import {
  clearVisualSelection,
  type ChartEditData,
  type ChartVectorRead,
  type ShapeEditChanges,
} from './WorkbookVisuals'

/** The App refs/state the visual edit/sync handlers need; built fresh per call. */
export interface VisualSyncContext {
  adapterRef: { readonly current: InMemoryWorkbookAdapter }
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  chartSyncRef: {
    readonly current: {
      timer: ReturnType<typeof setTimeout> | null
      dirty: Map<string, CellBounds>
    }
  }
  setMessage: (message: string) => void
  refreshLazyVisuals: (state: LazyWorkbookState) => void
  refreshDemoVisuals: () => void
}

/// Live chart↔data sync: cell edits touching a series' `c:f` reference
/// refresh that series' cache after a debounce. Silent by design — no undo
/// step (undoing the cell edit re-syncs the chart back), no message.
export function queueChartDataSync(
  ctx: VisualSyncContext,
  sheetId: string,
  bounds: CellBounds,
): void {
  const previous = ctx.chartSyncRef.current.dirty.get(sheetId)
  ctx.chartSyncRef.current.dirty.set(
    sheetId,
    previous
      ? {
          startRow: Math.min(previous.startRow, bounds.startRow),
          endRow: Math.max(previous.endRow, bounds.endRow),
          startColumn: Math.min(previous.startColumn, bounds.startColumn),
          endColumn: Math.max(previous.endColumn, bounds.endColumn),
        }
      : bounds,
  )
  if (ctx.chartSyncRef.current.timer) clearTimeout(ctx.chartSyncRef.current.timer)
  ctx.chartSyncRef.current.timer = setTimeout(() => {
    ctx.chartSyncRef.current.timer = null
    void runChartDataSync(ctx)
  }, 500)
}

async function runChartDataSync(ctx: VisualSyncContext): Promise<void> {
  const runtime = ctx.univerRef.current
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!runtime || !workbook) return
  const dirty = new Map(ctx.chartSyncRef.current.dirty)
  ctx.chartSyncRef.current.dirty.clear()
  if (dirty.size === 0) return
  const state = ctx.lazyWorkbookRef.current
  const sheetNames = new Map(
    workbook.getSheets().map((sheet) => [sheet.getSheetId(), sheet.getSheetName()] as const),
  )
  const sheetIdByName = new Map([...sheetNames].map(([id, name]) => [name, id]))
  const touched = (ref: string | undefined): boolean =>
    ref !== undefined &&
    [...dirty].some(([sheetId, bounds]) => {
      const name = sheetNames.get(sheetId)
      return name !== undefined && refIntersects(ref, name, bounds)
    })
  const readVector = async (
    ref: string,
  ): Promise<(string | number | boolean | null | undefined)[] | null> => {
    const split = splitSheetRef(ref)
    const sheetId = split ? sheetIdByName.get(split.sheetName) : undefined
    if (!split || sheetId === undefined) return null
    try {
      if (state) return (await readChartGridValues(state, runtime, sheetId, split.range)).flat()
      const target = workbook.getSheetBySheetId(sheetId)
      if (!target) return null
      return (
        target.getRange(split.range).getValues() as (
          string | number | boolean | null | undefined
        )[][]
      ).flat()
    } catch {
      return null
    }
  }
  const charts: { editKey: string; visualId: string; chart: ChartVisualState }[] = state
    ? [...state.file.visuals, ...state.editJournal.visualAdds]
        .filter(
          (visual) =>
            visual.kind === 'chart' &&
            visual.chart &&
            !state.editJournal.visualEdits.get(visual.id)?.remove,
        )
        .map((visual) => ({
          editKey: visual.chartPath ?? visual.id,
          visualId: visual.id,
          chart: applyChartStateEdit(
            visual.chart as ChartVisualState,
            visual.chartPath ? state.editJournal.chartEdits.get(visual.chartPath) : undefined,
          ),
        }))
    : ctx.adapterRef.current.getSnapshot().sheets.flatMap((sheet) =>
        (sheet.visuals ?? []).map((visual) => ({
          editKey: visual.id,
          visualId: visual.id,
          chart: visual.chart,
        })),
      )
  let changed = false
  for (const { editKey, visualId, chart } of charts) {
    const seriesEdits: NonNullable<ChartEditData['series']>[number][] = []
    for (const [index, series] of chart.series.entries()) {
      const wantValues = touched(series.valuesRef)
      const wantCategories = touched(series.categoriesRef)
      if (!wantValues && !wantCategories) continue
      const entry: NonNullable<ChartEditData['series']>[number] = { index }
      if (wantValues && series.valuesRef) {
        const vector = await readVector(series.valuesRef)
        if (vector) {
          const values = vector.map((value) => {
            const numeric = typeof value === 'number' ? value : Number(String(value ?? '').trim())
            return Number.isFinite(numeric) ? numeric : 0
          })
          if (JSON.stringify(values) !== JSON.stringify(series.values)) entry.values = values
        }
      }
      if (wantCategories && series.categoriesRef) {
        const vector = await readVector(series.categoriesRef)
        if (vector) {
          const categories = vector.map((value) => String(value ?? '').slice(0, 255))
          if (JSON.stringify(categories) !== JSON.stringify(series.categories)) {
            entry.categories = categories
          }
        }
      }
      if (entry.values !== undefined || entry.categories !== undefined) seriesEdits.push(entry)
    }
    if (seriesEdits.length === 0) continue
    changed = true
    const edit: ChartEditData = { series: seriesEdits }
    if (!state) {
      const before = ctx.adapterRef.current.findVisual(visualId)
      if (before) {
        ctx.adapterRef.current.replaceVisual(visualId, {
          ...before,
          chart: applyChartStateEdit(before.chart, edit),
        })
      }
    } else if (editKey.startsWith('xl/')) {
      recordChartEdit(state.editJournal, editKey, edit)
    } else {
      const adds = state.editJournal.visualAdds
      const at = adds.findIndex((candidate) => candidate.id === visualId)
      const current = adds[at]
      if (current?.chart) {
        adds[at] = {
          ...current,
          chart: applyChartStateEdit(current.chart, edit) as WorkbookVisualObject['chart'],
        }
      }
    }
  }
  if (!changed) return
  if (state) ctx.refreshLazyVisuals(state)
  else ctx.refreshDemoVisuals()
}

/// The silent chart↔data sync records no undo step, so an undo/redo that
/// restores a whole journal snapshot may wipe synced series data. Queueing
/// the chart's own references re-converges it onto the live cell values.
export function queueChartRefResync(
  ctx: VisualSyncContext,
  state: LazyWorkbookState,
  editKey: string,
): void {
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  if (!workbook) return
  const sheetIdByName = new Map(
    workbook.getSheets().map((sheet) => [sheet.getSheetName(), sheet.getSheetId()] as const),
  )
  const visual = [...state.file.visuals, ...state.editJournal.visualAdds].find(
    (candidate) => candidate.chartPath === editKey || candidate.id === editKey,
  )
  for (const series of visual?.chart?.series ?? []) {
    for (const ref of [series.valuesRef, series.categoriesRef]) {
      if (ref === undefined) continue
      const split = splitSheetRef(ref)
      const sheetId = split ? sheetIdByName.get(split.sheetName) : undefined
      if (!split || sheetId === undefined) continue
      try {
        queueChartDataSync(ctx, sheetId, parseRange(split.range))
      } catch {
        // Unparseable refs never sync, so there is nothing to restore.
      }
    }
  }
}

/// Demo charts bake edits straight into the snapshot visual. Deliberately
/// no adapter history entry: the Univer undo stack owns interactive ops,
/// while adapter history keeps owning AI plan applies.
function applyDemoVisualChange(
  ctx: VisualSyncContext,
  visualId: string,
  mutate: (before: SheetVisual) => SheetVisual | null,
): boolean {
  const runtime = ctx.univerRef.current
  if (!runtime) return false
  const before = ctx.adapterRef.current.findVisual(visualId)
  if (!before) return false
  const after = mutate(before)
  ctx.adapterRef.current.replaceVisual(visualId, after)
  ctx.refreshDemoVisuals()
  pushVisualUndo(runtime, {
    undo: () => {
      ctx.adapterRef.current.replaceVisual(visualId, before)
      ctx.refreshDemoVisuals()
    },
    redo: () => {
      ctx.adapterRef.current.replaceVisual(visualId, after)
      ctx.refreshDemoVisuals()
    },
  })
  return true
}

/** The chartEditRef implementation. */
export function applyChartEdit(ctx: VisualSyncContext, editKey: string, edit: ChartEditData): void {
  const state = ctx.lazyWorkbookRef.current
  const runtime = ctx.univerRef.current
  if (!runtime) return
  if (!state) {
    const applied = applyDemoVisualChange(ctx, editKey, (before) =>
      before.chart
        ? { ...before, chart: applyChartStateEdit(before.chart, edit) as SheetVisual['chart'] }
        : before,
    )
    ctx.setMessage(applied ? t('appChartUpdated') : t('appChartNotEditable'))
    return
  }
  if (editKey.startsWith('xl/')) {
    // File chart: the edit journals against its chart part.
    const before = state.editJournal.chartEdits.get(editKey)
    recordChartEdit(state.editJournal, editKey, edit)
    const after = state.editJournal.chartEdits.get(editKey)
    pushVisualUndo(runtime, {
      undo: () => {
        if (before === undefined) state.editJournal.chartEdits.delete(editKey)
        else state.editJournal.chartEdits.set(editKey, before)
        ctx.refreshLazyVisuals(state)
        queueChartRefResync(ctx, state, editKey)
      },
      redo: () => {
        if (after === undefined) state.editJournal.chartEdits.delete(editKey)
        else state.editJournal.chartEdits.set(editKey, after)
        ctx.refreshLazyVisuals(state)
        queueChartRefResync(ctx, state, editKey)
      },
    })
  } else {
    // Session-added chart: bake the edit into the journaled visual, so the
    // save writes the final object (no chart part exists yet).
    const adds = state.editJournal.visualAdds
    const index = adds.findIndex((candidate) => candidate.id === editKey)
    const before = adds[index]
    if (!before?.chart) {
      ctx.setMessage(t('appChartNotEditable'))
      return
    }
    const after = {
      ...before,
      chart: applyChartStateEdit(before.chart, edit) as WorkbookVisualObject['chart'],
    }
    adds[index] = after
    pushVisualUndo(runtime, {
      undo: () => {
        const at = state.editJournal.visualAdds.findIndex((candidate) => candidate.id === editKey)
        if (at >= 0) state.editJournal.visualAdds[at] = before
        ctx.refreshLazyVisuals(state)
        queueChartRefResync(ctx, state, editKey)
      },
      redo: () => {
        const at = state.editJournal.visualAdds.findIndex((candidate) => candidate.id === editKey)
        if (at >= 0) state.editJournal.visualAdds[at] = after
        ctx.refreshLazyVisuals(state)
        queueChartRefResync(ctx, state, editKey)
      },
    })
  }
  ctx.setMessage(t('appChartEditRecorded'))
  ctx.refreshLazyVisuals(state)
}

/** The chartVectorRef implementation. */
export function readChartVector(
  ctx: VisualSyncContext,
  editKey: string,
  rangeText: string,
): Promise<ChartVectorRead> {
  const state = ctx.lazyWorkbookRef.current
  const runtime = ctx.univerRef.current
  if (!runtime) return Promise.reject(new Error('Workbook not ready.'))
  if (!state) return readDemoChartRangeVector(runtime, ctx.adapterRef.current, editKey, rangeText)
  return readChartRangeVector(state, runtime, editKey, rangeText)
}

/** The shapeEditRef implementation. */
export function applyShapeEdit(
  ctx: VisualSyncContext,
  visualId: string,
  changes: ShapeEditChanges,
): void {
  const state = ctx.lazyWorkbookRef.current
  const runtime = ctx.univerRef.current
  if (!runtime) return
  if (!state) {
    const applied = applyDemoVisualChange(ctx, visualId, (before) =>
      changes.remove
        ? null
        : {
            ...before,
            ...(changes.anchor ? { anchor: changes.anchor } : {}),
            ...(changes.frameSize
              ? { frameWidth: changes.frameSize.width, frameHeight: changes.frameSize.height }
              : {}),
          },
    )
    if (changes.remove) clearVisualSelection(visualId)
    ctx.setMessage(
      applied
        ? changes.remove
          ? t('appChartDeleted')
          : t('appChartMoved')
        : t('appVisualNotEditable'),
    )
    return
  }
  const chartPath = state.file.visuals.find((candidate) => candidate.id === visualId)?.chartPath
  const before = captureVisualJournal(state, visualId, chartPath)
  if (changes.remove) {
    // Session visuals just drop out of the journal; file visuals record a
    // surgical anchor removal against their drawing part.
    if (!removeVisualAdd(state.editJournal, visualId)) {
      const visual = state.file.visuals.find((candidate) => candidate.id === visualId)
      if (!visual || !recordVisualEdit(state.editJournal, visual, { remove: true })) {
        ctx.setMessage(t('appVisualNoDelete'))
        return
      }
      // Pending edits on a deleted chart would patch a removed part.
      if (visual.chartPath !== undefined) {
        state.editJournal.chartEdits.delete(visual.chartPath)
      }
    }
    clearVisualSelection(visualId)
  } else if (!updateVisualAdd(state.editJournal, visualId, changes)) {
    const visual = state.file.visuals.find((candidate) => candidate.id === visualId)
    if (changes.text !== undefined || changes.anchor === undefined) {
      ctx.setMessage(t('appFileVisualMoveDeleteOnly'))
      return
    }
    if (
      !visual ||
      !recordVisualEdit(state.editJournal, visual, {
        anchor: changes.anchor,
        ...(changes.frameSize ? { frameSize: changes.frameSize } : {}),
      })
    ) {
      ctx.setMessage(t('appVisualNoMove'))
      return
    }
  }
  const after = captureVisualJournal(state, visualId, chartPath)
  pushVisualUndo(runtime, {
    undo: () => {
      restoreVisualJournal(state, visualId, chartPath, before)
      clearVisualSelection(visualId)
      ctx.refreshLazyVisuals(state)
      queueChartRefResync(ctx, state, visualId)
    },
    redo: () => {
      restoreVisualJournal(state, visualId, chartPath, after)
      clearVisualSelection(visualId)
      ctx.refreshLazyVisuals(state)
      queueChartRefResync(ctx, state, visualId)
    },
  })
  ctx.setMessage(
    changes.remove
      ? t('appVisualDeleted')
      : changes.text !== undefined
        ? t('appShapeTextUpdated')
        : t('appShapeMoved'),
  )
  ctx.refreshLazyVisuals(state)
}
