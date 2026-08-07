/**
 * Chart/shape/image insertion (ribbon and AI ops). Extracted from App.tsx;
 * the App component passes a VisualActionContext built fresh per call so
 * refs and state never go stale.
 */
import { columnLabel, parseAddress, parseRange } from '../domain/cell-address'
import {
  hasNumericYearAxis,
  recommendCharts,
  type ChartRecommendations,
} from '../domain/chart-recommend'
import { buildChartVisual, chartDataFromValues } from '../domain/chart-visual'
import type { InMemoryWorkbookAdapter } from '../domain/in-memory-workbook'
import { buildPivotChartData } from '../domain/pivot-chart'
import type {
  AddChartOperation,
  AddImageOperation,
  AddShapeOperation,
  EditChartOperation,
  EditShapeOperation,
} from '../domain/workbook-dsl'
import type { ChangePlan } from '../domain/workbook.types'
import type { WorkbookVisualObject } from '../shared/desktop-api'
import {
  isSheetRemoved,
  journalSize,
  recordVisualAdd,
  removeVisualAdd,
  updateVisualAdd,
} from './edit-journal'
import { t } from './i18n/locale'
import { findPivotAtSelection, type PivotActionContext } from './pivot-actions'
import { startSheetShapeDraw } from './shape-draw'
import {
  a1RangeRef,
  a1RowRangeRef,
  absRangeRef,
  pushVisualUndo,
  queueVisualInstall,
  readChartGridValues,
} from './univer-sync'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'
import {
  clearVisualSelection,
  type ChartEditData,
  type ChartVectorRead,
  type ShapeEditChanges,
} from './WorkbookVisuals'

/** The App refs/state the visual-insert actions need; built fresh per call. */
export interface VisualActionContext {
  adapterRef: { readonly current: InMemoryWorkbookAdapter }
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { current: LazyWorkbookState | null }
  visualDisposablesRef: { current: { dispose(): void }[] }
  visualInstallTimerRef: { current: ReturnType<typeof setTimeout> | null }
  chartEditRef: { current: (chartPath: string, edit: ChartEditData) => void }
  chartVectorRef: { current: (chartPath: string, range: string) => Promise<ChartVectorRead> }
  shapeEditRef: { current: (visualId: string, changes: ShapeEditChanges) => void }
  setMessage: (message: string) => void
  setRevision: (revision: number) => void
  setPreview: (plan: ChangePlan | null) => void
  setPendingEdits: (count: number) => void
  pivotContext: () => PivotActionContext
  queueDemoVisualInstall: (runtime: UniverRuntime, sheetId: string) => void
  refreshLazyVisuals: (state: LazyWorkbookState) => void
}

function queueCtxVisualInstall(
  ctx: VisualActionContext,
  runtime: UniverRuntime,
  sheetId: string,
): void {
  queueVisualInstall(
    runtime,
    ctx.lazyWorkbookRef,
    ctx.visualDisposablesRef,
    ctx.visualInstallTimerRef,
    sheetId,
    ctx.chartEditRef,
    ctx.chartVectorRef,
    ctx.shapeEditRef,
  )
}

function pushVisualAddUndo(
  ctx: VisualActionContext,
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  visual: WorkbookVisualObject,
): void {
  recordVisualAdd(state.editJournal, visual)
  pushVisualUndo(runtime, {
    undo: () => {
      removeVisualAdd(state.editJournal, visual.id)
      clearVisualSelection(visual.id)
      ctx.refreshLazyVisuals(state)
    },
    redo: () => {
      recordVisualAdd(state.editJournal, visual)
      ctx.refreshLazyVisuals(state)
    },
  })
}

/// Ribbon Insert Chart in demo mode: routes the selection through the same
/// adapter add_chart op as the AI path, so the chart lives in the snapshot
/// (undo/rebuild replay it). Values come from the snapshot — formula cells
/// carry no cached value there and read as empty.
function insertDemoChartFromSelection(
  ctx: VisualActionContext,
  runtime: UniverRuntime,
  chartKind: string,
): void {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) {
    ctx.setMessage(t('appSelectDataRangeFirst'))
    return
  }
  if (
    chartKind !== 'column' &&
    chartKind !== 'bar' &&
    chartKind !== 'line' &&
    chartKind !== 'area' &&
    chartKind !== 'pie' &&
    chartKind !== 'doughnut' &&
    chartKind !== 'scatter' &&
    chartKind !== 'radar' &&
    chartKind !== 'combo'
  ) {
    ctx.setMessage(t('appUnsupportedChartType', { kind: chartKind }))
    return
  }
  const startRow = range.getRow()
  const startColumn = range.getColumn()
  const dataRange =
    `${columnLabel(startColumn)}${startRow + 1}` +
    `:${columnLabel(startColumn + range.getWidth() - 1)}${startRow + range.getHeight()}`
  try {
    const plan = ctx.adapterRef.current.plan({
      dslVersion: 1,
      transactionId: `ribbon-chart-${crypto.randomUUID()}`,
      baseRevision: ctx.adapterRef.current.getSnapshot().revision,
      summary: `Insert ${chartKind} chart`,
      operations: [
        {
          op: 'add_chart',
          sheetId: worksheet.getSheetId(),
          chartType: chartKind,
          dataRange,
        },
      ],
    })
    const receipt = ctx.adapterRef.current.apply(plan)
    ctx.setRevision(receipt.revision)
    ctx.setPreview(null)
    ctx.queueDemoVisualInstall(runtime, worksheet.getSheetId())
    ctx.setMessage(t('appChartInsertedDemo'))
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appChartInsertFailed'))
  }
}

export async function handleInsertChart(
  ctx: VisualActionContext,
  chartKind: string,
): Promise<void> {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  if (!runtime) return
  if (!state) {
    insertDemoChartFromSelection(ctx, runtime, chartKind)
    return
  }
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) {
    ctx.setMessage(t('appSelectDataRangeFirst'))
    return
  }
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return
  const startRow = range.getRow()
  const startColumn = range.getColumn()
  const endColumn = startColumn + range.getWidth() - 1
  const endRow = startRow + range.getHeight() - 1
  // Sidecar-backed read: the Univer grid returns empty for cells outside
  // the streamed window on partially loaded workbooks.
  let values: (string | number | boolean | null | undefined)[][]
  try {
    values = await readChartGridValues(
      state,
      runtime,
      sheetId,
      `${columnLabel(startColumn)}${startRow + 1}:${columnLabel(endColumn)}${endRow + 1}`,
    )
  } catch (error) {
    ctx.setMessage(error instanceof Error ? error.message : t('appChartNeedsNumericColumn'))
    return
  }
  // Scatter reads its X values from the first data vector.
  let parsed = chartDataFromValues(
    values,
    chartKind === 'scatter' ? { numericCategoryColumn: true } : undefined,
  )
  if (!parsed) {
    ctx.setMessage(t('appChartNeedsNumericColumn'))
    return
  }
  // Keep Year/Value selections consistent with the recommendation previews:
  // the year vector is the category axis, not a plotted series.
  if (chartKind !== 'scatter' && hasNumericYearAxis(parsed)) {
    parsed = chartDataFromValues(values, { numericCategoryColumn: true }) ?? parsed
  }
  const sheetName = worksheet.getSheetName()
  const dataStartRow = startRow + (parsed.hasHeaderRow && !parsed.byRow ? 1 : 0)
  // by-row orientation: series names come from the first column, categories
  // from the first row
  const dataStartColumn = startColumn + (parsed.hasHeaderRow && parsed.byRow ? 1 : 0)
  const visual: WorkbookVisualObject = {
    id: `added-chart-${Date.now().toString(36)}-${state.editJournal.visualAdds.length + 1}`,
    sheetId,
    kind: 'chart',
    anchor: {
      fromRow: startRow,
      fromColumn: endColumn + 2,
      fromRowOffset: 0,
      fromColumnOffset: 0,
      toRow: startRow + 15,
      toColumn: endColumn + 9,
      toRowOffset: 0,
      toColumnOffset: 0,
    },
    chart: {
      chartTypes:
        chartKind === 'line'
          ? ['lineChart']
          : chartKind === 'area'
            ? ['areaChart']
            : chartKind === 'pie'
              ? ['pieChart']
              : chartKind === 'scatter'
                ? ['scatterChart']
                : chartKind === 'radar'
                  ? ['radarChart']
                  : chartKind === 'doughnut'
                    ? ['doughnutChart']
                    : chartKind === 'combo'
                      ? ['barChart', 'lineChart']
                      : ['barChart'],
      ...(chartKind === 'column' || chartKind === 'bar' || chartKind === 'combo'
        ? { barDirection: chartKind === 'bar' ? 'bar' : 'col' }
        : {}),
      title: parsed.series.length === 1 && parsed.series[0] ? parsed.series[0].name : 'Chart Title',
      series: parsed.series.map((series) => ({
        name: series.name,
        categories: parsed.categories,
        values: series.values,
        ...(parsed.byRow
          ? {
              valuesRef: a1RowRangeRef(
                sheetName,
                startRow + series.column,
                dataStartColumn,
                endColumn,
              ),
              ...(parsed.hasCategoryColumn
                ? {
                    categoriesRef: a1RowRangeRef(sheetName, startRow, dataStartColumn, endColumn),
                  }
                : {}),
            }
          : {
              valuesRef: a1RangeRef(sheetName, series.column + startColumn, dataStartRow, endRow),
              ...(parsed.hasCategoryColumn
                ? { categoriesRef: a1RangeRef(sheetName, startColumn, dataStartRow, endRow) }
                : {}),
            }),
      })),
    },
  }
  pushVisualAddUndo(ctx, runtime, state, visual)
  ctx.setPendingEdits(journalSize(state.editJournal))
  queueCtxVisualInstall(ctx, runtime, sheetId)
  ctx.setMessage(t('appChartInserted'))
}

/// PivotChart: creates a chart from the output area of the pivot under the
/// cursor. series/cat reference output-area cells, so the chart follows the
/// references after refresh write-back (chartSpace pivotSource wiring: see the
/// TODO in pivot-chart).
export async function handleInsertPivotChart(
  ctx: VisualActionContext,
  chartKind: string,
): Promise<void> {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  if (!runtime) return
  if (!state) {
    ctx.setMessage(t('appPivotChartNeedsFile'))
    return
  }
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!workbook || !worksheet) return
  const found = findPivotAtSelection(ctx.pivotContext())
  if (!found) {
    ctx.setMessage(t('appCursorNotInPivot'))
    return
  }
  if (!found.definition) {
    ctx.setMessage(t('appPivotDefNotLoaded'))
    return
  }
  if (isSheetRemoved(state.editJournal, found.sheetId)) return
  const bounds = parseRange(found.pivot.outputRef)
  let outputValues: (string | number | boolean | null | undefined)[][]
  try {
    outputValues = await readChartGridValues(state, runtime, found.sheetId, found.pivot.outputRef)
  } catch (error) {
    ctx.setMessage(error instanceof Error ? error.message : t('appPivotNoChartData'))
    return
  }
  const chartData = buildPivotChartData(
    found.definition,
    worksheet.getSheetName(),
    bounds,
    outputValues,
  )
  if (chartData.series.length === 0) {
    ctx.setMessage(t('appPivotNoChartData'))
    return
  }
  const visual: WorkbookVisualObject = {
    id: `added-pivotchart-${Date.now().toString(36)}-${state.editJournal.visualAdds.length + 1}`,
    sheetId: found.sheetId,
    kind: 'chart',
    // Anchored in the blank space to the pivot's right (same placement as
    // regular charts landing right of the selection).
    anchor: {
      fromRow: bounds.startRow,
      fromColumn: bounds.endColumn + 2,
      fromRowOffset: 0,
      fromColumnOffset: 0,
      toRow: bounds.startRow + 15,
      toColumn: bounds.endColumn + 9,
      toRowOffset: 0,
      toColumnOffset: 0,
    },
    chart: {
      chartTypes:
        chartKind === 'line'
          ? ['lineChart']
          : chartKind === 'pie'
            ? ['pieChart']
            : chartKind === 'doughnut'
              ? ['doughnutChart']
              : chartKind === 'radar'
                ? ['radarChart']
                : ['barChart'],
      ...(chartKind === 'column' || chartKind === 'bar'
        ? { barDirection: chartKind === 'bar' ? 'bar' : 'col' }
        : {}),
      title: chartData.title,
      series: chartData.series.map((series) => ({
        name: series.name,
        categories: [...series.categories],
        values: [...series.values],
        ...(series.valuesRef === undefined ? {} : { valuesRef: series.valuesRef }),
        ...(series.categoriesRef === undefined ? {} : { categoriesRef: series.categoriesRef }),
      })),
    },
  }
  pushVisualAddUndo(ctx, runtime, state, visual)
  ctx.setPendingEdits(journalSize(state.editJournal))
  queueCtxVisualInstall(ctx, runtime, found.sheetId)
  ctx.setMessage(
    chartData.truncated ? t('appPivotChartInsertedTruncated') : t('appPivotChartInserted'),
  )
}

/**
 * Excel-parity draw mode entry: arm the crosshair over the grid; the drawn
 * rectangle (or single click = default 1in square) becomes the shape's anchor.
 */
export function startShapeDraw(ctx: VisualActionContext, shapeType: string): void {
  const runtime = ctx.univerRef.current
  if (!runtime) return
  if (!ctx.lazyWorkbookRef.current) {
    ctx.setMessage(t('appShapeNeedsFile'))
    return
  }
  startSheetShapeDraw(runtime, shapeType, (anchor) => insertShapeAtAnchor(ctx, shapeType, anchor))
}

/** Insert a gallery shape at an explicit twoCellAnchor (draw-mode commit). */
export function insertShapeAtAnchor(
  ctx: VisualActionContext,
  shapeType: string,
  anchor: WorkbookVisualObject['anchor'],
): void {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  if (!runtime || !state) return
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!worksheet) return
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return
  const visual: WorkbookVisualObject = {
    id: `added-shape-${Date.now().toString(36)}-${state.editJournal.visualAdds.length + 1}`,
    sheetId,
    kind: 'shape',
    anchor,
    shapeType,
    fillColor: '#DDEBF7',
  }
  pushVisualAddUndo(ctx, runtime, state, visual)
  ctx.setPendingEdits(journalSize(state.editJournal))
  queueCtxVisualInstall(ctx, runtime, sheetId)
  ctx.setMessage(t('appShapeInserted'))
}

export function handleInsertShape(
  ctx: VisualActionContext,
  shapeType: string,
  isTextBox: boolean,
): void {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  if (!runtime) return
  if (!state) {
    ctx.setMessage(t('appShapeNeedsFile'))
    return
  }
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) {
    ctx.setMessage(t('appSelectCellFirst'))
    return
  }
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return
  const row = range.getRow()
  const column = range.getColumn()
  const visual: WorkbookVisualObject = {
    id: `added-shape-${Date.now().toString(36)}-${state.editJournal.visualAdds.length + 1}`,
    sheetId,
    kind: 'shape',
    anchor: {
      fromRow: row,
      fromColumn: column,
      fromRowOffset: 0,
      fromColumnOffset: 0,
      toRow: row + (isTextBox ? 4 : 6),
      toColumn: column + (isTextBox ? 4 : 3),
      toRowOffset: 0,
      toColumnOffset: 0,
    },
    shapeType,
    ...(isTextBox
      ? { name: 'TextBox', fillColor: '#FFFFFF', text: 'Text' }
      : { fillColor: '#DDEBF7' }),
  }
  pushVisualAddUndo(ctx, runtime, state, visual)
  ctx.setPendingEdits(journalSize(state.editJournal))
  queueCtxVisualInstall(ctx, runtime, sheetId)
  ctx.setMessage(isTextBox ? t('appTextBoxInserted') : t('appShapeInserted'))
}

/// AI edit_chart → renderer edit data. Series ranges are read from the
/// sheet here so the file cache and the on-screen render update together.
export async function buildAiChartEdit(
  ctx: VisualActionContext,
  state: LazyWorkbookState,
  workbook: ReturnType<UniverRuntime['univerAPI']['getActiveWorkbook']>,
  op: EditChartOperation,
): Promise<ChartEditData> {
  const edit: ChartEditData = {
    ...(op.title !== undefined ? { title: op.title } : {}),
    ...(op.chartType !== undefined ? { chartType: op.chartType } : {}),
    ...(op.seriesColors ? { seriesColors: op.seriesColors } : {}),
    ...(op.legend !== undefined ? { legend: op.legend } : {}),
    ...(op.dataLabels !== undefined ? { dataLabels: op.dataLabels } : {}),
    ...(op.grouping !== undefined ? { grouping: op.grouping } : {}),
    ...(op.axisTitles !== undefined ? { axisTitles: op.axisTitles } : {}),
  }
  if (!op.seriesData || op.seriesData.length === 0) return edit
  const runtime = ctx.univerRef.current
  const chartSheetId = [...state.file.visuals, ...state.editJournal.visualAdds].find(
    (candidate) => candidate.chartPath === op.chartPath || candidate.id === op.chartPath,
  )?.sheetId
  const series = []
  for (const entry of op.seriesData) {
    const sheetId = entry.sheetId ?? chartSheetId
    const target = sheetId === undefined ? null : workbook?.getSheetBySheetId(sheetId)
    if (!target || !sheetId || !runtime) {
      throw new Error(`Unknown sheet for the chart data: ${sheetId ?? '(none)'}`)
    }
    // Sidecar-backed read: streamed-out cells read as empty from the grid.
    const vector = async (
      range: string,
    ): Promise<(string | number | boolean | null | undefined)[]> =>
      (await readChartGridValues(state, runtime, sheetId, range)).flat()
    series.push({
      index: entry.index,
      ...(entry.name === undefined ? {} : { name: entry.name }),
      ...(entry.valuesRange === undefined
        ? {}
        : {
            values: (await vector(entry.valuesRange)).map((value) => {
              const numeric = typeof value === 'number' ? value : Number(String(value ?? '').trim())
              return Number.isFinite(numeric) ? numeric : 0
            }),
            valuesRef: absRangeRef(target.getSheetName(), entry.valuesRange),
          }),
      ...(entry.categoriesRange === undefined
        ? {}
        : {
            categories: (await vector(entry.categoriesRange)).map((value) => String(value ?? '')),
            categoriesRef: absRangeRef(target.getSheetName(), entry.categoriesRange),
          }),
    })
  }
  return { ...edit, series }
}

/// AI add_image: same visual pipeline as the picker-based insert, sized
/// from the measured natural dimensions.
export function insertAiImageVisual(
  ctx: VisualActionContext,
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: AddImageOperation,
  image: { dataUrl: string; mediaType: string; width: number; height: number },
): void {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook?.getSheetBySheetId(op.sheetId)) throw new Error(`Unknown sheet: ${op.sheetId}`)
  // ~80px per column, ~22px per row; scale down to a ≤480px-wide frame
  // (same sizing as the picker-based insert).
  const scale = Math.min(1, 480 / Math.max(1, image.width))
  const columns = Math.min(16, Math.max(2, Math.round((image.width * scale) / 80)))
  const rows = Math.min(40, Math.max(2, Math.round((image.height * scale) / 22)))
  const base = parseAddress(op.anchorCell)
  const visual: WorkbookVisualObject = {
    id: `added-image-${Date.now().toString(36)}-${state.editJournal.visualAdds.length + 1}`,
    sheetId: op.sheetId,
    kind: 'image',
    anchor: {
      fromRow: base.row,
      fromColumn: base.column,
      fromRowOffset: 0,
      fromColumnOffset: 0,
      toRow: base.row + rows,
      toColumn: base.column + columns,
      toRowOffset: 0,
      toColumnOffset: 0,
    },
    mediaType: image.mediaType,
    mediaDataUrl: image.dataUrl,
    name: op.path.split('/').pop() ?? 'image',
  }
  pushVisualAddUndo(ctx, runtime, state, visual)
  queueCtxVisualInstall(ctx, runtime, op.sheetId)
}

/// AI edit_shape: in-place journal update of a session-added shape,
/// preserving the frame size on a move.
export function applyAiShapeEdit(
  ctx: VisualActionContext,
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: EditShapeOperation,
): void {
  const visual = state.editJournal.visualAdds.find((candidate) => candidate.id === op.visualId)
  if (!visual || visual.kind !== 'shape') {
    throw new Error(t('appShapeNotEditable', { id: op.visualId }))
  }
  let anchor: WorkbookVisualObject['anchor'] | undefined
  if (op.anchorCell !== undefined) {
    const base = parseAddress(op.anchorCell)
    anchor = {
      ...visual.anchor,
      fromRow: base.row,
      fromColumn: base.column,
      toRow: base.row + (visual.anchor.toRow - visual.anchor.fromRow),
      toColumn: base.column + (visual.anchor.toColumn - visual.anchor.fromColumn),
    }
  }
  updateVisualAdd(state.editJournal, op.visualId, {
    ...(anchor === undefined ? {} : { anchor }),
    ...(op.text === undefined ? {} : { text: op.text }),
    ...(op.fillColor === undefined ? {} : { fillColor: op.fillColor }),
  })
  queueCtxVisualInstall(ctx, runtime, visual.sheetId)
}

/// AI add_chart: same visual-add pipeline as the ribbon's Insert Chart,
/// fed an explicit data range instead of the selection. Values come from
/// the sidecar-mapped read, so ranges outside the streamed viewport work.
export async function insertAiChartVisual(
  ctx: VisualActionContext,
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: AddChartOperation,
): Promise<void> {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const target = workbook?.getSheetBySheetId(op.sheetId)
  if (!workbook || !target) throw new Error(`Unknown sheet: ${op.sheetId}`)
  const values = await readChartGridValues(state, runtime, op.sheetId, op.dataRange)
  const visual: WorkbookVisualObject = buildChartVisual({
    id: `added-chart-${Date.now().toString(36)}-${state.editJournal.visualAdds.length + 1}`,
    sheetId: op.sheetId,
    sheetName: target.getSheetName(),
    chartType: op.chartType,
    dataRange: op.dataRange,
    title: op.title,
    anchorCell: op.anchorCell,
    values,
  })
  pushVisualAddUndo(ctx, runtime, state, visual)
  queueCtxVisualInstall(ctx, runtime, op.sheetId)
}

export function insertAiShapeVisual(
  ctx: VisualActionContext,
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: AddShapeOperation,
): void {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook?.getSheetBySheetId(op.sheetId)) throw new Error(`Unknown sheet: ${op.sheetId}`)
  const isTextBox = op.shapeType === 'textbox'
  const base = parseAddress(op.anchorCell)
  const visual: WorkbookVisualObject = {
    id: `added-shape-${Date.now().toString(36)}-${state.editJournal.visualAdds.length + 1}`,
    sheetId: op.sheetId,
    kind: 'shape',
    anchor: {
      fromRow: base.row,
      fromColumn: base.column,
      fromRowOffset: 0,
      fromColumnOffset: 0,
      toRow: base.row + (isTextBox ? 4 : 6),
      toColumn: base.column + (isTextBox ? 4 : 3),
      toRowOffset: 0,
      toColumnOffset: 0,
    },
    shapeType: isTextBox ? 'rect' : op.shapeType,
    ...(isTextBox
      ? { name: 'TextBox', fillColor: op.fillColor ?? '#FFFFFF', text: op.text ?? 'Text' }
      : {
          fillColor: op.fillColor ?? '#DDEBF7',
          ...(op.text === undefined ? {} : { text: op.text }),
        }),
  }
  pushVisualAddUndo(ctx, runtime, state, visual)
  queueCtxVisualInstall(ctx, runtime, op.sheetId)
}

export function handleInsertPicture(ctx: VisualActionContext): void {
  if (!ctx.univerRef.current) return
  if (!ctx.lazyWorkbookRef.current) {
    ctx.setMessage(t('appPictureNeedsFile'))
    return
  }
  const input = document.createElement('input')
  input.type = 'file'
  input.accept = 'image/png,image/jpeg,image/gif'
  input.onchange = () => {
    const file = input.files?.[0]
    if (!file) return
    if (file.size > 20 * 1024 * 1024) {
      ctx.setMessage(t('appPictureTooLarge'))
      return
    }
    const mediaType = file.type === 'image/jpg' ? 'image/jpeg' : file.type
    if (!['image/png', 'image/jpeg', 'image/gif'].includes(mediaType)) {
      ctx.setMessage(t('appPictureBadType'))
      return
    }
    const reader = new FileReader()
    reader.onload = () => {
      const dataUrl = typeof reader.result === 'string' ? reader.result : null
      if (!dataUrl) return
      const image = new Image()
      image.onload = () =>
        insertPictureVisual(
          ctx,
          dataUrl,
          mediaType,
          file.name,
          image.naturalWidth,
          image.naturalHeight,
        )
      image.onerror = () => insertPictureVisual(ctx, dataUrl, mediaType, file.name, 480, 320)
      image.src = dataUrl
    }
    reader.readAsDataURL(file)
  }
  input.click()
}

/// Reads the selection the same way handleInsertChart does and ranks chart
/// kinds for it; null (with a status message) when there is nothing to chart.
export async function handleRecommendedCharts(
  ctx: VisualActionContext,
): Promise<ChartRecommendations | null> {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  if (!runtime || !state) {
    ctx.setMessage(t('appSelectDataRangeFirst'))
    return null
  }
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) {
    ctx.setMessage(t('appSelectDataRangeFirst'))
    return null
  }
  const startRow = range.getRow()
  const startColumn = range.getColumn()
  const endColumn = startColumn + range.getWidth() - 1
  const endRow = startRow + range.getHeight() - 1
  let values: (string | number | boolean | null | undefined)[][]
  try {
    values = await readChartGridValues(
      state,
      runtime,
      worksheet.getSheetId(),
      `${columnLabel(startColumn)}${startRow + 1}:${columnLabel(endColumn)}${endRow + 1}`,
    )
  } catch (error) {
    ctx.setMessage(error instanceof Error ? error.message : t('appChartNeedsNumericColumn'))
    return null
  }
  const recommendations = recommendCharts(values)
  if (!recommendations) {
    ctx.setMessage(t('appChartNeedsNumericColumn'))
    return null
  }
  return recommendations
}

export function handleInsertScreenshot(
  ctx: VisualActionContext,
  dataUrl: string,
  width: number,
  height: number,
): void {
  insertPictureVisual(ctx, dataUrl, 'image/png', 'Screenshot.png', width, height)
}

export function handleInsertEquation(
  ctx: VisualActionContext,
  dataUrl: string,
  width: number,
  height: number,
): void {
  insertPictureVisual(ctx, dataUrl, 'image/png', 'Equation.png', width, height)
}

export function handleInsertIcon(
  ctx: VisualActionContext,
  dataUrl: string,
  size: number,
  name: string,
): void {
  insertPictureVisual(ctx, dataUrl, 'image/png', `${name.replace(/\s+/g, '-')}.png`, size, size)
}

function insertPictureVisual(
  ctx: VisualActionContext,
  dataUrl: string,
  mediaType: string,
  fileName: string,
  naturalWidth: number,
  naturalHeight: number,
): void {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  if (!runtime || !state) return
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!workbook || !worksheet || !range) {
    ctx.setMessage(t('appSelectCellFirst'))
    return
  }
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return
  // ~80px per column, ~22px per row; scale down to a ≤480px-wide frame.
  const scale = Math.min(1, 480 / Math.max(1, naturalWidth))
  const columns = Math.min(16, Math.max(2, Math.round((naturalWidth * scale) / 80)))
  const rows = Math.min(40, Math.max(2, Math.round((naturalHeight * scale) / 22)))
  const row = range.getRow()
  const column = range.getColumn()
  const visual: WorkbookVisualObject = {
    id: `added-image-${Date.now().toString(36)}-${state.editJournal.visualAdds.length + 1}`,
    sheetId,
    kind: 'image',
    anchor: {
      fromRow: row,
      fromColumn: column,
      fromRowOffset: 0,
      fromColumnOffset: 0,
      toRow: row + rows,
      toColumn: column + columns,
      toRowOffset: 0,
      toColumnOffset: 0,
    },
    mediaType,
    mediaDataUrl: dataUrl,
    name: fileName,
  }
  pushVisualAddUndo(ctx, runtime, state, visual)
  ctx.setPendingEdits(journalSize(state.editJournal))
  queueCtxVisualInstall(ctx, runtime, sheetId)
  ctx.setMessage(t('appPictureInserted'))
}
