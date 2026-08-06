import { useEffect, useRef, useState, useSyncExternalStore } from 'react'
import { createPortal } from 'react-dom'
import { shapePreviewPath } from '@genoffice/ui'

import type { createUniver } from './create-univer'

import {
  applyChartStateEdit,
  chartCategoryFormat,
  chartSupportsDataLabels,
  formatCategoryLabel,
  scatterAxisBounds,
  splitSheetRef,
} from '../domain/chart-visual'
import { parseAddress } from '../domain/cell-address'
import { t } from './i18n/locale'
import type { WorkbookChartEdit, WorkbookFile, WorkbookVisualObject } from '../shared/desktop-api'

type UniverRuntime = ReturnType<typeof createUniver>
type ActiveWorkbook = NonNullable<ReturnType<UniverRuntime['univerAPI']['getActiveWorkbook']>>
type UniverWorksheet = NonNullable<ReturnType<ActiveWorkbook['getSheetBySheetId']>>

interface Disposable {
  dispose(): void
}

/// What installWorkbookVisuals actually needs from a workbook: the demo
/// workbook satisfies this with its snapshot visuals and a stable fake
/// sessionId; lazy mode passes the full WorkbookFile.
export interface VisualHost {
  readonly sessionId: WorkbookFile['sessionId']
  readonly visuals: readonly WorkbookVisualObject[]
}

export type ChartEditData = Omit<WorkbookChartEdit, 'chartPath'>

export interface ChartVectorRead {
  readonly vector: readonly (string | number | boolean | null | undefined)[]
  readonly ref: string
}

/// Chart edits key on the chart part path for file charts and on the visual
/// id for charts that only exist in memory (session adds, demo workbook);
/// the app-side handler routes on the key's shape.
export interface ChartEditing {
  readonly edits: ReadonlyMap<string, ChartEditData>
  readonly onEdit: (editKey: string, edit: ChartEditData) => void
  /// Reads a single-row/column range so series data-range edits can bake
  /// current sheet values into the chart cache. Throws on invalid ranges.
  readonly readVector?: (editKey: string, range: string) => Promise<ChartVectorRead>
}

export interface ShapeEditChanges {
  readonly anchor?: WorkbookVisualObject['anchor']
  readonly text?: string
  readonly remove?: true
}

export interface ShapeEditing {
  readonly onEdit: (visualId: string, changes: ShapeEditChanges) => void
}

/// Visuals created this session are fully editable in place: their whole
/// object lives in the journal, so anchor/text changes save for free.
export function isEditableShape(visual: WorkbookVisualObject): boolean {
  return visual.kind === 'shape' && visual.id.startsWith('added-shape-')
}

/// Images and shapes already in the file move/delete through a surgical
/// anchor edit, located by the sidecar's (drawingPath, drawingIndex) pair.
/// Charts keep their own editor; deletion routes through the same edit.
export function isEditableFileVisual(visual: WorkbookVisualObject): boolean {
  return (
    (visual.kind === 'image' || visual.kind === 'shape') &&
    !visual.id.startsWith('added-') &&
    visual.drawingPath !== undefined &&
    visual.drawingIndex !== undefined
  )
}

/// Charts move through the same anchor edit as images/shapes and delete
/// through a cascading one (the save also drops the chart part, rel, and
/// content-type override); session-added and demo charts live in memory.
function isEditableChart(visual: WorkbookVisualObject): boolean {
  if (visual.kind !== 'chart') return false
  if (visual.id.startsWith('added-') || visual.id.startsWith('demo-')) return true
  return visual.drawingPath !== undefined && visual.drawingIndex !== undefined
}

const chartColors = [
  '#4472c4',
  '#ed7d31',
  '#a5a5a5',
  '#ffc000',
  '#5b9bd5',
  '#70ad47',
  '#264478',
  '#9e480e',
  '#636363',
  '#997300',
]

type SeriesLike = {
  readonly color?: string | undefined
  readonly pointColors?: readonly { readonly index: number; readonly color: string }[] | undefined
}

function seriesColor(series: SeriesLike | undefined, index: number): string {
  return series?.color ?? chartColors[index % chartColors.length] ?? '#4472c4'
}

export function installWorkbookVisuals(
  runtime: UniverRuntime,
  file: VisualHost,
  sheetId: string,
  chartEditing?: ChartEditing,
  shapeEditing?: ShapeEditing,
): Disposable[] {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) return []
  const disposables: Disposable[] = []
  for (const visual of file.visuals.filter((candidate) => candidate.sheetId === sheetId)) {
    const worksheet = workbook.getSheetBySheetId(visual.sheetId)
    if (!worksheet) continue
    const componentKey = `xlsx-${file.sessionId}-${visual.id}`
    const editable =
      isEditableShape(visual) || isEditableFileVisual(visual) || isEditableChart(visual)
    const component =
      shapeEditing && editable
        ? () => (
            <EditableShapeVisual
              file={file}
              visual={visual}
              worksheet={worksheet}
              allowText={isEditableShape(visual)}
              chartEditing={chartEditing}
              onEdit={shapeEditing.onEdit}
            />
          )
        : () => <WorkbookVisual file={file} visual={visual} chartEditing={chartEditing} />
    disposables.push(runtime.univerAPI.registerComponent(componentKey, component))
    // Lazy grids are sized to the data, but session-added visuals anchor
    // beyond it (default: two columns right of the data) — grow the grid so
    // the float keeps its frame instead of being clamped into a sliver.
    if (visual.anchor.toRow >= worksheet.getMaxRows()) {
      worksheet.setRowCount(visual.anchor.toRow + 1)
    }
    if (visual.anchor.toColumn >= worksheet.getMaxColumns()) {
      worksheet.setColumnCount(visual.anchor.toColumn + 1)
    }
    // An out-of-bounds anchor would throw in getRange and take every other
    // visual down with it — clamp the display range to the sheet.
    const maxRows = worksheet.getMaxRows()
    const maxColumns = worksheet.getMaxColumns()
    const fromRow = Math.min(visual.anchor.fromRow, maxRows - 1)
    const fromColumn = Math.min(visual.anchor.fromColumn, maxColumns - 1)
    const toRow = Math.min(visual.anchor.toRow, maxRows - 1)
    const toColumn = Math.min(visual.anchor.toColumn, maxColumns - 1)
    const range = worksheet.getRange(
      fromRow,
      fromColumn,
      Math.max(1, toRow + 1 - fromRow),
      Math.max(1, toColumn + 1 - fromColumn),
    )
    // Pixel-exact frame (Excel twoCellAnchor semantics): each marker is a
    // cell plus an EMU offset inside it, so the box runs from the `from`
    // marker to the `to` marker — not across the whole cell range.
    const columnWidth = (index: number): number => Math.max(worksheet.getColumnWidth(index), 1)
    const rowHeight = (index: number): number => Math.max(worksheet.getRowHeight(index), 1)
    const marginX = Math.max(0, visual.anchor.fromColumnOffset / EMU_PER_PIXEL)
    const marginY = Math.max(0, visual.anchor.fromRowOffset / EMU_PER_PIXEL)
    const width = markerSpan(
      { index: fromColumn, offset: marginX },
      markerFrom(toColumn, visual.anchor.toColumnOffset),
      columnWidth,
    )
    const height = markerSpan(
      { index: fromRow, offset: marginY },
      markerFrom(toRow, visual.anchor.toRowOffset),
      rowHeight,
    )
    // Degenerate anchors (oneCellAnchor fallback parses to a zero span):
    // keep the legacy behavior of filling the clamped cell range.
    const layout =
      width >= MIN_FRAME_PIXELS / 2 && height >= MIN_FRAME_PIXELS / 2
        ? { width, height, marginX, marginY }
        : {}
    const floating = worksheet.addFloatDomToRange(
      range,
      {
        componentKey,
        allowTransform: false,
        eventPassThrough: false,
      },
      layout,
      `${file.sessionId}-${visual.id}`,
    )
    if (floating) disposables.push(floating)
  }
  return disposables
}

export interface SparklineGroupState {
  readonly type: 'line' | 'column' | 'stacked'
  readonly color?: string | undefined
  readonly negativeColor?: string | undefined
  readonly cells: readonly { readonly cell: string; readonly sourceRef: string }[]
}

const MAX_SPARKLINES_PER_SHEET = 200

/// In-cell minicharts: one pass-through float DOM per sparkline cell, values
/// read live from the grid (journal edits land there before the save).
export function installSparklines(
  runtime: UniverRuntime,
  groups: readonly SparklineGroupState[],
  sheetId: string,
): Disposable[] {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getSheetBySheetId(sheetId)
  if (!workbook || !worksheet) return []
  const sheetIdByName = new Map(
    workbook.getSheets().map((sheet) => [sheet.getSheetName(), sheet.getSheetId()] as const),
  )
  const disposables: Disposable[] = []
  let budget = MAX_SPARKLINES_PER_SHEET
  for (const [groupIndex, group] of groups.entries()) {
    for (const [cellIndex, entry] of group.cells.entries()) {
      if (budget <= 0) return disposables
      budget -= 1
      const split = splitSheetRef(entry.sourceRef)
      const source = split
        ? workbook.getSheetBySheetId(sheetIdByName.get(split.sheetName) ?? '')
        : worksheet
      if (!source) continue
      let values: number[]
      let host: { row: number; column: number }
      try {
        host = parseAddress(entry.cell.replace(/\$/g, ''))
        const grid = source
          .getRange((split?.range ?? entry.sourceRef).replace(/\$/g, ''))
          .getValues() as (string | number | boolean | null | undefined)[][]
        values = grid.flat().map((value) => {
          const numeric = typeof value === 'number' ? value : Number(String(value ?? '').trim())
          return Number.isFinite(numeric) ? numeric : 0
        })
      } catch {
        continue
      }
      if (values.length === 0) continue
      const componentKey = `sparkline-${sheetId}-${groupIndex}-${cellIndex}`
      disposables.push(
        runtime.univerAPI.registerComponent(componentKey, () => (
          <Sparkline
            values={values}
            type={group.type}
            color={group.color}
            negativeColor={group.negativeColor}
          />
        )),
      )
      const range = worksheet.getRange(host.row, host.column, 1, 1)
      const floating = worksheet.addFloatDomToRange(
        range,
        { componentKey, allowTransform: false, eventPassThrough: true },
        {},
        componentKey,
      )
      if (floating) disposables.push(floating)
    }
  }
  return disposables
}

function Sparkline({
  values,
  type,
  color,
  negativeColor,
}: {
  readonly values: readonly number[]
  readonly type: SparklineGroupState['type']
  readonly color?: string | undefined
  readonly negativeColor?: string | undefined
}): React.JSX.Element {
  const stroke = color ?? '#376092'
  const negative = negativeColor ?? '#d00000'
  const width = 100
  const height = 26
  if (type === 'line') {
    const minimum = Math.min(...values)
    const maximum = Math.max(...values)
    const span = maximum - minimum || 1
    const count = Math.max(1, values.length - 1)
    const points = values
      .map(
        (value, index) =>
          `${2 + (index / count) * (width - 4)},${height - 3 - ((value - minimum) / span) * (height - 6)}`,
      )
      .join(' ')
    return (
      <svg className="sparkline-svg" viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="none">
        <polyline points={points} fill="none" stroke={stroke} strokeWidth="1.6" />
      </svg>
    )
  }
  // column / stacked (win-loss): bars around a zero baseline.
  const isWinLoss = type === 'stacked'
  const maximum = Math.max(...values.map((value) => Math.abs(value)), 0) || 1
  const barWidth = Math.max(1.5, (width - 4) / values.length - 1)
  const zeroY = height / 2
  return (
    <svg className="sparkline-svg" viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="none">
      {values.map((value, index) => {
        const magnitude = isWinLoss
          ? value === 0
            ? 0
            : zeroY - 3
          : (Math.abs(value) / maximum) * (zeroY - 2)
        return (
          <rect
            key={index}
            x={2 + index * ((width - 4) / values.length)}
            y={value >= 0 ? zeroY - magnitude : zeroY}
            width={barWidth}
            height={Math.max(magnitude, value === 0 ? 0.5 : 1)}
            fill={value < 0 ? negative : stroke}
          />
        )
      })}
    </svg>
  )
}

function WorkbookVisual({
  file,
  visual,
  chartEditing,
}: {
  readonly file: VisualHost
  readonly visual: WorkbookVisualObject
  readonly chartEditing?: ChartEditing | undefined
}): React.JSX.Element {
  if (visual.kind === 'chart' && visual.chart) {
    return (
      <ChartVisual
        chart={visual.chart}
        visualId={visual.id}
        chartPath={visual.chartPath}
        chartEditing={chartEditing}
      />
    )
  }
  if (visual.kind === 'image') {
    return <ImageVisual file={file} visual={visual} />
  }
  return <ShapeVisual visual={visual} />
}

/// A drag commit tears down and re-registers every float DOM, dropping the
/// browser focus that doubles as the selection state — this hands it back
/// to the moved visual when its replacement mounts.
let pendingFocusId: string | null = null

/// Reinstalls must not run mid-drag: disposing the dragged node kills its
/// pointer capture and the drag dies silently. Install queues poll this and
/// wait for the drop.
let activeDragCount = 0

export function isVisualDragActive(): boolean {
  return activeDragCount > 0
}

/// Reinstalls also wait while an inline chart editor is open — disposing its
/// float DOM would silently discard everything typed into it.
let openEditorCount = 0

export function isChartEditorOpen(): boolean {
  return openEditorCount > 0
}

export function trackChartEditorOpen(): () => void {
  openEditorCount += 1
  return () => {
    openEditorCount -= 1
  }
}

/// A move commit reinstalls the float ~100ms later; removing the body-level
/// ghost on pointerup would flash the visual back at its old anchor for that
/// window. The ghost is parked here and removed once the replacement mounts;
/// the timer covers rejected edits, where no reinstall ever comes.
let dropGhost: { id: string; ghost: HTMLElement; host: HTMLElement; timer: number } | null = null

function discardDropGhost(): void {
  if (!dropGhost) return
  const { ghost, host, timer } = dropGhost
  dropGhost = null
  clearTimeout(timer)
  ghost.remove()
  host.style.visibility = ''
}

function parkDropGhost(id: string, ghost: HTMLElement, host: HTMLElement): void {
  discardDropGhost()
  dropGhost = { id, ghost, host, timer: window.setTimeout(discardDropGhost, 1200) }
}

function claimDropGhost(id: string, node: HTMLElement): void {
  if (dropGhost?.id !== id) return
  const { ghost, timer } = dropGhost
  dropGhost = null
  clearTimeout(timer)
  // The float container may attach/lay out late (see tryFocus below): keep
  // the preview until the replacement actually paints at the new anchor.
  const settle = (attempts: number): void => {
    if (attempts > 0 && (!node.isConnected || node.getBoundingClientRect().width === 0)) {
      setTimeout(() => settle(attempts - 1), 80)
      return
    }
    requestAnimationFrame(() => requestAnimationFrame(() => ghost.remove()))
  }
  settle(12)
}

export interface VisualSelectionListener {
  readonly select: (visual: WorkbookVisualObject) => void
  readonly deselect: () => void
}

/// Selection lives here, not in DOM focus: reinstalls and Univer's own DOM
/// churn drop focus, but the id survives and the remounted wrapper re-renders
/// selected. A module singleton avoids threading callbacks through every
/// install call site; the app mirrors it for the contextual ribbon tab.
let selectedVisualId: string | null = null
const selectionSubscribers = new Set<() => void>()
let selectionListener: VisualSelectionListener | null = null

function selectVisual(visual: WorkbookVisualObject | null): void {
  const nextId = visual?.id ?? null
  if (nextId !== selectedVisualId) {
    selectedVisualId = nextId
    for (const notify of selectionSubscribers) notify()
    // Element selection never outlives its chart's selection.
    if (selectedChartElement && selectedChartElement.visualId !== nextId) {
      selectedChartElement = null
      for (const notify of elementSubscribers) notify()
    }
  }
  if (visual) selectionListener?.select(visual)
  else selectionListener?.deselect()
}

/// Clears the selection; with `onlyId`, only when that visual is selected
/// (deleting one visual must not deselect another).
export function clearVisualSelection(onlyId?: string): void {
  if (onlyId !== undefined && onlyId !== selectedVisualId) return
  selectVisual(null)
}

/// Selection of an element INSIDE a chart (title, legend, one series, one
/// point/slice, an axis); the format pane reads it to scope its controls.
/// Module store for the same reason as the visual selection.
export type ChartElementRef =
  | { kind: 'chart' | 'title' | 'legend' | 'value-axis' | 'category-axis' }
  | { kind: 'series'; seriesIndex: number }
  | { kind: 'point'; seriesIndex: number; pointIndex: number }

let selectedChartElement: { visualId: string; element: ChartElementRef } | null = null
const elementSubscribers = new Set<() => void>()

export function setChartElementSelection(visualId: string, element: ChartElementRef | null): void {
  selectedChartElement = element === null ? null : { visualId, element }
  for (const notify of elementSubscribers) notify()
}

export function getChartElementSelection(): { visualId: string; element: ChartElementRef } | null {
  return selectedChartElement
}

export function subscribeChartElementSelection(notify: () => void): () => void {
  elementSubscribers.add(notify)
  return () => elementSubscribers.delete(notify)
}

/// Context-menu / dialog requests bubble to the app through one listener
/// (same recipe as the visual-selection listener).
export type ChartDialogKind = 'select-data' | 'format'
let chartDialogListener: ((editKey: string, dialog: ChartDialogKind) => void) | null = null

export function setChartDialogListener(
  listener: ((editKey: string, dialog: ChartDialogKind) => void) | null,
): void {
  chartDialogListener = listener
}

export function setVisualSelectionListener(listener: VisualSelectionListener | null): void {
  selectionListener = listener
}

function useIsSelected(visualId: string): boolean {
  return useSyncExternalStore(
    (notify) => {
      selectionSubscribers.add(notify)
      return () => selectionSubscribers.delete(notify)
    },
    () => selectedVisualId === visualId,
  )
}

const RESIZE_CORNERS = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const
type ResizeCorner = (typeof RESIZE_CORNERS)[number]

const cornerEast = (corner: ResizeCorner): boolean =>
  corner === 'ne' || corner === 'e' || corner === 'se'
const cornerWest = (corner: ResizeCorner): boolean =>
  corner === 'nw' || corner === 'w' || corner === 'sw'
const cornerNorth = (corner: ResizeCorner): boolean =>
  corner === 'nw' || corner === 'n' || corner === 'ne'
const cornerSouth = (corner: ResizeCorner): boolean =>
  corner === 'sw' || corner === 's' || corner === 'se'

/// xlsx drawing offsets are EMU; 9525 EMU per CSS pixel at 96dpi.
export const EMU_PER_PIXEL = 9525
/// Frames smaller than this collapse resize handles into each other.
const MIN_FRAME_PIXELS = 24
/// xlsx sheet limits: drags may leave the data-sized lazy grid (install
/// grows it to the anchor), but never the real spreadsheet bounds.
const XLSX_MAX_COLUMN = 16383
const XLSX_MAX_ROW = 1048575

/// One edge of a drawing anchor: a cell index plus a pixel offset inside
/// that cell (the px equivalent of xlsx's `<xdr:col>` + `<xdr:colOff>`).
export interface AnchorMarker {
  index: number
  offset: number
}

export const markerFrom = (index: number, offsetEmu: number): AnchorMarker => ({
  index,
  offset: offsetEmu / EMU_PER_PIXEL,
})

/// Move a marker by a pixel delta, carrying across real row/column sizes.
/// Clamps at the sheet start and inside the last row/column.
export function walkMarker(
  marker: AnchorMarker,
  delta: number,
  sizeOf: (index: number) => number,
  maxIndex: number,
): AnchorMarker {
  let index = Math.min(marker.index, maxIndex)
  let offset = marker.offset + delta
  while (offset < 0 && index > 0) {
    index -= 1
    offset += sizeOf(index)
  }
  if (offset < 0) offset = 0
  while (index < maxIndex && offset >= sizeOf(index)) {
    offset -= sizeOf(index)
    index += 1
  }
  if (index >= maxIndex) offset = Math.min(offset, sizeOf(maxIndex))
  return { index, offset }
}

/// Pixel distance between two markers (negative when `to` sits before `from`).
function markerSpan(
  from: AnchorMarker,
  to: AnchorMarker,
  sizeOf: (index: number) => number,
): number {
  const span = to.offset - from.offset
  const low = Math.min(from.index, to.index)
  const high = Math.max(from.index, to.index)
  let cells = 0
  for (let index = low; index < high; index += 1) cells += sizeOf(index)
  return span + (from.index <= to.index ? cells : -cells)
}

function EditableShapeVisual({
  file,
  visual,
  worksheet,
  allowText,
  chartEditing,
  onEdit,
}: {
  readonly file: VisualHost
  readonly visual: WorkbookVisualObject
  readonly worksheet: UniverWorksheet
  readonly allowText: boolean
  readonly chartEditing?: ChartEditing | undefined
  readonly onEdit: (visualId: string, changes: ShapeEditChanges) => void
}): React.JSX.Element {
  const [drag, setDrag] = useState<{
    mode: 'move' | 'resize'
    corner: ResizeCorner
    startX: number
    startY: number
    dx: number
    dy: number
  } | null>(null)
  const [textEditing, setTextEditing] = useState(false)
  const [chartEditRequest, setChartEditRequest] = useState(0)
  const editorRef = useRef<HTMLDivElement | null>(null)
  // Move previews render as a clone on <body>: the Univer float container is
  // overflow-hidden and pinned to the old anchor, so translating the visual
  // inside it clips everything that leaves the original box.
  const ghostRef = useRef<HTMLDivElement | null>(null)
  const moveDeltaRef = useRef({ dx: 0, dy: 0 })
  const isSelected = useIsSelected(visual.id)
  const isChart = visual.kind === 'chart' && visual.chart !== undefined

  const dragActiveRef = useRef(false)
  const setDragActive = (active: boolean): void => {
    if (dragActiveRef.current === active) return
    dragActiveRef.current = active
    activeDragCount += active ? 1 : -1
  }

  const removeGhost = (host: HTMLElement | null): void => {
    ghostRef.current?.remove()
    ghostRef.current = null
    if (host) host.style.visibility = ''
  }

  const hostNodeRef = useRef<HTMLElement | null>(null)

  useEffect(() => {
    if (hostNodeRef.current) claimDropGhost(visual.id, hostNodeRef.current)
    return () => {
      ghostRef.current?.remove()
      ghostRef.current = null
      setDragActive(false)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const commitDrag = (
    mode: 'move' | 'resize',
    corner: ResizeCorner,
    dxRaw: number,
    dyRaw: number,
  ): boolean => {
    const zoom = worksheet.getZoom() || 1
    const anchor = visual.anchor
    // The lazy grid is data-sized; past its edge walk with the sheet
    // defaults — install grows the grid to the committed anchor, and Univer
    // gives the new rows/columns exactly these default sizes.
    const config = worksheet.getSheet().getConfig()
    const gridColumns = worksheet.getMaxColumns()
    const gridRows = worksheet.getMaxRows()
    const columnWidth = (index: number): number =>
      index < gridColumns
        ? Math.max(worksheet.getColumnWidth(index), 1)
        : Math.max(config.defaultColumnWidth, 1)
    const rowHeight = (index: number): number =>
      index < gridRows
        ? Math.max(worksheet.getRowHeight(index), 1)
        : Math.max(config.defaultRowHeight, 1)
    const maxColumn = XLSX_MAX_COLUMN
    const maxRow = XLSX_MAX_ROW
    let fromX = markerFrom(anchor.fromColumn, anchor.fromColumnOffset)
    let fromY = markerFrom(anchor.fromRow, anchor.fromRowOffset)
    let toX = markerFrom(anchor.toColumn, anchor.toColumnOffset)
    let toY = markerFrom(anchor.toRow, anchor.toRowOffset)
    const dx = dxRaw / zoom
    const dy = dyRaw / zoom
    if (mode === 'move') {
      // Free placement: keep the frame size. The grid edge caps
      // the shift — walk the leading marker first and apply the distance it
      // actually covered to both markers, so a clamped drag cannot shrink
      // the frame.
      const effectiveShift = (
        delta: number,
        from: AnchorMarker,
        to: AnchorMarker,
        sizeOf: (index: number) => number,
        maxIndex: number,
      ): number => {
        if (delta === 0) return 0
        const leading = delta > 0 ? to : from
        return markerSpan(leading, walkMarker(leading, delta, sizeOf, maxIndex), sizeOf)
      }
      const shiftX = effectiveShift(dx, fromX, toX, columnWidth, maxColumn)
      const shiftY = effectiveShift(dy, fromY, toY, rowHeight, maxRow)
      fromX = walkMarker(fromX, shiftX, columnWidth, maxColumn)
      fromY = walkMarker(fromY, shiftY, rowHeight, maxRow)
      toX = walkMarker(toX, shiftX, columnWidth, maxColumn)
      toY = walkMarker(toY, shiftY, rowHeight, maxRow)
    } else {
      // Resize handles: each edge moves its own anchor marker,
      // never past the opposite edge minus the minimum frame.
      if (cornerEast(corner)) {
        toX = walkMarker(toX, dx, columnWidth, maxColumn)
        if (markerSpan(fromX, toX, columnWidth) < MIN_FRAME_PIXELS) {
          toX = walkMarker(fromX, MIN_FRAME_PIXELS, columnWidth, maxColumn)
        }
      } else if (cornerWest(corner)) {
        fromX = walkMarker(fromX, dx, columnWidth, maxColumn)
        if (markerSpan(fromX, toX, columnWidth) < MIN_FRAME_PIXELS) {
          fromX = walkMarker(toX, -MIN_FRAME_PIXELS, columnWidth, maxColumn)
        }
      }
      if (cornerSouth(corner)) {
        toY = walkMarker(toY, dy, rowHeight, maxRow)
        if (markerSpan(fromY, toY, rowHeight) < MIN_FRAME_PIXELS) {
          toY = walkMarker(fromY, MIN_FRAME_PIXELS, rowHeight, maxRow)
        }
      } else if (cornerNorth(corner)) {
        fromY = walkMarker(fromY, dy, rowHeight, maxRow)
        if (markerSpan(fromY, toY, rowHeight) < MIN_FRAME_PIXELS) {
          fromY = walkMarker(toY, -MIN_FRAME_PIXELS, rowHeight, maxRow)
        }
      }
    }
    const next = {
      fromRow: fromY.index,
      fromColumn: fromX.index,
      fromRowOffset: Math.round(fromY.offset * EMU_PER_PIXEL),
      fromColumnOffset: Math.round(fromX.offset * EMU_PER_PIXEL),
      toRow: toY.index,
      toColumn: toX.index,
      toRowOffset: Math.round(toY.offset * EMU_PER_PIXEL),
      toColumnOffset: Math.round(toX.offset * EMU_PER_PIXEL),
    }
    if (
      next.fromRow === anchor.fromRow &&
      next.fromColumn === anchor.fromColumn &&
      next.fromRowOffset === anchor.fromRowOffset &&
      next.fromColumnOffset === anchor.fromColumnOffset &&
      next.toRow === anchor.toRow &&
      next.toColumn === anchor.toColumn &&
      next.toRowOffset === anchor.toRowOffset &&
      next.toColumnOffset === anchor.toColumnOffset
    )
      return false
    pendingFocusId = visual.id
    onEdit(visual.id, { anchor: next })
    return true
  }

  return (
    <div
      className={`shape-editable${isSelected ? ' selected' : ''}`}
      tabIndex={0}
      ref={(node) => {
        hostNodeRef.current = node
        if (!node || pendingFocusId !== visual.id) return
        pendingFocusId = null
        // The float container may not be attached/laid out yet; retry until
        // focus actually lands. preventScroll: a scroll-into-view here would
        // shift the sheet under the just-dropped visual.
        const tryFocus = (attempts: number): void => {
          node.focus({ preventScroll: true })
          if (document.activeElement !== node && attempts > 0) {
            setTimeout(() => tryFocus(attempts - 1), 80)
          }
        }
        tryFocus(12)
      }}
      onPointerDown={(event) => {
        if (textEditing || event.button !== 0) return
        // Chart editor buttons/inputs keep their own interactions.
        if ((event.target as HTMLElement).closest('button, .chart-editor')) return
        const handleCorner = (event.target as HTMLElement).dataset['corner'] as
          ResizeCorner | undefined
        const mode = handleCorner ? ('resize' as const) : ('move' as const)
        event.currentTarget.setPointerCapture(event.pointerId)
        setDragActive(true)
        moveDeltaRef.current = { dx: 0, dy: 0 }
        setDrag({
          mode,
          corner: handleCorner ?? 'se',
          startX: event.clientX,
          startY: event.clientY,
          dx: 0,
          dy: 0,
        })
        event.preventDefault()
        event.stopPropagation()
        // preventScroll: grabbing a half-visible visual must not scroll it
        // into view mid-drag — the sheet would shift under the pointer.
        event.currentTarget.focus({ preventScroll: true })
      }}
      onFocus={() => selectVisual(visual)}
      onKeyDown={(event) => {
        if (event.target !== event.currentTarget) return
        if (event.key === 'Delete' || event.key === 'Backspace') {
          event.preventDefault()
          event.stopPropagation()
          onEdit(visual.id, { remove: true })
          clearVisualSelection(visual.id)
        } else if (event.key === 'Escape') {
          event.preventDefault()
          event.currentTarget.blur()
          clearVisualSelection(visual.id)
        }
      }}
      onPointerMove={(event) => {
        if (!drag) return
        const dx = event.clientX - drag.startX
        const dy = event.clientY - drag.startY
        if (drag.mode === 'resize') {
          setDrag({ ...drag, dx, dy })
          return
        }
        // Move: drive the body-level ghost directly, no re-render per event.
        moveDeltaRef.current = { dx, dy }
        const host = event.currentTarget
        if (!ghostRef.current && (Math.abs(dx) > 2 || Math.abs(dy) > 2)) {
          const rect = host.getBoundingClientRect()
          const container = document.createElement('div')
          container.className = 'shape-drag-ghost'
          container.style.left = `${rect.left}px`
          container.style.top = `${rect.top}px`
          container.style.width = `${rect.width}px`
          container.style.height = `${rect.height}px`
          container.appendChild(host.cloneNode(true))
          document.body.appendChild(container)
          // Pointer capture keeps delivering events to the hidden host.
          host.style.visibility = 'hidden'
          ghostRef.current = container
        }
        if (ghostRef.current) {
          ghostRef.current.style.transform = `translate(${dx}px, ${dy}px)`
        }
      }}
      onPointerUp={(event) => {
        if (!drag) return
        const { mode, corner } = drag
        const { dx, dy } = mode === 'move' ? moveDeltaRef.current : drag
        setDrag(null)
        setDragActive(false)
        const committed = (Math.abs(dx) > 2 || Math.abs(dy) > 2) && commitDrag(mode, corner, dx, dy)
        if (committed && ghostRef.current) {
          parkDropGhost(visual.id, ghostRef.current, event.currentTarget)
          ghostRef.current = null
        } else {
          removeGhost(event.currentTarget)
        }
      }}
      onPointerCancel={(event) => {
        removeGhost(event.currentTarget)
        setDrag(null)
        setDragActive(false)
      }}
      onDoubleClick={
        allowText
          ? () => setTextEditing(true)
          : isChart
            ? (event) => {
                if ((event.target as HTMLElement).closest('button, .chart-editor')) return
                setChartEditRequest((count) => count + 1)
              }
            : undefined
      }
      title={
        textEditing
          ? undefined
          : allowText
            ? t('appVisualHintText')
            : isChart
              ? t('appVisualHintEdit')
              : t('appVisualHint')
      }
    >
      {visual.kind === 'chart' && visual.chart ? (
        <ChartVisual
          chart={visual.chart}
          visualId={visual.id}
          chartPath={visual.chartPath}
          chartEditing={chartEditing}
          onRemove={() => {
            onEdit(visual.id, { remove: true })
            clearVisualSelection(visual.id)
          }}
          openEditorSignal={chartEditRequest}
        />
      ) : visual.kind === 'image' ? (
        <ImageVisual file={file} visual={visual} />
      ) : (
        <ShapeVisual visual={textEditing ? { ...visual, text: '' } : visual} />
      )}
      {!textEditing && (
        <button
          className="shape-delete-button"
          title={t('appDeleteVisualTitle')}
          onPointerDown={(event) => event.stopPropagation()}
          onClick={(event) => {
            event.stopPropagation()
            onEdit(visual.id, { remove: true })
          }}
        >
          ✕
        </button>
      )}
      {textEditing && (
        <div
          className="shape-text-editor"
          contentEditable
          suppressContentEditableWarning
          role="textbox"
          ref={(node) => {
            if (!node || editorRef.current === node) return
            editorRef.current = node
            node.textContent = visual.text ?? ''
            node.focus()
            const selection = window.getSelection()
            const range = document.createRange()
            range.selectNodeContents(node)
            selection?.removeAllRanges()
            selection?.addRange(range)
          }}
          onBlur={() => {
            const next = editorRef.current?.textContent ?? ''
            editorRef.current = null
            setTextEditing(false)
            if (next !== (visual.text ?? '')) onEdit(visual.id, { text: next })
          }}
          onKeyDown={(event) => {
            if (event.key === 'Escape') {
              event.preventDefault()
              if (editorRef.current) editorRef.current.textContent = visual.text ?? ''
              event.currentTarget.blur()
            } else if (event.key === 'Enter' && !event.shiftKey) {
              event.preventDefault()
              event.currentTarget.blur()
            }
          }}
        />
      )}
      {drag?.mode === 'resize' && (
        <div
          className="shape-resize-outline"
          style={{
            left: cornerWest(drag.corner) ? drag.dx : 0,
            top: cornerNorth(drag.corner) ? drag.dy : 0,
            width: cornerEast(drag.corner)
              ? `calc(100% + ${drag.dx}px)`
              : cornerWest(drag.corner)
                ? `calc(100% - ${drag.dx}px)`
                : '100%',
            height: cornerSouth(drag.corner)
              ? `calc(100% + ${drag.dy}px)`
              : cornerNorth(drag.corner)
                ? `calc(100% - ${drag.dy}px)`
                : '100%',
          }}
        />
      )}
      {!textEditing &&
        RESIZE_CORNERS.map((corner) => (
          <span
            key={corner}
            className={`shape-handle handle-${corner}${corner === 'se' ? ' shape-resize-handle' : ''}`}
            data-handle="resize"
            data-corner={corner}
          />
        ))}
    </div>
  )
}

function ShapeVisual({ visual }: { readonly visual: WorkbookVisualObject }): React.JSX.Element {
  const type = visual.shapeType ?? ''
  const fill = visual.fillColor ?? '#DDEBF7'
  // Same geometry source as the gallery previews (and the other apps'
  // renderers), so every insertable preset draws its real silhouette.
  const d = shapePreviewPath(type, 100, 100)
  if (!d) {
    return (
      <div className="xlsx-shape">
        <span>{visual.text ?? visual.name ?? t('appDrawingObject')}</span>
      </div>
    )
  }
  return (
    <div
      className="xlsx-shape-drawn"
      style={visual.rotation ? { transform: `rotate(${visual.rotation}deg)` } : undefined}
    >
      <svg viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
        <path d={d} fill={fill} stroke="#00000022" />
      </svg>
      {visual.text && <span className="shape-text">{visual.text}</span>}
    </div>
  )
}

function ImageVisual({
  file,
  visual,
}: {
  readonly file: VisualHost
  readonly visual: WorkbookVisualObject
}): React.JSX.Element {
  const [source, setSource] = useState<string | null>(visual.mediaDataUrl ?? null)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    // Session-added images carry their bytes inline; nothing to fetch.
    if (visual.mediaDataUrl) return
    let isCurrent = true
    void window.desktopApi
      .readWorkbookMedia({
        sessionId: file.sessionId,
        visualId: visual.id,
      })
      .then((media) => {
        if (isCurrent) setSource(`data:${media.mediaType};base64,${media.base64}`)
      })
      .catch((reason: unknown) => {
        if (isCurrent) setError(reason instanceof Error ? reason.message : t('appImageLoadFailed'))
      })
    return () => {
      isCurrent = false
    }
  }, [file.sessionId, visual.id, visual.mediaDataUrl])

  if (error) return <div className="xlsx-visual-error">{error}</div>
  if (!source) return <div className="xlsx-visual-loading">{t('appImageLoading')}</div>
  return <img className="xlsx-image" src={source} alt={visual.name ?? t('appWorkbookImageAlt')} />
}

type ChartMetadata = NonNullable<WorkbookVisualObject['chart']>
type ChartSeries = ChartMetadata['series'][number]

const CHART_TYPE_OPTIONS = [
  { value: 'column', labelKey: 'appChartColumn' },
  { value: 'bar', labelKey: 'appChartBar' },
  { value: 'line', labelKey: 'appChartLine' },
  { value: 'area', labelKey: 'appChartTypeArea' },
  { value: 'pie', labelKey: 'appChartPie' },
  { value: 'doughnut', labelKey: 'appChartDoughnut' },
] as const

/// Applies journaled edits over the parsed chart cache so the on-screen
/// chart previews the pending save (shared with the baked-edit path).
function withChartEdit(chart: ChartMetadata, edit: ChartEditData | undefined): ChartMetadata {
  return applyChartStateEdit(chart, edit) as ChartMetadata
}

/// A chart converts only when it has exactly one supported plot type.
/// Single-plot column/bar/line/area/pie/doughnut charts convert; everything
/// else (scatter, combos, 3D) fails closed at save time — reject up front.
export function convertibleType(chart: ChartMetadata): string | null {
  if (chart.chartTypes.length !== 1) return null
  const only = chart.chartTypes[0]
  if (only === 'lineChart') return 'line'
  if (only === 'areaChart') return 'area'
  if (only === 'barChart') return chart.barDirection === 'bar' ? 'bar' : 'column'
  if (only === 'pieChart') return 'pie'
  if (only === 'doughnutChart') return 'doughnut'
  return null
}

/// Excel narrowing: first click on a series member selects the series, a
/// second click on the same series narrows to the single point.
function narrowSelection(
  current: ChartElementRef | null,
  seriesIndex: number,
  pointIndex: number,
): ChartElementRef {
  const sameSeries =
    current &&
    (current.kind === 'series' || current.kind === 'point') &&
    current.seriesIndex === seriesIndex
  return sameSeries ? { kind: 'point', seriesIndex, pointIndex } : { kind: 'series', seriesIndex }
}

interface ChartElementProps {
  readonly onElement?: ((element: ChartElementRef) => void) | undefined
  readonly selectedEl?: ChartElementRef | null | undefined
}

function isSelectedPoint(
  element: ChartElementRef | null | undefined,
  seriesIndex: number,
  pointIndex: number,
): boolean {
  if (!element) return false
  if (element.kind === 'series') return element.seriesIndex === seriesIndex
  if (element.kind === 'point') {
    return element.seriesIndex === seriesIndex && element.pointIndex === pointIndex
  }
  return false
}

function ChartVisual({
  chart: sourceChart,
  visualId,
  chartPath,
  chartEditing,
  onRemove,
  openEditorSignal = 0,
}: {
  readonly chart: ChartMetadata
  readonly visualId: string
  readonly chartPath?: string | undefined
  readonly chartEditing?: ChartEditing | undefined
  readonly onRemove?: (() => void) | undefined
  /// Bumped by the wrapper on double-click to open the inline editor.
  readonly openEditorSignal?: number
}): React.JSX.Element {
  const [isEditing, setIsEditing] = useState(false)
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null)
  useEffect(() => {
    if (openEditorSignal > 0) setIsEditing(true)
  }, [openEditorSignal])
  const selectedEl = useSyncExternalStore(subscribeChartElementSelection, () =>
    selectedChartElement?.visualId === visualId ? selectedChartElement.element : null,
  )
  // In-memory charts (session adds, demo) bake edits into the visual itself,
  // so only file charts carry a pending overlay.
  const editKey = chartPath ?? visualId
  const pendingEdit = chartPath ? chartEditing?.edits.get(chartPath) : undefined
  // Files written without numCache (openpyxl, pandas) have series refs but no
  // cached points; hydrate those series from the referenced cells. Categories
  // hydrate on their own too (scatter xVal numRef without numCache).
  const readVector = chartEditing?.readVector
  const needsHydration = sourceChart.series.some(
    (series) =>
      (series.values.length === 0 && series.valuesRef) ||
      (series.categories.length === 0 && series.categoriesRef),
  )
  const [hydration, setHydration] = useState<ReadonlyMap<
    number,
    { values: readonly number[] | null; categories: readonly string[] | null }
  > | null>(null)
  useEffect(() => {
    if (!needsHydration || !readVector) return
    let cancelled = false
    void (async () => {
      // The lazy sheet may still be indexing right after open; retry briefly.
      for (let attempt = 0; attempt < 8 && !cancelled; attempt += 1) {
        const entries = new Map<
          number,
          { values: readonly number[] | null; categories: readonly string[] | null }
        >()
        let retry = false
        for (const [index, series] of sourceChart.series.entries()) {
          const valuesRef = series.values.length === 0 ? series.valuesRef : undefined
          const categoriesRef = series.categories.length === 0 ? series.categoriesRef : undefined
          if (valuesRef === undefined && categoriesRef === undefined) continue
          try {
            const values = valuesRef === undefined ? null : await readVector(editKey, valuesRef)
            // A failed categories read only aborts (and retries) when it is
            // the sole target; alongside values it stays best-effort.
            const categories =
              categoriesRef === undefined
                ? null
                : valuesRef === undefined
                  ? await readVector(editKey, categoriesRef)
                  : await readVector(editKey, categoriesRef).catch(() => null)
            if (values === null && categories === null) continue
            entries.set(index, {
              values:
                values?.vector.map((value) => {
                  const numeric =
                    typeof value === 'number' ? value : Number(String(value ?? '').trim())
                  return Number.isFinite(numeric) ? numeric : 0
                }) ?? null,
              categories:
                categories?.vector.map((value) => String(value ?? '').slice(0, 255)) ?? null,
            })
          } catch {
            retry = true
          }
        }
        if (!retry) {
          if (!cancelled && entries.size > 0) setHydration(entries)
          return
        }
        await new Promise((resolve) => setTimeout(resolve, 900))
      }
    })()
    return () => {
      cancelled = true
    }
  }, [needsHydration, readVector, editKey, sourceChart])
  const hydratedSource = hydration
    ? {
        ...sourceChart,
        series: sourceChart.series.map((series, index) => {
          const entry = hydration.get(index)
          if (!entry) return series
          return {
            ...series,
            ...(entry.values && series.values.length === 0 ? { values: [...entry.values] } : {}),
            ...(series.categories.length === 0 && entry.categories
              ? { categories: [...entry.categories] }
              : {}),
          }
        }),
      }
    : sourceChart
  const chart = withChartEdit(hydratedSource, pendingEdit)
  const populated = chart.series.filter((series) => series.values.length > 0)
  const primarySeries = populated[0]
  if (!primarySeries) {
    // Blank while hydration is still in flight; error only when it can't run.
    if (needsHydration && readVector && !hydration) {
      return <figure className="xlsx-chart" />
    }
    return <div className="xlsx-visual-error">{t('appChartCacheEmpty')}</div>
  }
  const types = chart.chartTypes
  const isDoughnut = types.includes('doughnutChart')
  const isPie = types.includes('pieChart') || isDoughnut
  const isScatter = types.includes('scatterChart')
  const isCombo = types.includes('barChart') && types.includes('lineChart')
  const isArea = types.includes('areaChart') && !types.includes('barChart')
  const isLine = types.includes('lineChart') && !types.includes('barChart')
  const isRadar = types.includes('radarChart') && !types.includes('barChart')
  const categoryFormat = chartCategoryFormat(chart)
  const canEdit = chartEditing !== undefined
  const selectElement = canEdit
    ? (element: ChartElementRef): void => setChartElementSelection(visualId, element)
    : undefined
  const labelsOn = chart.dataLabels !== undefined && chart.dataLabels !== 'none'
  return (
    <figure
      className="xlsx-chart"
      onClick={() => {
        if (selectedEl) setChartElementSelection(visualId, null)
      }}
      onContextMenu={
        canEdit
          ? (event) => {
              event.preventDefault()
              event.stopPropagation()
              // Right-click also selects the chart (menu actions and the
              // format pane both key off the selection).
              ;(event.currentTarget.closest('.shape-editable') as HTMLElement | null)?.focus()
              setMenu({ x: event.clientX, y: event.clientY })
            }
          : undefined
      }
    >
      {canEdit && !isEditing && (
        <button
          className="chart-edit-button"
          title={t('appEditChartTitle')}
          onClick={() => setIsEditing(true)}
        >
          ✎
        </button>
      )}
      {canEdit && isEditing && chartEditing && (
        <ChartEditor
          chart={chart}
          isPie={isPie}
          convertible={convertibleType(withChartEdit(sourceChart, undefined))}
          pendingType={pendingEdit?.chartType}
          readVector={rangeReader(chartEditing, editKey)}
          onApply={(edit) => {
            chartEditing.onEdit(editKey, edit)
            setIsEditing(false)
          }}
          onCancel={() => setIsEditing(false)}
        />
      )}
      <figcaption
        className={selectedEl?.kind === 'title' ? 'chart-el-selected' : undefined}
        onClick={
          selectElement
            ? (event) => {
                event.stopPropagation()
                selectElement({ kind: 'title' })
              }
            : undefined
        }
      >
        {chart.title}
      </figcaption>
      <div className={`chart-body chart-legend-${chart.legend ?? 'default'}`}>
        {isPie ? (
          <PieChart
            series={primarySeries}
            isDoughnut={isDoughnut}
            legend={chart.legend}
            dataLabels={chart.dataLabels}
            dataLabelPosition={chart.dataLabelPosition}
            dataLabelFormat={chart.dataLabelFormat}
            holeSizePct={chart.holeSizePct}
            categoryFormat={categoryFormat}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        ) : isRadar ? (
          <RadarChart
            seriesList={populated}
            dataLabels={chart.dataLabels}
            categoryFormat={categoryFormat}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        ) : isScatter ? (
          <ScatterChart
            seriesList={populated}
            axisTitles={chart.axisTitles}
            dataLabels={chart.dataLabels}
            gridlines={chart.gridlines}
            valueAxis={chart.valueAxis}
            categoryFormat={categoryFormat}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        ) : isArea ? (
          <AreaChart
            seriesList={populated}
            axisTitles={chart.axisTitles}
            dataLabels={chart.dataLabels}
            grouping={chart.grouping}
            gridlines={chart.gridlines}
            valueAxis={chart.valueAxis}
            categoryFormat={categoryFormat}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        ) : isLine ? (
          <LineChart
            seriesList={populated}
            axisTitles={chart.axisTitles}
            dataLabels={chart.dataLabels}
            grouping={chart.grouping}
            gridlines={chart.gridlines}
            valueAxis={chart.valueAxis}
            categoryFormat={categoryFormat}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        ) : (
          <BarChart
            seriesList={isCombo && populated.length > 1 ? populated.slice(0, -1) : populated}
            isHorizontal={chart.barDirection === 'bar'}
            lineSeries={
              isCombo && populated.length > 1 ? populated[populated.length - 1] : undefined
            }
            axisTitles={chart.axisTitles}
            dataLabels={chart.dataLabels}
            dataLabelPosition={chart.dataLabelPosition}
            dataLabelFormat={chart.dataLabelFormat}
            grouping={chart.grouping}
            gridlines={chart.gridlines}
            valueAxis={chart.valueAxis}
            gapWidthPct={chart.gapWidthPct}
            categoryFormat={categoryFormat}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        )}
        {!isPie &&
          chart.legend !== 'none' &&
          (chart.legend !== undefined || populated.length > 1) && (
            <SeriesLegend
              seriesList={populated}
              selected={selectedEl?.kind === 'legend'}
              onSelect={
                selectElement
                  ? (event) => {
                      event.stopPropagation()
                      selectElement({ kind: 'legend' })
                    }
                  : undefined
              }
            />
          )}
      </div>
      {menu && chartEditing && (
        <ChartContextMenu
          x={menu.x}
          y={menu.y}
          onClose={() => setMenu(null)}
          items={[
            {
              label: t('appFormatSelection'),
              action: () => chartDialogListener?.(editKey, 'format'),
            },
            {
              label: t('appSelectData'),
              action: () => chartDialogListener?.(editKey, 'select-data'),
            },
            ...(chartSupportsDataLabels(chart.chartTypes)
              ? [
                  {
                    label: labelsOn ? t('appDeleteDataLabels') : t('appAddDataLabels'),
                    action: () =>
                      chartEditing.onEdit(editKey, {
                        dataLabels: labelsOn ? 'none' : isPie ? 'category-percent' : 'value',
                      }),
                  },
                ]
              : []),
            ...(onRemove ? [{ label: t('appDeleteChart'), action: onRemove }] : []),
          ]}
        />
      )}
    </figure>
  )
}

function ChartContextMenu({
  x,
  y,
  items,
  onClose,
}: {
  readonly x: number
  readonly y: number
  readonly items: readonly { label: string; action: () => void }[]
  readonly onClose: () => void
}): React.JSX.Element {
  // Portal to the body: the float-DOM container clips overflowing children.
  return createPortal(
    <div
      className="chart-menu-backdrop"
      onClick={onClose}
      onContextMenu={(event) => event.preventDefault()}
    >
      <div
        className="chart-menu"
        style={{
          left: Math.min(x, window.innerWidth - 200),
          top: Math.min(y, window.innerHeight - items.length * 32 - 12),
        }}
      >
        {items.map((item) => (
          <button
            key={item.label}
            onClick={(event) => {
              event.stopPropagation()
              item.action()
              onClose()
            }}
          >
            {item.label}
          </button>
        ))}
      </div>
    </div>,
    document.body,
  )
}

type ChartAxisTitles = ChartMetadata['axisTitles']
type ChartDataLabels = ChartMetadata['dataLabels']
type ChartGrouping = ChartMetadata['grouping']

/// Category title along the bottom, value title rotated up the left edge
/// (already swapped by the caller for horizontal bars).
function AxisTitleTexts({
  bottom,
  left,
}: {
  readonly bottom?: string | null | undefined
  readonly left?: string | null | undefined
}): React.JSX.Element {
  return (
    <g>
      {bottom ? (
        <text x="320" y="317" textAnchor="middle" className="axis-title">
          {truncateLabel(bottom, 60)}
        </text>
      ) : null}
      {left ? (
        <text
          x="12"
          y="155"
          transform="rotate(-90 12 155)"
          textAnchor="middle"
          className="axis-title"
        >
          {truncateLabel(left, 40)}
        </text>
      ) : null}
    </g>
  )
}

function rangeReader(
  chartEditing: ChartEditing,
  editKey: string,
): ((range: string) => Promise<ChartVectorRead>) | undefined {
  const readVector = chartEditing.readVector
  if (!readVector) return undefined
  return (range) => readVector(editKey, range)
}

/// Pie/doughnut slices color per point: explicit `c:dPt` fills win, the
/// Office palette cycles underneath (Excel's varyColors default).
function pieSliceColor(series: SeriesLike, index: number): string {
  return (
    series.pointColors?.find((point) => point.index === index)?.color ??
    chartColors[index % chartColors.length] ??
    '#4472c4'
  )
}

/// Inline editor: title, chart type (axis-based family only), series colors.
interface SeriesRangeInput {
  readonly values: string
  readonly categories: string
}

function ChartEditor({
  chart,
  isPie,
  convertible,
  pendingType,
  readVector,
  onApply,
  onCancel,
}: {
  readonly chart: ChartMetadata
  readonly isPie: boolean
  readonly convertible: string | null
  readonly pendingType: ChartEditData['chartType'] | undefined
  readonly readVector?: ((range: string) => Promise<ChartVectorRead>) | undefined
  readonly onApply: (edit: ChartEditData) => void
  readonly onCancel: () => void
}): React.JSX.Element {
  const [title, setTitle] = useState(chart.title)
  const [chartType, setChartType] = useState(pendingType ?? convertible ?? '')
  const [colors, setColors] = useState<Record<string, string>>({})
  const [sliceColors, setSliceColors] = useState<Record<string, string>>({})
  const [ranges, setRanges] = useState<Record<string, SeriesRangeInput>>({})
  const [rangeError, setRangeError] = useState<string | null>(null)
  const [isReading, setIsReading] = useState(false)
  useEffect(() => trackChartEditorOpen(), [])

  async function seriesRangeEdits(): Promise<NonNullable<ChartEditData['series']>> {
    if (!readVector) return []
    const entries: NonNullable<ChartEditData['series']>[number][] = []
    for (const [key, input] of Object.entries(ranges)) {
      const valuesRange = input.values.trim()
      const categoriesRange = input.categories.trim()
      if (!valuesRange && !categoriesRange) continue
      entries.push({
        index: Number(key),
        ...(valuesRange
          ? await readVector(valuesRange).then((read) => ({
              values: read.vector.map((value) => {
                const numeric =
                  typeof value === 'number' ? value : Number(String(value ?? '').trim())
                return Number.isFinite(numeric) ? numeric : 0
              }),
              valuesRef: read.ref,
            }))
          : {}),
        ...(categoriesRange
          ? await readVector(categoriesRange).then((read) => ({
              categories: read.vector.map((value) => String(value ?? '').slice(0, 255)),
              categoriesRef: read.ref,
            }))
          : {}),
      })
    }
    return entries
  }

  async function handleApply(): Promise<void> {
    let series: NonNullable<ChartEditData['series']>
    try {
      setIsReading(true)
      setRangeError(null)
      series = await seriesRangeEdits()
    } catch (error: unknown) {
      setRangeError(error instanceof Error ? error.message : t('appReadRangeFailed'))
      return
    } finally {
      setIsReading(false)
    }
    const edit: ChartEditData = {
      ...(title !== chart.title ? { title } : {}),
      // Compare against the effective type (pending conversion included) so
      // selecting the original type back is a real edit, not a no-op.
      ...(convertible && chartType && chartType !== (pendingType ?? convertible)
        ? { chartType: chartType as NonNullable<ChartEditData['chartType']> }
        : {}),
      ...(Object.keys(colors).length > 0 ? { seriesColors: colors } : {}),
      ...(Object.keys(sliceColors).length > 0 ? { pointColors: { '0': sliceColors } } : {}),
      ...(series.length > 0 ? { series } : {}),
    }
    if (Object.keys(edit).length === 0) {
      onCancel()
      return
    }
    onApply(edit)
  }

  return (
    // stopPropagation: the Univer float container forwards wheel events to
    // the sheet canvas — without this the grid scrolls (dragging the chart
    // and this panel along with it) instead of the panel's own list.
    <div className="chart-editor" onWheel={(event) => event.stopPropagation()}>
      <label>
        {t('appTitle')}
        <input
          className="chart-editor-title"
          value={title}
          maxLength={255}
          onChange={(event) => setTitle(event.target.value)}
        />
      </label>
      {convertible && (
        <label>
          {t('appTypeLabel')}
          <select
            className="chart-editor-type"
            value={chartType}
            onChange={(event) => setChartType(event.target.value)}
          >
            {CHART_TYPE_OPTIONS.map((option) => (
              <option key={option.value} value={option.value}>
                {t(option.labelKey)}
              </option>
            ))}
          </select>
        </label>
      )}
      {!isPie &&
        chart.series.slice(0, 6).map((series, index) => (
          <label key={index}>
            {truncateLabel(series.name || t('appSeriesN', { n: index + 1 }), 14)}
            <input
              type="color"
              value={colors[String(index)] ?? seriesColor(series, index)}
              onChange={(event) =>
                setColors((previous) => ({ ...previous, [String(index)]: event.target.value }))
              }
            />
          </label>
        ))}
      {isPie &&
        chart.series[0]?.categories.slice(0, 12).map((category, index) => (
          <label key={`slice-${index}`}>
            {truncateLabel(category || t('appSliceN', { n: index + 1 }), 14)}
            <input
              type="color"
              value={sliceColors[String(index)] ?? pieSliceColor(chart.series[0] ?? {}, index)}
              onChange={(event) =>
                setSliceColors((previous) => ({ ...previous, [String(index)]: event.target.value }))
              }
            />
          </label>
        ))}
      {readVector &&
        chart.series.slice(0, 6).map((series, index) => (
          <div className="chart-editor-series" key={`ranges-${index}`}>
            <span className="chart-editor-series-name">
              {t('appSeriesData', {
                name: truncateLabel(series.name || t('appSeriesN', { n: index + 1 }), 20),
              })}
            </span>
            <input
              className="chart-editor-range"
              placeholder={series.valuesRef ?? t('appValuesPlaceholder')}
              title={t('appValuesRangeTitle', {
                name: series.name || t('appSeriesN', { n: index + 1 }),
              })}
              value={ranges[String(index)]?.values ?? ''}
              onChange={(event) =>
                setRanges((previous) => ({
                  ...previous,
                  [String(index)]: {
                    categories: '',
                    ...previous[String(index)],
                    values: event.target.value,
                  },
                }))
              }
            />
            <input
              className="chart-editor-range"
              placeholder={series.categoriesRef ?? t('appLabelsPlaceholder')}
              title={t('appCategoriesRangeTitle', {
                name: series.name || t('appSeriesN', { n: index + 1 }),
              })}
              value={ranges[String(index)]?.categories ?? ''}
              onChange={(event) =>
                setRanges((previous) => ({
                  ...previous,
                  [String(index)]: {
                    values: '',
                    ...previous[String(index)],
                    categories: event.target.value,
                  },
                }))
              }
            />
          </div>
        ))}
      {rangeError && <div className="chart-editor-error">{rangeError}</div>}
      <div className="chart-editor-actions">
        <button className="chart-editor-cancel" onClick={onCancel}>
          {t('appCancel')}
        </button>
        <button
          className="chart-editor-apply"
          disabled={isReading}
          onClick={() => void handleApply()}
        >
          {isReading ? t('appReading') : t('appApply')}
        </button>
      </div>
    </div>
  )
}

function SeriesLegend({
  seriesList,
  selected,
  onSelect,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly selected?: boolean | undefined
  readonly onSelect?: ((event: React.MouseEvent) => void) | undefined
}): React.JSX.Element {
  return (
    <div className={`chart-legend${selected ? ' chart-el-selected' : ''}`} onClick={onSelect}>
      {seriesList.slice(0, 8).map((series, index) => (
        <span key={`${series.name}-${index}`}>
          <i style={{ background: seriesColor(series, index) }} />
          {truncateLabel(series.name, 18)}
        </span>
      ))}
      {seriesList.length > 8 && <span>{t('appMoreSeries', { count: seriesList.length - 8 })}</span>}
    </div>
  )
}

function BarChart({
  seriesList,
  lineSeries,
  isHorizontal,
  axisTitles,
  dataLabels,
  dataLabelPosition,
  dataLabelFormat,
  grouping,
  gridlines,
  valueAxis,
  gapWidthPct,
  categoryFormat,
  onElement,
  selectedEl,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly lineSeries?: ChartSeries | undefined
  readonly isHorizontal: boolean
  readonly axisTitles?: ChartAxisTitles
  readonly dataLabels?: ChartDataLabels
  readonly dataLabelPosition?: ChartLabelPosition
  readonly dataLabelFormat?: string | undefined
  readonly grouping?: ChartGrouping
  readonly gridlines?: boolean | undefined
  readonly valueAxis?: ChartValueAxis
  readonly gapWidthPct?: number | undefined
  readonly categoryFormat?: string | undefined
} & ChartElementProps): React.JSX.Element {
  const primary = seriesList[0]
  if (!primary) return <div className="xlsx-visual-error">{t('appChartCacheEmpty')}</div>
  const categories = primary.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const visibleCount = Math.min(primary.values.length, 20)
  const isStacked =
    (grouping === 'stacked' || grouping === 'percentStacked') && seriesList.length > 1
  const isPercent = grouping === 'percentStacked' && seriesList.length > 1
  const categoryTotal = (index: number): number =>
    seriesList.reduce((sum, series) => sum + Math.max(0, series.values[index] ?? 0), 0)
  const bounds = isPercent
    ? { min: 0, max: 1 }
    : axisBounds(
        isStacked
          ? Math.max(...Array.from({ length: visibleCount }, (_, index) => categoryTotal(index)), 0)
          : Math.max(...seriesList.flatMap((series) => [...series.values]), 0),
        valueAxis,
      )
  const span = bounds.max - bounds.min
  const norm = (value: number): number => Math.max(0, Math.min(1, (value - bounds.min) / span))
  // Stacked segments share the category slot; each value scales against the
  // axis maximum (percentStacked normalizes per category first).
  const segment = (seriesIndex: number, index: number): number => {
    const value = Math.max(0, seriesList[seriesIndex]?.values[index] ?? 0)
    if (!isPercent) return Math.min(1, value / (bounds.max || 1))
    const total = categoryTotal(index)
    return total === 0 ? 0 : value / total
  }
  const axisNumberFormat = isPercent ? '0%' : primary.numberFormat
  // Excel gap width: the space between category groups, in % of one bar.
  const gap = (gapWidthPct ?? 150) / 100
  const pickBar = onElement
    ? (event: React.MouseEvent, seriesIndex: number, pointIndex: number): void => {
        event.stopPropagation()
        onElement(narrowSelection(selectedEl ?? null, seriesIndex, pointIndex))
      }
    : undefined
  const barStroke = (seriesIndex: number, pointIndex: number): Record<string, string> =>
    isSelectedPoint(selectedEl, seriesIndex, pointIndex)
      ? { stroke: '#107C41', strokeWidth: '2' }
      : {}
  const selectCategoryAxis = onElement
    ? (event: React.MouseEvent): void => {
        event.stopPropagation()
        onElement({ kind: 'category-axis' })
      }
    : undefined
  const selectValueAxis = onElement
    ? (event: React.MouseEvent): void => {
        event.stopPropagation()
        onElement({ kind: 'value-axis' })
      }
    : undefined

  if (isHorizontal) {
    const rowHeight = 280 / visibleCount
    const barHeight = Math.max(
      3,
      isStacked ? rowHeight / (1 + gap) : rowHeight / (seriesList.length + gap),
    )
    const groupTop = (index: number): number =>
      14 + rowHeight * index + (rowHeight - barHeight * (isStacked ? 1 : seriesList.length)) / 2
    return (
      <svg className="chart-svg" viewBox="0 0 600 320" role="img">
        {Array.from({ length: visibleCount }, (_, index) => {
          let cursor = 0
          return (
            <g key={`${categories[index] ?? index}-${index}`}>
              <text
                x="148"
                y={26 + rowHeight * index}
                textAnchor="end"
                onClick={selectCategoryAxis}
              >
                {truncateLabel(categories[index] ?? String(index + 1))}
              </text>
              {seriesList.map((series, seriesIndex) => {
                const share = isStacked
                  ? segment(seriesIndex, index)
                  : norm(series.values[index] ?? bounds.min)
                const x = 158 + (isStacked ? cursor * 390 : 0)
                if (isStacked) cursor += share
                return (
                  <rect
                    key={seriesIndex}
                    x={x}
                    y={groupTop(index) + (isStacked ? 0 : barHeight * seriesIndex)}
                    width={share * 390}
                    height={barHeight}
                    fill={seriesColor(series, seriesIndex)}
                    {...barStroke(seriesIndex, index)}
                    onClick={pickBar ? (event) => pickBar(event, seriesIndex, index) : undefined}
                  />
                )
              })}
              {dataLabels === 'value' &&
                !isStacked &&
                (() => {
                  const width = norm(primary.values[index] ?? bounds.min) * 390
                  const inside =
                    dataLabelPosition === 'center' || dataLabelPosition === 'inside-end'
                  const x =
                    dataLabelPosition === 'center'
                      ? 158 + width / 2
                      : dataLabelPosition === 'inside-end'
                        ? 154 + width
                        : 164 + width
                  return (
                    <text
                      x={x}
                      y={groupTop(index) + barHeight / 2 + 3}
                      textAnchor={
                        dataLabelPosition === 'center'
                          ? 'middle'
                          : dataLabelPosition === 'inside-end'
                            ? 'end'
                            : 'start'
                      }
                      className="axis-label"
                      {...(inside ? { fill: '#fff' } : {})}
                    >
                      {formatLabelValue(
                        primary.values[index] ?? 0,
                        dataLabelFormat,
                        primary.numberFormat,
                      )}
                    </text>
                  )
                })()}
            </g>
          )
        })}
        <TruncationNote shown={visibleCount} total={primary.values.length} />
        <AxisTitleTexts bottom={axisTitles?.value} left={axisTitles?.category} />
      </svg>
    )
  }

  const columnWidth = 480 / visibleCount
  const barWidth = Math.max(
    2,
    isStacked ? columnWidth / (1 + gap) : columnWidth / (seriesList.length + gap),
  )
  const groupWidth = barWidth * (isStacked ? 1 : seriesList.length)
  const groupLeft = (index: number): number =>
    62 + columnWidth * index + (columnWidth - groupWidth) / 2
  const lineMaximum = lineSeries ? niceAxisMaximum(Math.max(...lineSeries.values, 0)) : 1
  const points = lineSeries?.values
    .slice(0, visibleCount)
    .map(
      (value, index) => `${groupLeft(index) + groupWidth / 2},${280 - (value / lineMaximum) * 230}`,
    )
    .join(' ')
  const showValueLabels =
    !isStacked &&
    (dataLabels === 'value' ||
      (dataLabels === undefined && visibleCount <= 12 && seriesList.length === 1))
  return (
    <svg className="chart-svg" viewBox="0 0 600 320" role="img">
      <VerticalAxis
        minimum={bounds.min}
        maximum={bounds.max}
        numberFormat={axisNumberFormat}
        showGridlines={gridlines !== false}
        onSelect={selectValueAxis}
      />
      {Array.from({ length: visibleCount }, (_, index) => {
        let cursor = 0
        return (
          <g key={`${categories[index] ?? index}-${index}`}>
            {seriesList.map((series, seriesIndex) => {
              const share = isStacked
                ? segment(seriesIndex, index)
                : norm(series.values[index] ?? bounds.min)
              const height = share * 240
              const y = 280 - (isStacked ? (cursor + share) * 240 : height)
              if (isStacked) cursor += share
              return (
                <rect
                  key={seriesIndex}
                  x={groupLeft(index) + (isStacked ? 0 : barWidth * seriesIndex)}
                  y={y}
                  width={barWidth}
                  height={height}
                  fill={seriesColor(series, seriesIndex)}
                  {...barStroke(seriesIndex, index)}
                  onClick={pickBar ? (event) => pickBar(event, seriesIndex, index) : undefined}
                />
              )
            })}
            {showValueLabels &&
              (() => {
                const share = norm(primary.values[index] ?? bounds.min)
                const inside = dataLabelPosition === 'center' || dataLabelPosition === 'inside-end'
                const y =
                  dataLabelPosition === 'center'
                    ? 283 - share * 120
                    : dataLabelPosition === 'inside-end'
                      ? 292 - share * 240
                      : 272 - share * 240
                return (
                  <text
                    x={groupLeft(index) + groupWidth / 2}
                    y={y}
                    textAnchor="middle"
                    className="axis-label"
                    {...(inside ? { fill: '#fff' } : {})}
                  >
                    {formatLabelValue(
                      primary.values[index] ?? 0,
                      dataLabelFormat,
                      primary.numberFormat,
                    )}
                  </text>
                )
              })()}
            <text
              x={62 + columnWidth * index + columnWidth / 2}
              y="298"
              textAnchor="middle"
              onClick={selectCategoryAxis}
            >
              {truncateLabel(categories[index] ?? String(index + 1), 5)}
            </text>
          </g>
        )
      })}
      {!isStacked &&
        seriesList.map((series, seriesIndex) =>
          series.trendline === 'linear' ? (
            <TrendLine
              key={`trend-${seriesIndex}`}
              values={series.values.slice(0, visibleCount)}
              minimum={bounds.min}
              maximum={bounds.max}
              color={seriesColor(series, seriesIndex)}
              xFor={(index) => groupLeft(index) + groupWidth / 2}
            />
          ) : null,
        )}
      {points && (
        <polyline
          points={points}
          fill="none"
          stroke={seriesColor(lineSeries, seriesList.length)}
          strokeWidth="3"
        />
      )}
      {lineSeries &&
        [0, 0.5, 1].map((fraction) => (
          <text
            key={fraction}
            x="596"
            y={284 - fraction * 230}
            textAnchor="end"
            className="axis-label"
            fill={seriesColor(lineSeries, seriesList.length)}
          >
            {formatAxisValue(fraction * lineMaximum, lineSeries.numberFormat)}
          </text>
        ))}
      <TruncationNote shown={visibleCount} total={primary.values.length} />
      <AxisTitleTexts bottom={axisTitles?.category} left={axisTitles?.value} />
    </svg>
  )
}

/// Least-squares regression over (index, value), rendered as a dashed line.
function TrendLine({
  values,
  minimum = 0,
  maximum,
  color,
  xFor,
}: {
  readonly values: readonly number[]
  readonly minimum?: number
  readonly maximum: number
  readonly color: string
  readonly xFor: (index: number) => number
}): React.JSX.Element | null {
  const count = values.length
  if (count < 2) return null
  const meanX = (count - 1) / 2
  const meanY = values.reduce((sum, value) => sum + value, 0) / count
  let numerator = 0
  let denominator = 0
  for (let index = 0; index < count; index += 1) {
    numerator += (index - meanX) * ((values[index] ?? 0) - meanY)
    denominator += (index - meanX) ** 2
  }
  const slope = denominator === 0 ? 0 : numerator / denominator
  const intercept = meanY - slope * meanX
  const span = maximum - minimum || 1
  const yFor = (index: number): number =>
    280 - Math.max(0, Math.min(1, (intercept + slope * index - minimum) / span)) * 240
  return (
    <line
      x1={xFor(0)}
      y1={yFor(0)}
      x2={xFor(count - 1)}
      y2={yFor(count - 1)}
      stroke={color}
      strokeWidth="2"
      strokeDasharray="6 4"
      opacity="0.85"
    />
  )
}

function linePoints(values: readonly number[], maximum: number, minimum = 0): string {
  const count = Math.max(1, values.length - 1)
  const span = maximum - minimum || 1
  return values
    .map(
      (value, index) =>
        `${60 + (index / count) * 500},${280 - Math.max(0, Math.min(1, (value - minimum) / span)) * 240}`,
    )
    .join(' ')
}

/// 5-tick nice ceiling so the top gridline sits at or above the data maximum.
function niceAxisMaximum(maximum: number): number {
  if (maximum <= 0) return 1
  const magnitude = 10 ** Math.floor(Math.log10(maximum))
  const normalized = maximum / magnitude
  const nice =
    normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 2.5 ? 2.5 : normalized <= 5 ? 5 : 10
  return nice * magnitude
}

type ChartLabelPosition = ChartMetadata['dataLabelPosition']

/// Minimal formatCode support for data labels (percent / thousands / fixed
/// decimals); anything fancier falls back to the axis heuristics.
function formatLabelValue(
  value: number,
  formatCode: string | undefined,
  numberFormat: string | undefined,
): string {
  if (!formatCode) return formatAxisValue(value, numberFormat)
  const decimals = /0\.(0+)/.exec(formatCode)?.[1]?.length ?? 0
  if (formatCode.includes('%')) return `${(value * 100).toFixed(decimals)}%`
  const fixed = value.toFixed(decimals)
  if (!formatCode.includes(',')) return fixed
  const [whole, fraction] = fixed.split('.')
  const grouped = Number(whole).toLocaleString('en-US')
  return fraction ? `${grouped}.${fraction}` : grouped
}

/// Corner note when the drawing shows fewer categories than the data has.
function TruncationNote({
  shown,
  total,
}: {
  readonly shown: number
  readonly total: number
}): React.JSX.Element | null {
  if (total <= shown) return null
  return (
    <text x="580" y="14" textAnchor="end" className="axis-label" opacity="0.75">
      {t('appTruncationNote', { shown, total })}
    </text>
  )
}

function formatAxisValue(value: number, numberFormat: string | undefined): string {
  if (numberFormat?.includes('%')) return `${Math.round(value * 100)}%`
  const magnitude = Math.abs(value)
  if (magnitude >= 1e9) return `${(value / 1e9).toFixed(1)}B`
  if (magnitude >= 1e6) return `${(value / 1e6).toFixed(1)}M`
  if (magnitude >= 1e4) return `${Math.round(value / 1e3)}K`
  return Number.isInteger(value) ? value.toLocaleString() : value.toFixed(1)
}

function VerticalAxis({
  minimum = 0,
  maximum,
  numberFormat,
  showGridlines = true,
  onSelect,
}: {
  readonly minimum?: number
  readonly maximum: number
  readonly numberFormat: string | undefined
  readonly showGridlines?: boolean
  readonly onSelect?: ((event: React.MouseEvent) => void) | undefined
}): React.JSX.Element {
  const span = maximum - minimum || 1
  return (
    <g onClick={onSelect}>
      {[0, 0.25, 0.5, 0.75, 1].map((fraction) => {
        const tick = minimum + fraction * span
        const y = 280 - fraction * 240
        return (
          <g key={fraction}>
            {(showGridlines || fraction === 0) && (
              <line x1="58" y1={y} x2="580" y2={y} stroke="#e3e3e3" strokeWidth="1" />
            )}
            <text x="54" y={y + 4} textAnchor="end" className="axis-label">
              {formatAxisValue(tick, numberFormat)}
            </text>
          </g>
        )
      })}
    </g>
  )
}

type ChartValueAxis = ChartMetadata['valueAxis']

/// Explicit axis bounds win; otherwise 0 up to a nice ceiling.
function axisBounds(dataMax: number, valueAxis: ChartValueAxis): { min: number; max: number } {
  const min = valueAxis?.min ?? 0
  const max = valueAxis?.max ?? niceAxisMaximum(dataMax)
  return max > min ? { min, max } : { min, max: min + 1 }
}

function LineChart({
  seriesList,
  axisTitles,
  dataLabels,
  grouping,
  gridlines,
  valueAxis,
  categoryFormat,
  onElement,
  selectedEl,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly axisTitles?: ChartAxisTitles
  readonly dataLabels?: ChartDataLabels
  readonly grouping?: ChartGrouping
  readonly gridlines?: boolean | undefined
  readonly valueAxis?: ChartValueAxis
  readonly categoryFormat?: string | undefined
} & ChartElementProps): React.JSX.Element {
  const primary = seriesList[0]
  if (!primary) return <div className="xlsx-visual-error">{t('appChartCacheEmpty')}</div>
  const isStacked =
    (grouping === 'stacked' || grouping === 'percentStacked') && seriesList.length > 1
  const isPercent = grouping === 'percentStacked' && seriesList.length > 1
  // Running per-category totals; each line rides on the ones below it
  // (negatives clamp to 0, matching AreaChart).
  const stackTotals = isStacked
    ? seriesList.reduce<number[][]>((totals, series) => {
        const previous = totals[totals.length - 1] ?? primary.values.map(() => 0)
        totals.push(previous.map((base, index) => base + Math.max(0, series.values[index] ?? 0)))
        return totals
      }, [])
    : []
  const percentTotals = isPercent
    ? primary.values.map(
        (_, index) =>
          seriesList.reduce((sum, series) => sum + Math.max(0, series.values[index] ?? 0), 0) || 1,
      )
    : []
  const displayValues = (seriesIndex: number): number[] => {
    if (!isStacked) return [...(seriesList[seriesIndex]?.values ?? [])]
    const totals = stackTotals[seriesIndex] ?? []
    return isPercent ? totals.map((value, index) => value / (percentTotals[index] ?? 1)) : totals
  }
  const bounds = isPercent
    ? { min: 0, max: 1 }
    : axisBounds(
        isStacked
          ? Math.max(...(stackTotals[stackTotals.length - 1] ?? [0]), 0)
          : Math.max(...seriesList.flatMap((series) => [...series.values]), 0),
        isStacked ? undefined : valueAxis,
      )
  const span = bounds.max - bounds.min
  const categories = primary.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const count = Math.max(1, primary.values.length - 1)
  return (
    <svg className="chart-svg" viewBox="0 0 600 320" role="img">
      <VerticalAxis
        minimum={bounds.min}
        maximum={bounds.max}
        numberFormat={isPercent ? '0%' : primary.numberFormat}
        showGridlines={gridlines !== false}
        onSelect={
          onElement
            ? (event) => {
                event.stopPropagation()
                onElement({ kind: 'value-axis' })
              }
            : undefined
        }
      />
      {seriesList.map((series, seriesIndex) => (
        <polyline
          key={seriesIndex}
          points={linePoints(displayValues(seriesIndex), bounds.max, bounds.min)}
          fill="none"
          stroke={seriesColor(series, seriesIndex)}
          strokeWidth={
            selectedEl?.kind === 'series' && selectedEl.seriesIndex === seriesIndex ? '5' : '3'
          }
          onClick={
            onElement
              ? (event) => {
                  event.stopPropagation()
                  onElement({ kind: 'series', seriesIndex })
                }
              : undefined
          }
        />
      ))}
      {seriesList.map((series, seriesIndex) =>
        series.trendline === 'linear' ? (
          <TrendLine
            key={`trend-${seriesIndex}`}
            values={displayValues(seriesIndex)}
            minimum={bounds.min}
            maximum={bounds.max}
            color={seriesColor(series, seriesIndex)}
            xFor={(index) => 60 + (index / Math.max(1, series.values.length - 1)) * 500}
          />
        ) : null,
      )}
      {primary.values.map((_, index) => (
        <text
          key={index}
          x={60 + (index / count) * 500}
          y="300"
          textAnchor="middle"
          onClick={
            onElement
              ? (event) => {
                  event.stopPropagation()
                  onElement({ kind: 'category-axis' })
                }
              : undefined
          }
        >
          {truncateLabel(categories[index] ?? String(index + 1), 5)}
        </text>
      ))}
      {dataLabels === 'value' &&
        displayValues(0).map((displayed, index) => (
          <text
            key={`label-${index}`}
            x={60 + (index / count) * 500}
            y={272 - Math.max(0, Math.min(1, (displayed - bounds.min) / span)) * 240}
            textAnchor="middle"
            className="axis-label"
          >
            {formatAxisValue(primary.values[index] ?? 0, primary.numberFormat)}
          </text>
        ))}
      <AxisTitleTexts bottom={axisTitles?.category} left={axisTitles?.value} />
    </svg>
  )
}

function AreaChart({
  seriesList,
  axisTitles,
  dataLabels,
  grouping,
  gridlines,
  valueAxis,
  categoryFormat,
  onElement,
  selectedEl,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly axisTitles?: ChartAxisTitles
  readonly dataLabels?: ChartDataLabels
  readonly grouping?: ChartGrouping
  readonly gridlines?: boolean | undefined
  readonly valueAxis?: ChartValueAxis
  readonly categoryFormat?: string | undefined
} & ChartElementProps): React.JSX.Element {
  const primary = seriesList[0]
  if (!primary) return <div className="xlsx-visual-error">{t('appChartCacheEmpty')}</div>
  const categories = primary.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const count = Math.max(1, primary.values.length - 1)
  const isStacked =
    (grouping === 'stacked' || grouping === 'percentStacked') && seriesList.length > 1
  const isPercent = grouping === 'percentStacked' && seriesList.length > 1
  // Cumulative per-category boundaries: series N fills between the running
  // total below it and the one including it.
  const stackBounds = isStacked
    ? seriesList.reduce<number[][]>((bounds, series) => {
        const previous = bounds[bounds.length - 1] ?? primary.values.map(() => 0)
        bounds.push(previous.map((base, index) => base + Math.max(0, series.values[index] ?? 0)))
        return bounds
      }, [])
    : []
  const percentTotals = isPercent
    ? primary.values.map(
        (_, index) =>
          seriesList.reduce((sum, series) => sum + Math.max(0, series.values[index] ?? 0), 0) || 1,
      )
    : []
  const bounds = isPercent
    ? { min: 0, max: 1 }
    : axisBounds(
        isStacked
          ? Math.max(...(stackBounds[stackBounds.length - 1] ?? [0]), 0)
          : Math.max(...seriesList.flatMap((series) => [...series.values]), 0),
        isStacked ? undefined : valueAxis,
      )
  const axisSpan = bounds.max - bounds.min
  const stackY = (raw: number, index: number): number =>
    280 - (isPercent ? raw / (percentTotals[index] ?? 1) : raw / (bounds.max || 1)) * 240
  const selectSeries = onElement
    ? (event: React.MouseEvent, seriesIndex: number): void => {
        event.stopPropagation()
        onElement({ kind: 'series', seriesIndex })
      }
    : undefined
  return (
    <svg className="chart-svg" viewBox="0 0 600 320" role="img">
      <VerticalAxis
        minimum={bounds.min}
        maximum={bounds.max}
        numberFormat={isPercent ? '0%' : primary.numberFormat}
        showGridlines={gridlines !== false}
        onSelect={
          onElement
            ? (event) => {
                event.stopPropagation()
                onElement({ kind: 'value-axis' })
              }
            : undefined
        }
      />
      {seriesList.map((series, seriesIndex) => {
        const isSel = selectedEl?.kind === 'series' && selectedEl.seriesIndex === seriesIndex
        if (isStacked) {
          const upper = stackBounds[seriesIndex] ?? []
          const lower =
            seriesIndex === 0 ? upper.map(() => 0) : (stackBounds[seriesIndex - 1] ?? [])
          const xFor = (index: number): number => 60 + (index / count) * 500
          const upperPoints = upper.map((value, index) => `${xFor(index)},${stackY(value, index)}`)
          const lowerPoints = lower
            .map((value, index) => `${xFor(index)},${stackY(value, index)}`)
            .reverse()
          return (
            <g
              key={seriesIndex}
              onClick={selectSeries ? (event) => selectSeries(event, seriesIndex) : undefined}
            >
              <polygon
                points={[...upperPoints, ...lowerPoints].join(' ')}
                fill={seriesColor(series, seriesIndex)}
                opacity="0.75"
              />
              <polyline
                points={upperPoints.join(' ')}
                fill="none"
                stroke={isSel ? '#107C41' : seriesColor(series, seriesIndex)}
                strokeWidth={isSel ? '3.5' : '2'}
              />
            </g>
          )
        }
        const points = linePoints(series.values, bounds.max, bounds.min)
        const lastX = 60 + (Math.max(0, series.values.length - 1) / count) * 500
        return (
          <g
            key={seriesIndex}
            onClick={selectSeries ? (event) => selectSeries(event, seriesIndex) : undefined}
          >
            <polygon
              points={`60,280 ${points} ${lastX},280`}
              fill={seriesColor(series, seriesIndex)}
              opacity="0.35"
            />
            <polyline
              points={points}
              fill="none"
              stroke={isSel ? '#107C41' : seriesColor(series, seriesIndex)}
              strokeWidth={isSel ? '4' : '2.5'}
            />
          </g>
        )
      })}
      {primary.values.map((_, index) => (
        <text key={index} x={60 + (index / count) * 500} y="300" textAnchor="middle">
          {truncateLabel(categories[index] ?? String(index + 1), 5)}
        </text>
      ))}
      {dataLabels === 'value' &&
        primary.values.map((value, index) => (
          <text
            key={`label-${index}`}
            x={60 + (index / count) * 500}
            y={272 - Math.max(0, Math.min(1, (value - bounds.min) / axisSpan)) * 240}
            textAnchor="middle"
            className="axis-label"
          >
            {formatAxisValue(value, primary.numberFormat)}
          </text>
        ))}
      <AxisTitleTexts bottom={axisTitles?.category} left={axisTitles?.value} />
    </svg>
  )
}

/// Spider chart: one spoke per category, rings at quarter steps, one closed
/// polygon per series.
function RadarChart({
  seriesList,
  dataLabels,
  categoryFormat,
  onElement,
  selectedEl,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly dataLabels?: ChartDataLabels
  readonly categoryFormat?: string | undefined
} & ChartElementProps): React.JSX.Element {
  const primary = seriesList[0]
  if (!primary) return <div className="xlsx-visual-error">{t('appChartCacheEmpty')}</div>
  const categories = primary.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const count = Math.min(Math.max(primary.values.length, 3), 12)
  const cx = 300
  const cy = 168
  const r = 118
  const maximum = niceAxisMaximum(
    Math.max(...seriesList.flatMap((series) => [...series.values]), 0),
  )
  const vertex = (index: number, radius: number): [number, number] => {
    const theta = (index / count) * 2 * Math.PI
    return [cx + Math.sin(theta) * radius, cy - Math.cos(theta) * radius]
  }
  const ringPoints = (fraction: number): string =>
    Array.from({ length: count }, (_, index) => vertex(index, r * fraction).join(',')).join(' ')
  const seriesPoints = (series: ChartSeries): string =>
    Array.from({ length: count }, (_, index) => {
      const value = Math.max(0, series.values[index] ?? 0)
      return vertex(index, Math.min(1, value / maximum) * r).join(',')
    }).join(' ')
  return (
    <svg className="chart-svg" viewBox="0 0 600 320" role="img">
      {[0.25, 0.5, 0.75, 1].map((fraction) => (
        <polygon key={fraction} points={ringPoints(fraction)} fill="none" stroke="#e3e3e3" />
      ))}
      {Array.from({ length: count }, (_, index) => {
        const [x, y] = vertex(index, r)
        const [lx, ly] = vertex(index, r * 1.12)
        return (
          <g key={index}>
            <line x1={cx} y1={cy} x2={x} y2={y} stroke="#e3e3e3" />
            <text
              x={lx}
              y={ly + 3}
              textAnchor={Math.abs(lx - cx) < 8 ? 'middle' : lx > cx ? 'start' : 'end'}
              onClick={
                onElement
                  ? (event) => {
                      event.stopPropagation()
                      onElement({ kind: 'category-axis' })
                    }
                  : undefined
              }
            >
              {truncateLabel(categories[index] ?? String(index + 1), 8)}
            </text>
          </g>
        )
      })}
      {seriesList.map((series, seriesIndex) => {
        const isSel = selectedEl?.kind === 'series' && selectedEl.seriesIndex === seriesIndex
        return (
          <polygon
            key={seriesIndex}
            points={seriesPoints(series)}
            fill={seriesColor(series, seriesIndex)}
            fillOpacity="0.18"
            stroke={isSel ? '#107C41' : seriesColor(series, seriesIndex)}
            strokeWidth={isSel ? '4' : '2.5'}
            onClick={
              onElement
                ? (event) => {
                    event.stopPropagation()
                    onElement({ kind: 'series', seriesIndex })
                  }
                : undefined
            }
          />
        )
      })}
      {dataLabels === 'value' &&
        Array.from({ length: count }, (_, index) => {
          const value = primary.values[index] ?? 0
          const [x, y] = vertex(index, Math.min(1, Math.max(0, value) / maximum) * r + 10)
          return (
            <text key={index} x={x} y={y} textAnchor="middle" className="axis-label">
              {formatAxisValue(value, primary.numberFormat)}
            </text>
          )
        })}
    </svg>
  )
}

/// Scatter tick labels keep sub-integer precision that formatAxisValue's
/// one-decimal rounding would collapse (0.04 → "0.0").
function formatScatterTick(value: number, format: string | undefined): string {
  if (format !== undefined && format !== 'General') {
    return formatCategoryLabel(String(value), format)
  }
  if (Number.isInteger(value) || Math.abs(value) >= 1e4) return formatAxisValue(value, undefined)
  return String(Number(value.toPrecision(4)))
}

function ScatterChart({
  seriesList,
  axisTitles,
  dataLabels,
  gridlines,
  valueAxis,
  categoryFormat,
  onElement,
  selectedEl,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly axisTitles?: ChartAxisTitles
  readonly dataLabels?: ChartDataLabels
  readonly gridlines?: boolean | undefined
  readonly valueAxis?: ChartValueAxis
  readonly categoryFormat?: string | undefined
} & ChartElementProps): React.JSX.Element {
  const points = seriesList.map((series) => {
    const xValues = series.categories.map((value, index) => {
      const parsed = Number.parseFloat(value)
      return Number.isFinite(parsed) ? parsed : index
    })
    return { series, xValues }
  })
  const allX = points.flatMap((entry) => entry.xValues)
  const allY = points.flatMap((entry) => [...entry.series.values])
  if (allY.length === 0) return <div className="xlsx-visual-error">{t('appChartCacheEmpty')}</div>
  // X has no explicit bounds channel (the sidecar reports the left axis'
  // c:scaling as valueAxis); it always auto-scales to the data.
  const boundsX = scatterAxisBounds(allX)
  const boundsY = scatterAxisBounds(allY, valueAxis)
  const plotX = (value: number): number =>
    60 + Math.max(0, Math.min(1, (value - boundsX.min) / (boundsX.max - boundsX.min))) * 520
  const plotY = (value: number): number =>
    280 - Math.max(0, Math.min(1, (value - boundsY.min) / (boundsY.max - boundsY.min))) * 240
  return (
    <svg className="chart-svg" viewBox="0 0 600 320" role="img">
      <VerticalAxis
        minimum={boundsY.min}
        maximum={boundsY.max}
        numberFormat={seriesList[0]?.numberFormat}
        showGridlines={gridlines !== false}
        onSelect={
          onElement
            ? (event) => {
                event.stopPropagation()
                onElement({ kind: 'value-axis' })
              }
            : undefined
        }
      />
      {boundsX.ticks.map((tick, index) => (
        <text
          key={index}
          x={plotX(tick)}
          y="296"
          textAnchor="middle"
          className="axis-label"
          onClick={
            onElement
              ? (event) => {
                  event.stopPropagation()
                  onElement({ kind: 'category-axis' })
                }
              : undefined
          }
        >
          {formatScatterTick(tick, categoryFormat)}
        </text>
      ))}
      {points.map(({ series, xValues }, seriesIndex) => (
        <g key={seriesIndex}>
          {series.values.map((value, index) => (
            <circle
              key={index}
              cx={plotX(xValues[index] ?? 0)}
              cy={plotY(value)}
              r={isSelectedPoint(selectedEl, seriesIndex, index) ? 6 : 4}
              fill={seriesColor(series, seriesIndex)}
              opacity="0.85"
              {...(isSelectedPoint(selectedEl, seriesIndex, index)
                ? { stroke: '#107C41', strokeWidth: 2 }
                : {})}
              onClick={
                onElement
                  ? (event) => {
                      event.stopPropagation()
                      onElement(narrowSelection(selectedEl ?? null, seriesIndex, index))
                    }
                  : undefined
              }
            />
          ))}
          {dataLabels === 'value' &&
            series.values.map((value, index) => (
              <text
                key={`label-${index}`}
                x={plotX(xValues[index] ?? 0)}
                y={plotY(value) - 8}
                textAnchor="middle"
                className="axis-label"
              >
                {formatAxisValue(value, series.numberFormat)}
              </text>
            ))}
        </g>
      ))}
      <AxisTitleTexts bottom={axisTitles?.category} left={axisTitles?.value} />
    </svg>
  )
}

interface PieSliceLabel {
  key: number
  inside: boolean
  x: number
  y: number
  anchor: 'start' | 'middle' | 'end'
  lines: string[]
  leader?: string
}

/// Label placement: slices big enough hold their label at
/// two-thirds radius; small slices move it outside with a leader line, and
/// outside labels on each side push apart until they no longer overlap.
function pieSliceLabels(
  values: readonly number[],
  categories: readonly string[],
  total: number,
  mode: NonNullable<ChartDataLabels>,
  geometry: { cx: number; cy: number; r: number; inner: number; offsets: readonly number[] },
  position?: ChartLabelPosition,
  formatCode?: string | undefined,
): PieSliceLabel[] {
  const { cx, cy, r, inner, offsets } = geometry
  const labels: PieSliceLabel[] = []
  let cursor = 0
  for (const [index, value] of values.entries()) {
    const share = Math.max(0, value) / total
    const mid = (cursor + share / 2) * 2 * Math.PI
    cursor += share
    if (share < 0.02) continue
    const percentText = `${(share * 100).toFixed(share >= 0.1 ? 0 : 1)}%`
    const lines =
      mode === 'category-percent'
        ? [truncateLabel(categories[index] ?? '', 12), percentText]
        : [mode === 'value' ? formatLabelValue(value, formatCode, undefined) : percentText]
    const sin = Math.sin(mid)
    const cos = Math.cos(mid)
    // An exploded slice carries its label out with it.
    const offset = offsets[index] ?? 0
    // An explicit position overrides the best-fit heuristic.
    const inside =
      position === 'outside-end' ? false : position !== undefined ? true : share >= 0.09
    if (inside) {
      const baseRadius =
        inner > 0
          ? (r + inner) / 2
          : position === 'inside-end'
            ? r * 0.8
            : position === 'center'
              ? r * 0.5
              : r * 0.62
      const radius = baseRadius + offset
      labels.push({
        key: index,
        inside: true,
        x: cx + sin * radius,
        y: cy - cos * radius,
        anchor: 'middle',
        lines,
      })
      continue
    }
    const elbowR = r * 1.1 + offset
    const tick = 10
    const rightSide = sin >= 0
    const elbowX = cx + sin * elbowR
    const elbowY = cy - cos * elbowR
    const textX = elbowX + (rightSide ? tick : -tick)
    labels.push({
      key: index,
      inside: false,
      x: textX + (rightSide ? 3 : -3),
      y: elbowY,
      anchor: rightSide ? 'start' : 'end',
      lines,
      leader: `${cx + sin * (r + offset)},${cy - cos * (r + offset)} ${elbowX},${elbowY} ${textX},${elbowY}`,
    })
  }
  // Outside labels: per side, top-down, keep a minimum vertical gap.
  for (const side of ['start', 'end'] as const) {
    const outside = labels
      .filter((label) => !label.inside && label.anchor === side)
      .sort((a, b) => a.y - b.y)
    const gap = 15
    for (let i = 1; i < outside.length; i += 1) {
      const previous = outside[i - 1]
      const current = outside[i]
      if (previous && current && current.y < previous.y + gap) current.y = previous.y + gap
    }
  }
  return labels
}

function PieChart({
  series,
  isDoughnut,
  legend,
  dataLabels,
  dataLabelPosition,
  dataLabelFormat,
  holeSizePct,
  categoryFormat,
  onElement,
  selectedEl,
}: {
  readonly series: ChartSeries
  readonly isDoughnut: boolean
  readonly legend?: ChartMetadata['legend']
  readonly dataLabels?: ChartDataLabels
  readonly dataLabelPosition?: ChartLabelPosition
  readonly dataLabelFormat?: string | undefined
  readonly holeSizePct?: number | undefined
  readonly categoryFormat?: string | undefined
} & ChartElementProps): React.JSX.Element {
  const { values } = series
  const categories = series.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const total = values.reduce((sum, value) => sum + Math.max(0, value), 0) || 1
  const cx = 210
  const cy = 160
  const r = 100
  // 50 mirrors what buildChartXml writes for new doughnuts, so the render
  // does not jump after the first save.
  const inner = isDoughnut ? (r * Math.min(90, Math.max(10, holeSizePct ?? 50))) / 100 : 0
  const point = (radius: number, theta: number): [number, number] => [
    cx + Math.sin(theta) * radius,
    cy - Math.cos(theta) * radius,
  ]
  // Explosion: per-slice dPt override, else the series-level value; render
  // caps the shift at 40% of the radius so labels stay inside the viewbox.
  const explosionOffset = (index: number): number => {
    const pct =
      series.pointExplosions?.find((entry) => entry.index === index)?.pct ??
      series.explosionPct ??
      0
    return (Math.min(100, Math.max(0, pct)) / 100) * r * 0.4
  }
  const offsets = values.map((_, index) => explosionOffset(index))
  let angle = 0
  const slices = values
    .map((value, index) => {
      const share = Math.max(0, value) / total
      const start = angle
      angle += share * 2 * Math.PI
      return { index, share, start, end: angle, color: pieSliceColor(series, index) }
    })
    .filter((slice) => slice.share > 0.0005)
  const slicePath = (slice: (typeof slices)[number]): string => {
    // A lone full-circle slice has coincident endpoints; nudge the end angle
    // so the arc still spans the whole ring.
    const end = slice.share > 0.9995 ? slice.start + 2 * Math.PI - 0.001 : slice.end
    const large = end - slice.start > Math.PI ? 1 : 0
    const [x1, y1] = point(r, slice.start)
    const [x2, y2] = point(r, end)
    if (inner === 0) {
      return `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} Z`
    }
    const [ix1, iy1] = point(inner, slice.start)
    const [ix2, iy2] = point(inner, end)
    return (
      `M ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2}` +
      ` L ${ix2} ${iy2} A ${inner} ${inner} 0 ${large} 0 ${ix1} ${iy1} Z`
    )
  }
  const labels =
    dataLabels !== undefined && dataLabels !== 'none'
      ? pieSliceLabels(
          values,
          categories,
          total,
          dataLabels,
          { cx, cy, r, inner, offsets },
          dataLabelPosition,
          dataLabelFormat,
        )
      : []
  return (
    <div className={`pie-layout pie-legend-${legend ?? 'right'}`}>
      <svg className="chart-svg pie-svg" viewBox="0 0 420 320" role="img">
        {slices.map((slice) => {
          const mid = (slice.start + slice.end) / 2
          const offset = offsets[slice.index] ?? 0
          const selected = isSelectedPoint(selectedEl, 0, slice.index)
          return (
            <path
              key={slice.index}
              d={slicePath(slice)}
              transform={
                offset > 0
                  ? `translate(${Math.sin(mid) * offset} ${-Math.cos(mid) * offset})`
                  : undefined
              }
              fill={slice.color}
              stroke={selected ? '#107C41' : '#fff'}
              strokeWidth={selected ? 2.5 : 1.5}
              onClick={
                onElement
                  ? (event) => {
                      event.stopPropagation()
                      onElement({ kind: 'point', seriesIndex: 0, pointIndex: slice.index })
                    }
                  : undefined
              }
            />
          )
        })}
        {labels.map((label) => (
          <g key={label.key}>
            {label.leader && (
              <polyline points={label.leader} fill="none" stroke="#9a9a9a" strokeWidth="1" />
            )}
            <text
              x={label.x}
              y={label.y + (label.lines.length > 1 ? -2 : 3)}
              textAnchor={label.anchor}
              className={label.inside ? 'pie-label-inside' : 'pie-label-outside'}
            >
              {label.lines.map((line, lineIndex) => (
                <tspan key={lineIndex} x={label.x} dy={lineIndex === 0 ? 0 : 12}>
                  {line}
                </tspan>
              ))}
            </text>
          </g>
        ))}
      </svg>
      {legend !== 'none' && (
        <div
          className={`chart-legend${selectedEl?.kind === 'legend' ? ' chart-el-selected' : ''}`}
          onClick={
            onElement
              ? (event) => {
                  event.stopPropagation()
                  onElement({ kind: 'legend' })
                }
              : undefined
          }
        >
          {categories.slice(0, 12).map((category, index) => {
            const share = (Math.max(0, values[index] ?? 0) / total) * 100
            return (
              <span key={`${category}-${index}`}>
                <i style={{ background: pieSliceColor(series, index) }} />
                {truncateLabel(category, 12)} {share >= 0.05 ? `${share.toFixed(1)}%` : ''}
              </span>
            )
          })}
          {categories.length > 12 && (
            <span>{t('appMoreItems', { count: categories.length - 12 })}</span>
          )}
        </div>
      )}
    </div>
  )
}

function truncateLabel(value: string, maximumLength = 14): string {
  return value.length > maximumLength ? `${value.slice(0, maximumLength)}…` : value
}
