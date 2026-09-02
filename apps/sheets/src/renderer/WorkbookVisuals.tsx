import { Fragment, useEffect, useId, useRef, useState, useSyncExternalStore } from 'react'
import { createPortal } from 'react-dom'
import { BooleanNumber, numfmt } from '@univerjs/core'
import { isMetafileMime, metafileToDataUrl } from '@genoffice/docx-engine/metafile'
import { Dropdown, shapePreviewPath, useDismissablePopover } from '@genoffice/ui'

import type { createUniver } from './create-univer'

import {
  applyChartStateEdit,
  chartCategoryFormat,
  chartSupportsDataLabels,
  formatCategoryLabel,
  scatterAxisBounds,
  splitSheetRef,
  valueAxisScale,
} from '../domain/chart-visual'
import { parseAddress } from '../domain/cell-address'
import { ColorDropdown } from './ColorDropdown'
import { t } from './i18n/locale'
import { oleCaption, oleFrameStyle, oleRenderKind } from './ole-visual'
import { VisualDeleteButton } from './VisualDeleteButton'
import { shouldShowVisualDeleteButton } from './visual-delete-button'
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
  /// New a:xfrm ext in EMU — rides along when resizing a rotated shape,
  /// whose anchor stores the rotated AABB rather than the true frame.
  readonly frameSize?: { readonly width: number; readonly height: number }
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
    let fromRow = Math.min(visual.anchor.fromRow, maxRows - 1)
    let fromColumn = Math.min(visual.anchor.fromColumn, maxColumns - 1)
    let toRow = Math.min(visual.anchor.toRow, maxRows - 1)
    let toColumn = Math.min(visual.anchor.toColumn, maxColumns - 1)
    // Pixel-exact frame (Excel twoCellAnchor semantics): each marker is a
    // cell plus an EMU offset inside it, so the box runs from the `from`
    // marker to the `to` marker — not across the whole cell range.
    const columnWidth = (index: number): number => Math.max(worksheet.getColumnWidth(index), 1)
    const rowHeight = (index: number): number => Math.max(worksheet.getRowHeight(index), 1)
    let marginX = Math.max(0, visual.anchor.fromColumnOffset / EMU_PER_PIXEL)
    let marginY = Math.max(0, visual.anchor.fromRowOffset / EMU_PER_PIXEL)
    const explicitTo = visual.anchor.explicitTo === true
    let width = markerSpan(
      { index: fromColumn, offset: marginX },
      clampExplicitTo(markerFrom(toColumn, visual.anchor.toColumnOffset), explicitTo, columnWidth),
      columnWidth,
    )
    let height = markerSpan(
      { index: fromRow, offset: marginY },
      clampExplicitTo(markerFrom(toRow, visual.anchor.toRowOffset), explicitTo, rowHeight),
      rowHeight,
    )
    // A rotated shape's anchor holds its rotated bounds while xfrm ext keeps
    // the true frame; both share the same center. The float container clips
    // (Univer overflow-hidden), so re-anchor it to the AABB of the rotated
    // true frame — the component centers the ext box inside and rotates it.
    if (visual.rotation && visual.frameWidth && visual.frameHeight && width > 0 && height > 0) {
      const extWidth = visual.frameWidth / EMU_PER_PIXEL
      const extHeight = visual.frameHeight / EMU_PER_PIXEL
      const radians = (visual.rotation * Math.PI) / 180
      const aabbWidth =
        Math.abs(extWidth * Math.cos(radians)) + Math.abs(extHeight * Math.sin(radians))
      const aabbHeight =
        Math.abs(extWidth * Math.sin(radians)) + Math.abs(extHeight * Math.cos(radians))
      const fromX = walkMarker(
        { index: fromColumn, offset: marginX },
        (width - aabbWidth) / 2,
        columnWidth,
        maxColumns - 1,
      )
      const fromY = walkMarker(
        { index: fromRow, offset: marginY },
        (height - aabbHeight) / 2,
        rowHeight,
        maxRows - 1,
      )
      fromColumn = fromX.index
      fromRow = fromY.index
      marginX = fromX.offset
      marginY = fromY.offset
      toColumn = walkMarker(fromX, aabbWidth, columnWidth, maxColumns - 1).index
      toRow = walkMarker(fromY, aabbHeight, rowHeight, maxRows - 1).index
      width = aabbWidth
      height = aabbHeight
    }
    const range = worksheet.getRange(
      fromRow,
      fromColumn,
      Math.max(1, toRow + 1 - fromRow),
      Math.max(1, toColumn + 1 - fromColumn),
    )
    // Degenerate anchors (a zero marker span on BOTH axes): keep the legacy
    // behavior of filling the clamped cell range. Any real span stays a
    // marker frame with a 1px floor — tiny icons and hairlines (a line's
    // thin axis spans 0) must not balloon to their cell range, which turned
    // zero-height rules into cell-tall diagonals and icon parts into
    // cell-sized blobs.
    const framed = width > 0 || height > 0
    const frameWidth = Math.max(width, 1)
    const frameHeight = Math.max(height, 1)
    // RTL sheets mirror the float too (Excel keeps logical anchors; the box
    // lands mirrored). Univer positions the DOM from the anchor cell's
    // mirrored left edge, so restate the margin as "mirrored box left minus
    // that edge": colWidth(from) - marginX - width.
    const rtl = worksheet.getSheet().getConfig().rightToLeft === BooleanNumber.TRUE
    const anchoredMarginX = rtl ? columnWidth(fromColumn) - marginX - frameWidth : marginX
    const layout = framed
      ? { width: frameWidth, height: frameHeight, marginX: anchoredMarginX, marginY }
      : {}
    const frame = framed ? { width: frameWidth, height: frameHeight } : undefined
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
              frame={frame}
            />
          )
        : () => (
            <WorkbookVisual file={file} visual={visual} chartEditing={chartEditing} frame={frame} />
          )
    disposables.push(runtime.univerAPI.registerComponent(componentKey, component))
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

export interface CellImageState {
  readonly id: string
  readonly row: number
  readonly column: number
}

/// In-cell rich-value pictures ("place picture in cell"): one pass-through
/// float DOM per host cell, bytes fetched through the same media IPC as
/// floating images.
export function installCellImages(
  runtime: UniverRuntime,
  sessionId: string,
  images: readonly CellImageState[],
  sheetId: string,
): Disposable[] {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getSheetBySheetId(sheetId)
  if (!workbook || !worksheet) return []
  const disposables: Disposable[] = []
  for (const image of images) {
    if (image.row >= worksheet.getMaxRows() || image.column >= worksheet.getMaxColumns()) continue
    const componentKey = `cell-image-${sessionId}-${sheetId}-${image.id}`
    disposables.push(
      runtime.univerAPI.registerComponent(componentKey, () => (
        <CellImage sessionId={sessionId} imageId={image.id} />
      )),
    )
    const range = worksheet.getRange(image.row, image.column, 1, 1)
    const floating = worksheet.addFloatDomToRange(
      range,
      { componentKey, allowTransform: false, eventPassThrough: true },
      {},
      componentKey,
    )
    if (floating) disposables.push(floating)
  }
  return disposables
}

/// Excel fits the picture inside the cell box, preserving its aspect ratio;
/// while the bytes load the cell just stays blank.
function CellImage({
  sessionId,
  imageId,
}: {
  readonly sessionId: string
  readonly imageId: string
}): React.JSX.Element | null {
  const { url, unavailable } = useWorkbookMediaUrl(sessionId, imageId)
  if (unavailable || !url) return null
  return (
    <div className="xlsx-cell-image">
      <img src={url} alt="" />
    </div>
  )
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

/// Anchor-frame pixel size at install time — the shape renderer needs the
/// aspect ratio for preset geometry and Excel's rotated-anchor un-swap.
export interface ShapeFrame {
  readonly width: number
  readonly height: number
}

function WorkbookVisual({
  file,
  visual,
  chartEditing,
  frame,
}: {
  readonly file: VisualHost
  readonly visual: WorkbookVisualObject
  readonly chartEditing?: ChartEditing | undefined
  readonly frame?: ShapeFrame | undefined
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
  if (visual.kind === 'ole') {
    return <OleVisual file={file} visual={visual} />
  }
  return <ShapeVisual file={file} visual={visual} frame={frame} />
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
export type ResizeCorner = (typeof RESIZE_CORNERS)[number]

/// RTL mirrors the x-axis: a screen-space handle grabs the opposite logical
/// edge (pair with negating the screen dx).
export const mirrorCornerX = (corner: ResizeCorner): ResizeCorner =>
  (({ nw: 'ne', n: 'n', ne: 'nw', e: 'w', se: 'sw', s: 's', sw: 'se', w: 'e' }) as const)[corner]

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

/// A real `<xdr:to>` marker never reaches past its own cell: Excel clamps
/// the offset at the cell edge (broken writers store the picture's full
/// size there, e.g. ClosedXML). Synthesized to markers (oneCellAnchor ext,
/// absoluteAnchor, group children) encode sizes as offsets past the edge
/// and must keep walking, so only explicit-to anchors clamp.
export function clampExplicitTo(
  to: AnchorMarker,
  explicitTo: boolean,
  sizeOf: (index: number) => number,
): AnchorMarker {
  if (!explicitTo || to.offset <= sizeOf(to.index)) return to
  return { index: to.index, offset: sizeOf(to.index) }
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
  frame,
}: {
  readonly file: VisualHost
  readonly visual: WorkbookVisualObject
  readonly worksheet: UniverWorksheet
  readonly allowText: boolean
  readonly chartEditing?: ChartEditing | undefined
  readonly onEdit: (visualId: string, changes: ShapeEditChanges) => void
  readonly frame?: ShapeFrame | undefined
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
    screenCorner: ResizeCorner,
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
    // Drag from the clamped geometry the user sees, not the raw overflowing
    // offsets — otherwise walkMarker re-expands an explicit-to anchor to its
    // walked size on commit. The rebuilt anchor is fully normalized, so it
    // rightly drops the explicitTo flag.
    const explicitTo = anchor.explicitTo === true
    let toX = clampExplicitTo(
      markerFrom(anchor.toColumn, anchor.toColumnOffset),
      explicitTo,
      columnWidth,
    )
    let toY = clampExplicitTo(markerFrom(anchor.toRow, anchor.toRowOffset), explicitTo, rowHeight)
    // RTL sheets render mirrored geometry from logical anchors, so a screen
    // drag maps to the opposite logical x-shift and the opposite logical edge.
    const rtl = config.rightToLeft === BooleanNumber.TRUE
    const corner = rtl ? mirrorCornerX(screenCorner) : screenCorner
    const dx = (rtl ? -dxRaw : dxRaw) / zoom
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
    let frameSize: { width: number; height: number } | undefined
    if (mode === 'resize' && visual.rotation && visual.frameWidth && visual.frameHeight) {
      // The dragged box is the shape's new rotated AABB; the file's true
      // frame (xfrm ext) must change with it or the next install re-centers
      // the old AABB and the resize snaps back. Solve the ext whose rotated
      // AABB is the box; near 45° the system is singular, so scale the old
      // ext uniformly to the box's half-perimeter instead.
      const boxWidth = markerSpan(fromX, toX, columnWidth)
      const boxHeight = markerSpan(fromY, toY, rowHeight)
      const radians = (visual.rotation * Math.PI) / 180
      const cos = Math.abs(Math.cos(radians))
      const sin = Math.abs(Math.sin(radians))
      const det = cos * cos - sin * sin
      let extWidth = Math.abs(det) > 0.1 ? (cos * boxWidth - sin * boxHeight) / det : 0
      let extHeight = Math.abs(det) > 0.1 ? (cos * boxHeight - sin * boxWidth) / det : 0
      if (extWidth < MIN_FRAME_PIXELS || extHeight < MIN_FRAME_PIXELS) {
        const oldWidth = visual.frameWidth / EMU_PER_PIXEL
        const oldHeight = visual.frameHeight / EMU_PER_PIXEL
        const scale = Math.max(
          (boxWidth + boxHeight) / ((cos + sin) * (oldWidth + oldHeight)),
          MIN_FRAME_PIXELS / Math.min(oldWidth, oldHeight),
        )
        extWidth = oldWidth * scale
        extHeight = oldHeight * scale
      }
      // Keep the file invariant (anchor = the ext's rotated AABB, shared
      // center): re-derive the markers from the box center.
      const aabbWidth = cos * extWidth + sin * extHeight
      const aabbHeight = sin * extWidth + cos * extHeight
      fromX = walkMarker(fromX, (boxWidth - aabbWidth) / 2, columnWidth, maxColumn)
      fromY = walkMarker(fromY, (boxHeight - aabbHeight) / 2, rowHeight, maxRow)
      toX = walkMarker(fromX, aabbWidth, columnWidth, maxColumn)
      toY = walkMarker(fromY, aabbHeight, rowHeight, maxRow)
      frameSize = {
        width: Math.max(1, Math.round(extWidth * EMU_PER_PIXEL)),
        height: Math.max(1, Math.round(extHeight * EMU_PER_PIXEL)),
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
    onEdit(visual.id, { anchor: next, ...(frameSize ? { frameSize } : {}) })
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
      data-tip={
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
        <ShapeVisual
          file={file}
          visual={textEditing ? { ...visual, text: '' } : visual}
          frame={frame}
        />
      )}
      {shouldShowVisualDeleteButton({ selected: isSelected, textEditing }) && (
        <VisualDeleteButton
          hostRef={hostNodeRef}
          worksheet={worksheet}
          label={t('appDeleteVisualTitle')}
          onDelete={() => onEdit(visual.id, { remove: true })}
        />
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

/// prstDash → dash/gap units in multiples of the line width (DrawingML's
/// preset patterns); "solid" and unknown values fall through to a solid
/// stroke.
const LINE_DASH_PATTERNS: Record<string, readonly number[]> = {
  dash: [4, 3],
  dashDot: [4, 3, 1, 3],
  dot: [1, 3],
  lgDash: [8, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
  sysDot: [1, 1],
}

/// DrawingML gradient angle (degrees clockwise, 0 = left-to-right) to SVG
/// linearGradient endpoints in objectBoundingBox units. The bounding-box
/// stretch matches lin@scaled gradients closely enough for cell-anchored
/// shapes.
function gradientEndpoints(angle: number): { x1: number; y1: number; x2: number; y2: number } {
  const radians = (angle * Math.PI) / 180
  const dx = Math.cos(radians) / 2
  const dy = Math.sin(radians) / 2
  const round = (value: number): number => Math.round(value * 1000) / 1000
  return { x1: round(0.5 - dx), y1: round(0.5 - dy), x2: round(0.5 + dx), y2: round(0.5 + dy) }
}

function ShapeVisual({
  file,
  visual,
  frame,
}: {
  readonly file?: VisualHost | undefined
  readonly visual: WorkbookVisualObject
  readonly frame?: ShapeFrame | undefined
}): React.JSX.Element {
  // useId's delimiters (":"/"«») are hostile to url(#…) parsing.
  const shapeId = useId().replace(/[^A-Za-z0-9]/g, '')
  const gradientId = `shape-fill-${shapeId}`
  const clipId = `shape-clip-${shapeId}`
  // a:blipFill: the image paints clipped to the geometry, under the stroke.
  const { url: fillImage } = useWorkbookMediaUrl(
    visual.fillMediaPath && file ? file.sessionId : undefined,
    visual.id,
  )
  const type = visual.shapeType ?? ''
  const gradient = visual.fillGradient
  const customPath = visual.customPath
  // "none" is an explicit <a:noFill/> — transparent, not the default tint.
  // Custom geometry never gets the default tint either: a custGeom without
  // an explicit fill is stroke-only artwork, not an inserted preset.
  const fill = gradient
    ? `url(#${gradientId})`
    : visual.fillColor === 'none'
      ? 'transparent'
      : (visual.fillColor ?? (customPath ? 'transparent' : '#DDEBF7'))
  // A rotated shape's anchor stores its rotated bounds (Excel writes the
  // quadrant-swapped snap rect, LibreOffice the AABB) while xfrm ext keeps
  // the true unrotated frame; both center the anchor on the shape center.
  // Restore the ext-sized box around that center, then rotate it. Without
  // ext, fall back to swapping the anchor box inside the swap quadrants.
  const rotation = (((visual.rotation ?? 0) % 360) + 360) % 360
  const trueFrame =
    rotation !== 0 && frame && visual.frameWidth && visual.frameHeight
      ? { width: visual.frameWidth / EMU_PER_PIXEL, height: visual.frameHeight / EMU_PER_PIXEL }
      : undefined
  const swapFallback =
    rotation !== 0 &&
    !trueFrame &&
    ((rotation >= 45 && rotation < 135) || (rotation >= 225 && rotation < 315))
  // Preset geometry is aspect-dependent (snips/rounds scale with the short
  // side); use the real frame proportions when the install-time layout is
  // known instead of a stretched 100×100 square.
  const boxWidth = trueFrame
    ? Math.max(1, trueFrame.width)
    : frame
      ? Math.max(1, swapFallback ? frame.height : frame.width)
      : 100
  const boxHeight = trueFrame
    ? Math.max(1, trueFrame.height)
    : frame
      ? Math.max(1, swapFallback ? frame.width : frame.height)
      : 100
  // An unflipped OOXML line runs top-left → bottom-right; flips mirror the
  // endpoints (the gallery preview draws ascending, which is the flipV form).
  const isStraightLine = type === 'line' || type === 'straightConnector1'
  const lineX0 = visual.flipH ? boxWidth : 0
  const lineY0 = visual.flipV ? boxHeight : 0
  // Same geometry source as the gallery previews (and the other apps'
  // renderers), so every insertable preset draws its real silhouette.
  // custGeom paths keep their own coordinate space and scale into the box.
  const d = customPath
    ? customPath.d
    : isStraightLine
      ? `M ${lineX0} ${lineY0} L ${boxWidth - lineX0} ${boxHeight - lineY0}`
      : shapePreviewPath(type, boxWidth, boxHeight)
  if (!d) {
    // Unsupported geometry: never leak the internal shape name (Excel shows
    // no frame at all) — an empty frame only when the shape carries text.
    if (!visual.text) return <div aria-hidden="true" />
    return (
      <div className="xlsx-shape">
        <span>{visual.text}</span>
      </div>
    )
  }
  const stroke = visual.lineColor === 'none' ? 'transparent' : (visual.lineColor ?? '#00000022')
  // Session text edits overwrite `text` but not `paragraphs`; a mismatch
  // means the flat text is the newer truth.
  const paragraphs =
    visual.paragraphs &&
    visual.paragraphs
      .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
      .join('\n') === (visual.text ?? '')
      ? visual.paragraphs
      : undefined
  const transforms: string[] = []
  if (visual.rotation) transforms.push(`rotate(${visual.rotation}deg)`)
  if (!isStraightLine && (visual.flipH || visual.flipV)) {
    transforms.push(`scale(${visual.flipH ? -1 : 1}, ${visual.flipV ? -1 : 1})`)
  }
  const strokeWidth = visual.lineWidth ?? 1
  let dashPattern = visual.lineDash ? LINE_DASH_PATTERNS[visual.lineDash] : undefined
  // Round-capped dot patterns render as true circles via zero-length dashes:
  // the stretched viewBox distorts dash lengths but not the cap circle.
  if (dashPattern && visual.lineCap === 'rnd' && dashPattern[0] === 1 && dashPattern.length === 2) {
    dashPattern = [0, dashPattern[1]! + 1]
  }
  const content = (
    <>
      <svg viewBox={`0 0 ${boxWidth} ${boxHeight}`} preserveAspectRatio="none" aria-hidden="true">
        {gradient && !isStraightLine && (
          <defs>
            <linearGradient id={gradientId} {...gradientEndpoints(gradient.angle)}>
              {gradient.stops.map((stop, index) => (
                <stop key={index} offset={stop.position} stopColor={stop.color} />
              ))}
            </linearGradient>
          </defs>
        )}
        {customPath?.fillD && (
          // Mixed custGeom: fill only the fillable subpaths; the stroke pass
          // below covers every subpath with a transparent fill.
          <path
            d={customPath.fillD}
            fill={fill}
            stroke="none"
            transform={`scale(${boxWidth / customPath.width}, ${boxHeight / customPath.height})`}
          />
        )}
        {fillImage && !isStraightLine && (
          <>
            <clipPath id={clipId}>
              <path
                d={d}
                {...(customPath
                  ? {
                      transform: `scale(${boxWidth / customPath.width}, ${boxHeight / customPath.height})`,
                    }
                  : {})}
              />
            </clipPath>
            <image
              href={fillImage}
              x={0}
              y={0}
              width={boxWidth}
              height={boxHeight}
              preserveAspectRatio="none"
              clipPath={`url(#${clipId})`}
              {...(visual.opacity !== undefined && visual.opacity < 1
                ? { opacity: visual.opacity }
                : {})}
            />
          </>
        )}
        <path
          d={d}
          fill={
            isStraightLine || customPath?.strokeOnly || customPath?.fillD || fillImage
              ? 'transparent'
              : fill
          }
          {...(customPath
            ? {
                transform: `scale(${boxWidth / customPath.width}, ${boxHeight / customPath.height})`,
              }
            : {})}
          stroke={stroke}
          strokeWidth={strokeWidth}
          {...(dashPattern
            ? { strokeDasharray: dashPattern.map((unit) => unit * strokeWidth).join(' ') }
            : {})}
          {...(visual.lineCap === 'rnd'
            ? { strokeLinecap: 'round' as const }
            : visual.lineCap === 'sq'
              ? { strokeLinecap: 'square' as const }
              : {})}
          vectorEffect="non-scaling-stroke"
        />
      </svg>
      {paragraphs ? (
        <span
          className={`shape-text shape-text-rich shape-anchor-${visual.textAnchor ?? 'ctr'}`}
          style={visual.textColor ? { color: visual.textColor } : undefined}
        >
          {paragraphs.map((paragraph, index) => (
            <span
              key={index}
              className="shape-paragraph"
              style={{
                textAlign:
                  paragraph.align === 'ctr'
                    ? 'center'
                    : paragraph.align === 'r'
                      ? 'right'
                      : paragraph.align === 'just'
                        ? 'justify'
                        : 'left',
              }}
            >
              {paragraph.runs.length === 0
                ? '\u00a0'
                : paragraph.runs.map((run, runIndex) => (
                    <span
                      key={runIndex}
                      style={{
                        ...(run.color ? { color: run.color } : {}),
                        ...(run.bold ? { fontWeight: 700 } : {}),
                        ...(run.italic ? { fontStyle: 'italic' } : {}),
                        ...(run.underline ? { textDecoration: 'underline' } : {}),
                        ...(run.size ? { fontSize: `${run.size}pt` } : {}),
                      }}
                    >
                      {run.text}
                    </span>
                  ))}
            </span>
          ))}
        </span>
      ) : (
        visual.text && (
          <span
            className="shape-text"
            style={visual.textColor ? { color: visual.textColor } : undefined}
          >
            {visual.text}
          </span>
        )
      )}
    </>
  )
  if (trueFrame || swapFallback) {
    // cq units keep the inner box proportional to the float container so it
    // tracks zoom and live resizes.
    const size =
      trueFrame && frame
        ? {
            width: `${(trueFrame.width / frame.width) * 100}cqw`,
            height: `${(trueFrame.height / frame.height) * 100}cqh`,
          }
        : { width: '100cqh', height: '100cqw' }
    return (
      <div className="xlsx-shape-drawn xlsx-shape-rotated">
        <div
          className="xlsx-shape-rotated-inner"
          style={{ ...size, transform: ['translate(-50%, -50%)', ...transforms].join(' ') }}
        >
          {content}
        </div>
      </div>
    )
  }
  return (
    <div
      className="xlsx-shape-drawn"
      style={transforms.length ? { transform: transforms.join(' ') } : undefined}
    >
      {content}
    </div>
  )
}

function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i)
  return bytes
}

/// Lazily fetches a visual's media bytes as a data URL (metafiles rasterize
/// to PNG first). Pass an undefined sessionId to skip fetching.
function useWorkbookMediaUrl(
  sessionId: string | undefined,
  visualId: string,
): { url: string | null; unavailable: boolean } {
  const [url, setUrl] = useState<string | null>(null)
  const [unavailable, setUnavailable] = useState(false)

  useEffect(() => {
    if (sessionId === undefined) return
    let isCurrent = true
    void window.desktopApi
      .readWorkbookMedia({ sessionId, visualId })
      .then(async (media) => {
        const next = isMetafileMime(media.mediaType)
          ? await metafileToDataUrl(base64ToBytes(media.base64), media.mediaType)
          : `data:${media.mediaType};base64,${media.base64}`
        if (!isCurrent) return
        if (next) setUrl(next)
        else setUnavailable(true)
      })
      .catch((reason: unknown) => {
        // Unsupported or unreadable media degrades to an empty frame; the
        // grid must never show the failure text.
        console.warn(`workbook media unavailable (${visualId})`, reason)
        if (isCurrent) setUnavailable(true)
      })
    return () => {
      isCurrent = false
    }
  }, [sessionId, visualId])

  return { url, unavailable }
}

function ImageVisual({
  file,
  visual,
}: {
  readonly file: VisualHost
  readonly visual: WorkbookVisualObject
}): React.JSX.Element {
  // Session-added images carry their bytes inline; nothing to fetch.
  const { url, unavailable } = useWorkbookMediaUrl(
    visual.mediaDataUrl ? undefined : file.sessionId,
    visual.id,
  )
  const source = visual.mediaDataUrl ?? url

  if (unavailable) return <div className="xlsx-visual-unavailable" />
  if (!source) return <div className="xlsx-visual-loading">{t('appImageLoading')}</div>
  const opacity =
    visual.opacity !== undefined && visual.opacity < 1 ? { opacity: visual.opacity } : undefined
  const crop = visual.crop
  // a:srcRect: stretch the source so the cropped window fills the frame
  // (Excel fits the remaining region exactly, no letterboxing).
  const visibleWidth = crop ? 1 - crop.left - crop.right : 1
  const visibleHeight = crop ? 1 - crop.top - crop.bottom : 1
  if (crop && visibleWidth > 0 && visibleHeight > 0) {
    return (
      <span className="xlsx-image-crop">
        <img
          src={source}
          alt={visual.name ?? t('appWorkbookImageAlt')}
          style={{
            width: `${100 / visibleWidth}%`,
            height: `${100 / visibleHeight}%`,
            left: `${(-100 * crop.left) / visibleWidth}%`,
            top: `${(-100 * crop.top) / visibleHeight}%`,
            ...opacity,
          }}
        />
      </span>
    )
  }
  return (
    <img
      className="xlsx-image"
      src={source}
      alt={visual.name ?? t('appWorkbookImageAlt')}
      style={opacity}
    />
  )
}

/// Embedded OLE object: Excel's cached preview picture (EMF/WMF rasterized
/// like any other metafile) at the object's anchor, or — with no preview or
/// an undecodable one — Excel's icon-and-caption box. Always read-only: the
/// object lives in <oleObjects>/xl/embeddings, outside the drawing edit
/// pipeline, so it is never moved, deleted or re-saved as a picture.
function OleVisual({
  file,
  visual,
}: {
  readonly file: VisualHost
  readonly visual: WorkbookVisualObject
}): React.JSX.Element {
  const { url, unavailable } = useWorkbookMediaUrl(
    visual.mediaPath === undefined ? undefined : file.sessionId,
    visual.id,
  )
  const caption = oleCaption(visual.progId)
  // The legacy VML shape's stroke/fill are document colours (Excel's
  // window/windowText by default): the frame Excel draws around every
  // embedded object and the backdrop that hides the grid behind it.
  const frame = oleFrameStyle(visual)
  if (oleRenderKind(visual, unavailable) === 'preview') {
    if (!url) return <div className="xlsx-visual-loading">{t('appImageLoading')}</div>
    return (
      <div className="xlsx-ole-frame" style={frame}>
        <img className="xlsx-image xlsx-ole-preview" src={url} alt={caption} draggable={false} />
      </div>
    )
  }
  return (
    <div
      className="xlsx-ole-placeholder"
      style={frame}
      role="img"
      aria-label={caption}
      title={caption}
    >
      <svg className="xlsx-ole-icon" viewBox="0 0 24 24" aria-hidden="true">
        <path
          d="M6 2h8l5 5v15H6z"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.5"
          strokeLinejoin="round"
        />
        <path d="M14 2v5h5" fill="none" stroke="currentColor" strokeWidth="1.5" />
        <path d="M9 12h7M9 15h7M9 18h5" stroke="currentColor" strokeWidth="1.5" />
      </svg>
      <span className="xlsx-ole-caption">{caption}</span>
    </div>
  )
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
      (series.categories.length === 0 && series.categoriesRef) ||
      series.nameRef,
  )
  const [hydration, setHydration] = useState<ReadonlyMap<
    number,
    {
      values: readonly number[] | null
      categories: readonly string[] | null
      name: string | null
    }
  > | null>(null)
  useEffect(() => {
    if (!needsHydration || !readVector) return
    let cancelled = false
    void (async () => {
      // The lazy sheet may still be indexing right after open; retry briefly.
      for (let attempt = 0; attempt < 8 && !cancelled; attempt += 1) {
        const entries = new Map<
          number,
          {
            values: readonly number[] | null
            categories: readonly string[] | null
            name: string | null
          }
        >()
        let retry = false
        for (const [index, series] of sourceChart.series.entries()) {
          const valuesRef = series.values.length === 0 ? series.valuesRef : undefined
          const categoriesRef = series.categories.length === 0 ? series.categoriesRef : undefined
          const nameRef = series.nameRef
          if (valuesRef === undefined && categoriesRef === undefined && nameRef === undefined) {
            continue
          }
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
            // The name (an uncached c:tx cell reference) throws into the
            // retry loop when it is the sole target — a still-indexing sheet
            // must not freeze the SeriesN placeholder. Alongside other reads
            // it stays best-effort: the placeholder already stands in.
            const name =
              nameRef === undefined
                ? null
                : valuesRef === undefined && categoriesRef === undefined
                  ? await readVector(editKey, nameRef)
                  : await readVector(editKey, nameRef).catch(() => null)
            if (values === null && categories === null && name === null) continue
            const nameText = String(name?.vector[0] ?? '').slice(0, 255)
            entries.set(index, {
              values:
                values?.vector.map((value) => {
                  const numeric =
                    typeof value === 'number' ? value : Number(String(value ?? '').trim())
                  return Number.isFinite(numeric) ? numeric : 0
                }) ?? null,
              categories:
                categories?.vector.map((value) => String(value ?? '').slice(0, 255)) ?? null,
              name: nameText === '' ? null : nameText,
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
            ...(entry.name !== null ? { name: entry.name } : {}),
          }
        }),
      }
    : sourceChart
  const chart = withChartEdit(hydratedSource, pendingEdit)
  const populated = chart.series.filter((series) => series.values.length > 0)
  const primarySeries = populated[0]
  if (!primarySeries) {
    // No cached points and nothing hydrated: Excel shows a silent empty
    // frame, never a diagnostic.
    return <figure className="xlsx-chart" />
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
  // Per-side axes (xAxis/yAxis, paired by c:axPos) win over the legacy
  // element-kind fields — scatter charts have two value axes and the legacy
  // pairing flips their titles. The slots here are SEMANTIC
  // (category/value); on a horizontal bar chart the category axis sits on
  // the left (yAxis position) and the value axis on the bottom, so the
  // positional fields map crosswise — BarChart swaps them back for layout.
  const isHorizontalBar = chart.barDirection === 'bar'
  const effectiveAxisTitles = {
    category: (isHorizontalBar ? chart.yAxis : chart.xAxis)?.title ?? chart.axisTitles?.category,
    value: (isHorizontalBar ? chart.xAxis : chart.yAxis)?.title ?? chart.axisTitles?.value,
  }
  // The value axis sits on the bottom (xAxis slot) for horizontal bars.
  const valueAxisSide = isHorizontalBar ? chart.xAxis : chart.yAxis
  const valueScaleInput = {
    min: valueAxisSide?.min ?? chart.valueAxis?.min,
    max: valueAxisSide?.max ?? chart.valueAxis?.max,
    majorUnit: valueAxisSide?.majorUnit,
    numFmt: valueAxisSide?.numFmt,
  }
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
          data-tip={t('appEditChartTitle')}
          aria-label={t('appEditChartTitle')}
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
      {chart.title !== '' && (
        <figcaption
          className={selectedEl?.kind === 'title' ? 'chart-el-selected' : undefined}
          style={{
            ...(chart.titleStyle?.size ? { fontSize: `${chart.titleStyle.size}pt` } : {}),
            ...(chart.titleStyle?.bold === false ? { fontWeight: 400 } : {}),
            ...(chart.titleStyle?.color ? { color: chart.titleStyle.color } : {}),
          }}
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
      )}
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
            axisTitles={effectiveAxisTitles}
            dataLabels={chart.dataLabels}
            gridlines={chart.gridlines}
            valueAxis={valueScaleInput}
            xAxis={chart.xAxis}
            scatterStyle={chart.scatterStyle}
            categoryFormat={categoryFormat}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        ) : isArea ? (
          <AreaChart
            seriesList={populated}
            axisTitles={effectiveAxisTitles}
            dataLabels={chart.dataLabels}
            grouping={chart.grouping}
            gridlines={chart.gridlines}
            valueAxis={valueScaleInput}
            categoryFormat={categoryFormat}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        ) : isLine ? (
          <LineChart
            seriesList={populated}
            axisTitles={effectiveAxisTitles}
            dataLabels={chart.dataLabels}
            grouping={chart.grouping}
            gridlines={chart.gridlines}
            valueAxis={valueScaleInput}
            categoryFormat={categoryFormat}
            lineMarkers={chart.lineMarkers}
            dispBlanksAs={chart.dispBlanksAs}
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
            axisTitles={effectiveAxisTitles}
            dataLabels={chart.dataLabels}
            dataLabelPosition={chart.dataLabelPosition}
            dataLabelFormat={chart.dataLabelFormat}
            grouping={chart.grouping}
            gridlines={chart.gridlines}
            valueAxis={valueScaleInput}
            secondaryAxis={chart.secondaryYAxis}
            gapWidthPct={chart.gapWidthPct}
            categoryFormat={categoryFormat}
            categoryReversed={(isHorizontalBar ? chart.yAxis : chart.xAxis)?.reversed === true}
            lineMarkers={chart.lineMarkers}
            dispBlanksAs={chart.dispBlanksAs}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        )}
        {!isPie &&
          chart.legend !== 'none' &&
          (chart.legend !== undefined || populated.length > 1) && (
            <SeriesLegend
              seriesList={populated}
              lineSwatches={legendUsesLineSwatches(types)}
              lineSwatchFrom={isCombo && populated.length > 1 ? populated.length - 1 : undefined}
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
  const menuRef = useRef<HTMLDivElement>(null)
  // Outside press / window blur / shell chrome press — the shared dismissal
  // (the click-through backdrop below stays as belt and braces). The menu is
  // portaled to the body, so the inside guard must be its own element.
  useDismissablePopover(true, onClose, { inside: () => [menuRef.current] })
  const closeRef = useRef(onClose)
  closeRef.current = onClose
  useEffect(() => {
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') closeRef.current()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])
  // Portal to the body: the float-DOM container clips overflowing children.
  return createPortal(
    <div
      className="chart-menu-backdrop"
      onClick={onClose}
      onContextMenu={(event) => event.preventDefault()}
    >
      <div
        ref={menuRef}
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
  bottomY = 317,
}: {
  readonly bottom?: string | null | undefined
  readonly left?: string | null | undefined
  /// Charts with an outer category tier extend the viewBox and push the
  /// bottom title below the group-label band.
  readonly bottomY?: number
}): React.JSX.Element {
  return (
    <g>
      {bottom ? (
        <text x="320" y={bottomY} textAnchor="middle" className="axis-title">
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

/// Excel's automatic slice cycle past the six accents: the same accents
/// lightened (HSL L' = 0.6·L + 0.4), verified against Excel-rendered pies.
const pieSliceColorsLight = ['#8faadc', '#f4b183', '#c9c9c9', '#ffd966', '#9dc3e6', '#a9d18e']

/// Pie/doughnut slices color per point: explicit `c:dPt` fills win, the
/// Office palette cycles underneath (Excel's varyColors default) — six
/// accents, then their lighter variants, then around again.
function pieSliceColor(series: SeriesLike, index: number): string {
  const explicit = series.pointColors?.find((point) => point.index === index)?.color
  if (explicit) return explicit
  const cycle = index % 12
  return (cycle < 6 ? chartColors[cycle] : pieSliceColorsLight[cycle - 6]) ?? '#4472c4'
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
          <Dropdown
            className="chart-editor-type"
            ariaLabel={t('appTypeLabel')}
            value={chartType}
            options={CHART_TYPE_OPTIONS.map((option) => ({
              value: option.value,
              label: t(option.labelKey),
            }))}
            onPick={setChartType}
          />
        </label>
      )}
      {!isPie &&
        chart.series.slice(0, 6).map((series, index) => (
          <label key={index}>
            {truncateLabel(series.name || t('appSeriesN', { n: index + 1 }), 14)}
            <ColorDropdown
              label={series.name || t('appSeriesN', { n: index + 1 })}
              value={colors[String(index)] ?? seriesColor(series, index)}
              portal
              onPick={(hex) => {
                if (hex) setColors((previous) => ({ ...previous, [String(index)]: hex }))
              }}
            />
          </label>
        ))}
      {isPie &&
        chart.series[0]?.categories.slice(0, 12).map((category, index) => (
          <label key={`slice-${index}`}>
            {truncateLabel(category || t('appSliceN', { n: index + 1 }), 14)}
            <ColorDropdown
              label={category || t('appSliceN', { n: index + 1 })}
              value={sliceColors[String(index)] ?? pieSliceColor(chart.series[0] ?? {}, index)}
              portal
              onPick={(hex) => {
                if (hex) setSliceColors((previous) => ({ ...previous, [String(index)]: hex }))
              }}
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
              data-tip={t('appValuesRangeTitle', {
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
              data-tip={t('appCategoriesRangeTitle', {
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
  lineSwatches = false,
  lineSwatchFrom,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly selected?: boolean | undefined
  readonly onSelect?: ((event: React.MouseEvent) => void) | undefined
  /// Line-family charts: swatches mirror the drawn stroke color.
  readonly lineSwatches?: boolean | undefined
  /// Combo charts: series from this index on render as lines.
  readonly lineSwatchFrom?: number | undefined
}): React.JSX.Element {
  return (
    <div className={`chart-legend${selected ? ' chart-el-selected' : ''}`} onClick={onSelect}>
      {seriesList.slice(0, 8).map((series, index) => (
        <span key={`${series.name}-${index}`}>
          <i
            style={{
              background: legendSwatchColor(
                series,
                index,
                lineSwatches || (lineSwatchFrom !== undefined && index >= lineSwatchFrom),
              ),
            }}
          />
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
  secondaryAxis,
  gapWidthPct,
  categoryFormat,
  categoryReversed,
  lineMarkers,
  dispBlanksAs,
  onElement,
  selectedEl,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly lineSeries?: ChartSeries | undefined
  readonly isHorizontal: boolean
  /// Plot-level c:lineChart/c:marker flag for the combo line.
  readonly lineMarkers?: boolean | undefined
  readonly dispBlanksAs?: ChartMetadata['dispBlanksAs']
  /// c:catAx orientation maxMin — categories read top-down (bar) /
  /// right-to-left (column) instead of Excel's minMax default.
  readonly categoryReversed?: boolean | undefined
  readonly axisTitles?: ChartAxisTitles
  readonly dataLabels?: ChartDataLabels
  readonly dataLabelPosition?: ChartLabelPosition
  readonly dataLabelFormat?: string | undefined
  readonly grouping?: ChartGrouping
  readonly gridlines?: boolean | undefined
  readonly valueAxis?: ChartValueAxis
  readonly secondaryAxis?:
    | {
        min?: number | undefined
        max?: number | undefined
        majorUnit?: number | undefined
        numFmt?: string | undefined
        hidden?: boolean | undefined
      }
    | undefined
  readonly gapWidthPct?: number | undefined
  readonly categoryFormat?: string | undefined
} & ChartElementProps): React.JSX.Element {
  const primary = seriesList[0]
  if (!primary) return <></>
  const categories = primary.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  // 20 was too tight for real corpora (38-county bar charts); 48 keeps
  // bars ≥ ~5px in the 600px plot while very wide data still truncates.
  const visibleCount = Math.min(primary.values.length, 48)
  const isStacked =
    (grouping === 'stacked' || grouping === 'percentStacked') && seriesList.length > 1
  const isPercent = grouping === 'percentStacked' && seriesList.length > 1
  const categoryTotal = (index: number): number =>
    seriesList.reduce((sum, series) => sum + Math.max(0, series.values[index] ?? 0), 0)
  const barMax = isStacked
    ? Math.max(...Array.from({ length: visibleCount }, (_, index) => categoryTotal(index)), 0)
    : Math.max(...seriesList.flatMap((series) => [...series.values]), 0)
  // A combo line without its own value axis shares the primary scale.
  const lineOnPrimary = lineSeries !== undefined && secondaryAxis === undefined
  const bounds = isPercent
    ? { min: 0, max: 1, ticks: [0, 0.25, 0.5, 0.75, 1] }
    : axisBounds(
        lineOnPrimary ? Math.max(barMax, ...(lineSeries?.values ?? [])) : barMax,
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
  const axisNumberFormat = isPercent ? '0%' : (valueAxis?.numFmt ?? primary.numberFormat)
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
    const rowHeight = 270 / visibleCount
    const rowGroups = visibleCategoryGroups(primary.categoryGroups, visibleCount)
    // The outer tier owns the x ≤ 99 column; tick labels (end-anchored at
    // 148) shrink so the two never share pixels.
    const categoryBudget = rowGroups.length > 0 ? 8 : 14
    const barHeight = Math.max(
      3,
      isStacked ? rowHeight / (1 + gap) : rowHeight / (seriesList.length + gap),
    )
    // Excel draws bar-chart categories bottom-up and series 0 nearest the
    // category axis under the default minMax orientation; maxMin flips both
    // back to top-down reading order.
    const rowSlot = (index: number): number => (categoryReversed ? index : visibleCount - 1 - index)
    const seriesSlot = (seriesIndex: number): number =>
      categoryReversed ? seriesIndex : seriesList.length - 1 - seriesIndex
    const groupTop = (index: number): number =>
      14 +
      rowHeight * rowSlot(index) +
      (rowHeight - barHeight * (isStacked ? 1 : seriesList.length)) / 2
    const plotX = (value: number): number =>
      158 + Math.max(0, Math.min(1, (value - bounds.min) / (bounds.max - bounds.min || 1))) * 390
    const labelFor = (
      value: number,
      y: number,
      numberFormat?: string,
    ): React.JSX.Element | null => {
      const width = norm(value) * 390
      const inside = dataLabelPosition === 'center' || dataLabelPosition === 'inside-end'
      const x =
        dataLabelPosition === 'center'
          ? 158 + width / 2
          : dataLabelPosition === 'inside-end'
            ? 154 + width
            : 164 + width
      return (
        <text
          x={x}
          y={y}
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
          {formatLabelValue(value, dataLabelFormat, numberFormat)}
        </text>
      )
    }
    return (
      <svg className="chart-svg" viewBox="0 0 600 320" role="img">
        {bounds.ticks.map((tick, index) => (
          <g key={`vt-${index}`} onClick={selectValueAxis}>
            {gridlines !== false && (
              <line
                x1={plotX(tick)}
                y1="12"
                x2={plotX(tick)}
                y2="286"
                stroke="#e3e3e3"
                strokeWidth="1"
              />
            )}
            {/* Excel draws the bar-chart value axis along the bottom. */}
            <text x={plotX(tick)} y="298" textAnchor="middle" className="axis-label">
              {formatAxisValue(tick, axisNumberFormat)}
            </text>
          </g>
        ))}
        {Array.from({ length: visibleCount }, (_, index) => {
          let cursor = 0
          return (
            <g key={`${categories[index] ?? index}-${index}`}>
              <text
                x="148"
                y={26 + rowHeight * rowSlot(index)}
                textAnchor="end"
                onClick={selectCategoryAxis}
              >
                {truncateLabel(categories[index] ?? String(index + 1), categoryBudget)}
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
                    y={groupTop(index) + (isStacked ? 0 : barHeight * seriesSlot(seriesIndex))}
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
                seriesList.map((series, seriesIndex) => (
                  <Fragment key={`lbl-${seriesIndex}`}>
                    {labelFor(
                      series.values[index] ?? 0,
                      groupTop(index) + barHeight * seriesSlot(seriesIndex) + barHeight / 2 + 3,
                      series.numberFormat,
                    )}
                  </Fragment>
                ))}
            </g>
          )
        })}
        {/* Outer multiLvlStrCache level: a rotated label column left of the
            category labels with separator lines at the group edges. */}
        {rowGroups.map((group, groupIndex) => {
          const first = rowSlot(group.start)
          const last = rowSlot(group.end - 1)
          const yTop = 14 + rowHeight * Math.min(first, last)
          const yBottom = 14 + rowHeight * (Math.max(first, last) + 1)
          const cy = (yTop + yBottom) / 2
          return (
            <g key={`grp-${groupIndex}`} onClick={selectCategoryAxis}>
              <line x1="86" y1={yTop} x2="158" y2={yTop} stroke="#d9d9d9" strokeWidth="1" />
              <line x1="86" y1={yBottom} x2="158" y2={yBottom} stroke="#d9d9d9" strokeWidth="1" />
              <text x="94" y={cy} transform={`rotate(-90 94 ${cy})`} textAnchor="middle">
                {truncateLabel(
                  group.label,
                  Math.max(3, Math.floor((yBottom - yTop) / AXIS_LABEL_CHAR_UNITS)),
                )}
              </text>
            </g>
          )
        })}
        <TruncationNote shown={visibleCount} total={primary.values.length} />
        <AxisTitleTexts bottom={axisTitles?.value} left={axisTitles?.category} />
      </svg>
    )
  }

  const columnWidth = 480 / visibleCount
  const tickStride = categoryTickStride(categories, visibleCount, columnWidth)
  const barWidth = Math.max(
    2,
    isStacked ? columnWidth / (1 + gap) : columnWidth / (seriesList.length + gap),
  )
  const groupWidth = barWidth * (isStacked ? 1 : seriesList.length)
  // maxMin mirrors the column order (and the series order inside a group).
  const catSlot = (index: number): number => (categoryReversed ? visibleCount - 1 - index : index)
  const seriesSlot = (seriesIndex: number): number =>
    categoryReversed ? seriesList.length - 1 - seriesIndex : seriesIndex
  const groupLeft = (index: number): number =>
    62 + columnWidth * catSlot(index) + (columnWidth - groupWidth) / 2
  // The combo line rides the secondary value axis when the file has one;
  // a single axis pair means it shares the primary scale (Excel never
  // invents a right-hand axis).
  const lineScale = lineSeries
    ? secondaryAxis
      ? valueAxisScale(Math.max(...lineSeries.values, 0), secondaryAxis)
      : bounds
    : undefined
  const comboStroke = lineSeries ? lineStroke(lineSeries, seriesList.length) : null
  const comboSegments =
    lineSeries && lineScale
      ? lineSegments(
          Math.min(lineSeries.values.length, visibleCount),
          lineSeries.blanks,
          dispBlanksAs,
        )
      : []
  const comboY = (index: number): number =>
    280 -
    (lineScale
      ? Math.max(
          0,
          Math.min(
            1,
            ((lineSeries?.values[index] ?? 0) - lineScale.min) /
              (lineScale.max - lineScale.min || 1),
          ),
        )
      : 0) *
      240
  const comboPoint = (index: number): string =>
    `${groupLeft(index) + groupWidth / 2},${comboY(index)}`
  const comboSymbol =
    lineSeries === undefined || lineSeries.marker === 'none'
      ? null
      : lineMarkers !== true && lineSeries.marker === undefined
        ? null
        : lineSeries.marker !== undefined && lineSeries.marker !== 'auto'
          ? lineSeries.marker
          : (AUTO_MARKER_SYMBOLS[seriesList.length % AUTO_MARKER_SYMBOLS.length] ?? 'circle')
  const showValueLabels =
    !isStacked &&
    (dataLabels === 'value' ||
      (dataLabels === undefined && visibleCount <= 12 && seriesList.length === 1))
  const columnGroups = visibleCategoryGroups(primary.categoryGroups, visibleCount)
  // The group band occupies y 284-314; a bottom axis title moves below it
  // on an extended canvas instead of overprinting.
  const shiftBottomTitle = columnGroups.length > 0 && Boolean(axisTitles?.category)
  return (
    <svg className="chart-svg" viewBox={`0 0 600 ${shiftBottomTitle ? 336 : 320}`} role="img">
      <VerticalAxis
        minimum={bounds.min}
        maximum={bounds.max}
        ticks={isPercent ? undefined : bounds.ticks}
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
                  x={groupLeft(index) + (isStacked ? 0 : barWidth * seriesSlot(seriesIndex))}
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
              seriesList.map((series, seriesIndex) => {
                const share = norm(series.values[index] ?? bounds.min)
                const inside = dataLabelPosition === 'center' || dataLabelPosition === 'inside-end'
                const y =
                  dataLabelPosition === 'center'
                    ? 283 - share * 120
                    : dataLabelPosition === 'inside-end'
                      ? 292 - share * 240
                      : 272 - share * 240
                const x = groupLeft(index) + barWidth * seriesSlot(seriesIndex) + barWidth / 2
                return (
                  <text
                    key={`lbl-${seriesIndex}`}
                    x={x}
                    y={y}
                    textAnchor="middle"
                    className="axis-label"
                    {...(inside ? { fill: '#fff' } : {})}
                  >
                    {formatLabelValue(
                      series.values[index] ?? 0,
                      dataLabelFormat,
                      series.numberFormat,
                    )}
                  </text>
                )
              })}
            <CategoryTick
              x={62 + columnWidth * catSlot(index) + columnWidth / 2}
              label={categories[index] ?? String(index + 1)}
              slotWidth={columnWidth}
              stride={tickStride}
              index={index}
              onClick={selectCategoryAxis}
            />
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
      {comboStroke !== null &&
        comboSegments.map((segment, segmentIndex) => (
          <polyline
            key={`combo-${segmentIndex}`}
            points={segment.map(comboPoint).join(' ')}
            fill="none"
            stroke={comboStroke}
            strokeWidth={lineSeries?.lineWidth ?? 3}
          />
        ))}
      {lineSeries && comboSymbol !== null && (
        <g>
          {comboSegments.flat().map((index) => (
            <MarkerGlyph
              key={`combo-marker-${index}`}
              x={groupLeft(index) + groupWidth / 2}
              y={comboY(index)}
              symbol={comboSymbol}
              color={comboStroke ?? seriesColor(lineSeries, seriesList.length)}
            />
          ))}
        </g>
      )}
      <CategoryGroupBand
        spans={columnGroups.map((group) => {
          const first = catSlot(group.start)
          const last = catSlot(group.end - 1)
          return {
            label: group.label,
            xStart: 62 + columnWidth * Math.min(first, last),
            xEnd: 62 + columnWidth * (Math.max(first, last) + 1),
          }
        })}
        onClick={selectCategoryAxis}
      />
      {lineSeries &&
        lineScale &&
        secondaryAxis !== undefined &&
        secondaryAxis.hidden !== true &&
        lineScale.ticks.map((tick, index) => (
          <text
            key={index}
            x="596"
            y={284 - ((tick - lineScale.min) / (lineScale.max - lineScale.min || 1)) * 240}
            textAnchor="end"
            className="axis-label"
          >
            {formatAxisValue(tick, secondaryAxis?.numFmt ?? lineSeries.numberFormat)}
          </text>
        ))}
      <TruncationNote shown={visibleCount} total={primary.values.length} />
      <AxisTitleTexts
        bottom={axisTitles?.category}
        left={axisTitles?.value}
        bottomY={shiftBottomTitle ? 333 : 317}
      />
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

/// Excel's automatic marker cycle when the plot flag is on but the series
/// sets no symbol (diamond, square, triangle, then repeat).
const AUTO_MARKER_SYMBOLS = ['diamond', 'square', 'triangle'] as const

function MarkerGlyph({
  x,
  y,
  symbol,
  color,
}: {
  readonly x: number
  readonly y: number
  readonly symbol: string
  readonly color: string
}): React.JSX.Element {
  const r = 4
  if (symbol === 'diamond') {
    return (
      <polygon points={`${x},${y - r} ${x + r},${y} ${x},${y + r} ${x - r},${y}`} fill={color} />
    )
  }
  if (symbol === 'square') {
    return <rect x={x - r} y={y - r} width={2 * r} height={2 * r} fill={color} />
  }
  if (symbol === 'triangle') {
    return <polygon points={`${x},${y - r} ${x + r},${y + r} ${x - r},${y + r}`} fill={color} />
  }
  return <circle cx={x} cy={y} r={r} fill={color} />
}

/// Series line stroke: an explicit spPr/a:ln color wins over the fill/accent
/// default; null for an explicit noFill line.
export function lineStroke(
  series: SeriesLike & { readonly lineColor?: string | undefined },
  index: number,
): string | null {
  if (series.lineColor === 'none') return null
  return series.lineColor ?? seriesColor(series, index)
}

/// Index runs a line plots as one stroke. Blank cells (c:dispBlanksAs)
/// break the line on 'gap', are bridged on 'span', and plot their 0 filler
/// on 'zero' — the OOXML default when the element is absent.
export function lineSegments(
  count: number,
  blanks: readonly number[] | undefined,
  mode: string | undefined,
): number[][] {
  const indices = Array.from({ length: count }, (_, index) => index)
  if ((mode !== 'gap' && mode !== 'span') || blanks === undefined || blanks.length === 0) {
    return count > 0 ? [indices] : []
  }
  const blankSet = new Set(blanks)
  if (mode === 'span') {
    const kept = indices.filter((index) => !blankSet.has(index))
    return kept.length > 0 ? [kept] : []
  }
  const segments: number[][] = []
  let current: number[] = []
  for (const index of indices) {
    if (blankSet.has(index)) {
      if (current.length > 0) segments.push(current)
      current = []
    } else {
      current.push(index)
    }
  }
  if (current.length > 0) segments.push(current)
  return segments
}

/// Whether the legend mirrors stroke colors: only when a line-drawing
/// component wins the render cascade (pie → radar → scatter → area → line →
/// bar) — a chart that also carries an areaChart/scatterChart paints fills.
export function legendUsesLineSwatches(types: readonly string[]): boolean {
  if (types.includes('pieChart') || types.includes('doughnutChart')) return false
  const noBar = !types.includes('barChart')
  if (types.includes('radarChart') && noBar) return true
  if (types.includes('scatterChart')) return false
  if (types.includes('areaChart') && noBar) return false
  return types.includes('lineChart') && noBar
}

/// Legend swatch color: line-family charts mirror the drawn stroke (an
/// explicit noFill line keeps the series color); everything else keeps the
/// fill resolution.
export function legendSwatchColor(
  series: SeriesLike & { readonly lineColor?: string | undefined },
  index: number,
  lineSwatches: boolean,
): string {
  if (!lineSwatches) return seriesColor(series, index)
  return lineStroke(series, index) ?? seriesColor(series, index)
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
/// Pie value labels honor the full source number format (currency symbols,
/// accounting padding) the way Excel's sourceLinked labels do.
function formatPieValue(value: number, formatCode: string | undefined): string {
  if (formatCode && formatCode !== 'General') {
    try {
      return numfmt.format(formatCode, value, { throws: false }).trim()
    } catch {
      // Unparseable format: fall through to the plain rendering.
    }
  }
  return formatLabelValue(value, formatCode, undefined)
}

/// Excel's showPercent default format is "0%" — integer rounding — unless
/// the dLbls carry their own percent numFmt.
export function formatPiePercent(share: number, formatCode: string | undefined): string {
  if (formatCode !== undefined && formatCode.includes('%')) {
    return formatLabelValue(share, formatCode, undefined)
  }
  return `${Math.round(share * 100)}%`
}

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

/// Approximate glyph width of the small axis font in the 640-unit viewBox.
const AXIS_LABEL_CHAR_UNITS = 5.4

/// Excel wraps a category label at spaces when its slot is narrow ("KW 01"
/// stacks as KW / 01); greedy two-line split, single line when it fits or
/// has no break point.
export function categoryTickLines(label: string, slotWidth: number): readonly string[] {
  const fits = (text: string): boolean => text.length * AXIS_LABEL_CHAR_UNITS <= slotWidth
  if (fits(label) || !label.includes(' ')) return [label]
  const words = label.split(/\s+/)
  let first = words[0] ?? ''
  let index = 1
  while (index < words.length && fits(`${first} ${words[index]}`)) {
    first = `${first} ${words[index]}`
    index += 1
  }
  const rest = words.slice(index).join(' ')
  return rest ? [first, rest] : [first]
}

/// Excel thins overlapping category labels to every n-th instead of
/// overprinting: when even the wrapped label is wider than its slot, only
/// every ceil(width/slot)-th category keeps its label.
export function categoryTickStride(
  labels: readonly string[],
  count: number,
  slotWidth: number,
): number {
  if (!(slotWidth > 0)) return 1
  let widest = 0
  for (let index = 0; index < count; index += 1) {
    for (const line of categoryTickLines(labels[index] ?? String(index + 1), slotWidth)) {
      widest = Math.max(widest, Math.min(line.length, 16) * AXIS_LABEL_CHAR_UNITS)
    }
  }
  return Math.max(1, Math.ceil((widest + 2) / slotWidth))
}

function CategoryTick({
  x,
  label,
  slotWidth,
  stride,
  index,
  onClick,
}: {
  readonly x: number
  readonly label: string
  readonly slotWidth: number
  readonly stride: number
  readonly index: number
  readonly onClick?: ((event: React.MouseEvent) => void) | undefined
}): React.JSX.Element | null {
  if (index % stride !== 0) return null
  const budget = Math.max(5, Math.min(16, Math.floor((slotWidth * stride) / AXIS_LABEL_CHAR_UNITS)))
  const lines = categoryTickLines(label, slotWidth * stride)
  // Two-line ticks start higher so the second line (baseline 303) clears
  // the bottom axis title at y=317; the plot floor is 280, so line one at
  // 292 stays below the bars either way.
  return (
    <text x={x} y={lines.length === 1 ? 298 : 292} textAnchor="middle" onClick={onClick}>
      {lines.length === 1 ? (
        truncateLabel(lines[0] ?? '', budget)
      ) : (
        <>
          <tspan x={x} dy="0">
            {truncateLabel(lines[0] ?? '', budget)}
          </tspan>
          <tspan x={x} dy="11">
            {truncateLabel(lines[1] ?? '', budget)}
          </tspan>
        </>
      )}
    </text>
  )
}

/// Clamp parsed outer-level category spans to the drawn category count.
function visibleCategoryGroups(
  groups: ChartSeries['categoryGroups'],
  visibleCount: number,
): { label: string; start: number; end: number }[] {
  return (groups ?? [])
    .map((group) => ({ ...group, end: Math.min(group.end, visibleCount) }))
    .filter((group) => group.start < group.end && group.start < visibleCount)
}

/// Outer multiLvlStrCache level under the tick labels: each group label
/// centered under its span with separator ticks at the span edges (Excel's
/// grouped category axis).
function CategoryGroupBand({
  spans,
  onClick,
}: {
  readonly spans: readonly { label: string; xStart: number; xEnd: number }[]
  readonly onClick?: ((event: React.MouseEvent) => void) | undefined
}): React.JSX.Element {
  return (
    <g onClick={onClick}>
      {spans.map((span, index) => (
        <Fragment key={index}>
          <line
            x1={span.xStart}
            y1="284"
            x2={span.xStart}
            y2="314"
            stroke="#d9d9d9"
            strokeWidth="1"
          />
          <line x1={span.xEnd} y1="284" x2={span.xEnd} y2="314" stroke="#d9d9d9" strokeWidth="1" />
          <text x={(span.xStart + span.xEnd) / 2} y="312" textAnchor="middle">
            {truncateLabel(
              span.label,
              Math.max(3, Math.floor((span.xEnd - span.xStart) / AXIS_LABEL_CHAR_UNITS)),
            )}
          </text>
        </Fragment>
      ))}
    </g>
  )
}

function formatAxisValue(value: number, numberFormat: string | undefined): string {
  if (numberFormat && numberFormat !== 'General' && !numberFormat.includes('%')) {
    try {
      const text = numfmt.format(numberFormat, value, { throws: false })
      if (typeof text === 'string' && text !== '') return text
    } catch {
      // fall through to the plain rendering
    }
  }
  if (numberFormat?.includes('%')) return `${Math.round(value * 100)}%`
  const magnitude = Math.abs(value)
  // Excel prints full numbers on value axes (no K abbreviation); only guard
  // the layout against extreme magnitudes.
  if (magnitude >= 1e9) return `${(value / 1e9).toFixed(1)}B`
  if (magnitude >= 1e6) return `${(value / 1e6).toFixed(1)}M`
  const clean = Number(value.toPrecision(12))
  return clean.toLocaleString('en-US', { maximumFractionDigits: 4 })
}

function VerticalAxis({
  minimum = 0,
  maximum,
  ticks,
  numberFormat,
  showGridlines = true,
  onSelect,
}: {
  readonly minimum?: number
  readonly maximum: number
  readonly ticks?: readonly number[] | undefined
  readonly numberFormat: string | undefined
  readonly showGridlines?: boolean
  readonly onSelect?: ((event: React.MouseEvent) => void) | undefined
}): React.JSX.Element {
  const span = maximum - minimum || 1
  const list =
    ticks && ticks.length >= 2
      ? ticks
      : [0, 0.25, 0.5, 0.75, 1].map((fraction) => minimum + fraction * span)
  return (
    <g onClick={onSelect}>
      {list.map((tick, index) => {
        const y = 280 - ((tick - minimum) / span) * 240
        return (
          <g key={index}>
            {(showGridlines || index === 0) && (
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

type ChartValueAxis =
  | {
      min?: number | undefined
      max?: number | undefined
      majorUnit?: number | undefined
      numFmt?: string | undefined
    }
  | undefined

/// Explicit axis bounds/unit win; otherwise an Excel-like auto scale.
function axisBounds(
  dataMax: number,
  valueAxis: ChartValueAxis,
): { min: number; max: number; ticks: number[] } {
  return valueAxisScale(dataMax, valueAxis)
}

function LineChart({
  seriesList,
  axisTitles,
  dataLabels,
  grouping,
  gridlines,
  valueAxis,
  categoryFormat,
  lineMarkers,
  dispBlanksAs,
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
  readonly lineMarkers?: boolean | undefined
  readonly dispBlanksAs?: ChartMetadata['dispBlanksAs']
} & ChartElementProps): React.JSX.Element {
  const primary = seriesList[0]
  if (!primary) return <></>
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
    ? { min: 0, max: 1, ticks: [0, 0.25, 0.5, 0.75, 1] }
    : axisBounds(
        isStacked
          ? Math.max(...(stackTotals[stackTotals.length - 1] ?? [0]), 0)
          : Math.max(...seriesList.flatMap((series) => [...series.values]), 0),
        isStacked ? undefined : valueAxis,
      )
  const span = bounds.max - bounds.min
  const categories = primary.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const count = Math.max(1, primary.values.length - 1)
  // Stacked lines keep the 0 fillers (a blank contributes nothing to the
  // stack); only flat lines honor gap/span.
  const seriesSegments = (series: ChartSeries): number[][] =>
    lineSegments(series.values.length, isStacked ? undefined : series.blanks, dispBlanksAs)
  const isSkippedBlank = (series: ChartSeries, index: number): boolean =>
    !isStacked &&
    (dispBlanksAs === 'gap' || dispBlanksAs === 'span') &&
    (series.blanks?.includes(index) ?? false)
  const categoryGroups = visibleCategoryGroups(primary.categoryGroups, primary.values.length)
  // The group band occupies y 284-314; a bottom axis title moves below it
  // on an extended canvas instead of overprinting.
  const shiftBottomTitle = categoryGroups.length > 0 && Boolean(axisTitles?.category)
  return (
    <svg className="chart-svg" viewBox={`0 0 600 ${shiftBottomTitle ? 336 : 320}`} role="img">
      <VerticalAxis
        minimum={bounds.min}
        maximum={bounds.max}
        ticks={isPercent ? undefined : bounds.ticks}
        numberFormat={isPercent ? '0%' : (valueAxis?.numFmt ?? primary.numberFormat)}
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
        const stroke = lineStroke(series, seriesIndex)
        if (stroke === null) return null
        const width = series.lineWidth ?? 3
        const values = displayValues(seriesIndex)
        const denominator = Math.max(1, values.length - 1)
        const pointAt = (index: number): string =>
          `${60 + (index / denominator) * 500},${
            280 - Math.max(0, Math.min(1, ((values[index] ?? 0) - bounds.min) / (span || 1))) * 240
          }`
        return seriesSegments(series).map((segment, segmentIndex) => (
          <polyline
            key={`${seriesIndex}-${segmentIndex}`}
            points={segment.map(pointAt).join(' ')}
            fill="none"
            stroke={stroke}
            strokeWidth={
              selectedEl?.kind === 'series' && selectedEl.seriesIndex === seriesIndex
                ? Math.max(5, width + 2)
                : width
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
        ))
      })}
      {seriesList.map((series, seriesIndex) => {
        // Plot flag on or explicit symbol; a series-level "none" always wins.
        if (series.marker === 'none') return null
        if (lineMarkers !== true && series.marker === undefined) return null
        const symbol =
          series.marker !== undefined && series.marker !== 'auto'
            ? series.marker
            : (AUTO_MARKER_SYMBOLS[seriesIndex % AUTO_MARKER_SYMBOLS.length] ?? 'circle')
        const values = displayValues(seriesIndex)
        const pointCount = Math.max(1, values.length - 1)
        // Markers match the polyline stroke; an explicit noFill line keeps
        // its markers in the series color.
        const color = lineStroke(series, seriesIndex) ?? seriesColor(series, seriesIndex)
        return (
          <g key={`markers-${seriesIndex}`}>
            {values.map((value, index) =>
              isSkippedBlank(series, index) ? null : (
                <MarkerGlyph
                  key={index}
                  x={60 + (index / pointCount) * 500}
                  y={280 - Math.max(0, Math.min(1, (value - bounds.min) / (span || 1))) * 240}
                  symbol={symbol}
                  color={color}
                />
              ),
            )}
          </g>
        )
      })}
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
        <CategoryTick
          key={index}
          x={60 + (index / count) * 500}
          label={categories[index] ?? String(index + 1)}
          slotWidth={500 / Math.max(1, count)}
          stride={categoryTickStride(categories, primary.values.length, 500 / Math.max(1, count))}
          index={index}
          onClick={
            onElement
              ? (event) => {
                  event.stopPropagation()
                  onElement({ kind: 'category-axis' })
                }
              : undefined
          }
        />
      ))}
      {dataLabels === 'value' &&
        displayValues(0).map((displayed, index) =>
          isSkippedBlank(primary, index) ? null : (
            <text
              key={`label-${index}`}
              x={60 + (index / count) * 500}
              y={272 - Math.max(0, Math.min(1, (displayed - bounds.min) / span)) * 240}
              textAnchor="middle"
              className="axis-label"
            >
              {formatAxisValue(primary.values[index] ?? 0, primary.numberFormat)}
            </text>
          ),
        )}
      <CategoryGroupBand
        spans={categoryGroups.map((group) => {
          const xAt = (index: number): number => 60 + (index / count) * 500
          // Group edges fall midway between the last point of one group
          // and the first point of the next (plot edges at the ends).
          return {
            label: group.label,
            xStart: group.start === 0 ? 60 : (xAt(group.start - 1) + xAt(group.start)) / 2,
            xEnd:
              group.end >= primary.values.length ? 560 : (xAt(group.end - 1) + xAt(group.end)) / 2,
          }
        })}
        onClick={
          onElement
            ? (event) => {
                event.stopPropagation()
                onElement({ kind: 'category-axis' })
              }
            : undefined
        }
      />
      <AxisTitleTexts
        bottom={axisTitles?.category}
        left={axisTitles?.value}
        bottomY={shiftBottomTitle ? 333 : 317}
      />
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
  if (!primary) return <></>
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
    ? { min: 0, max: 1, ticks: [0, 0.25, 0.5, 0.75, 1] }
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
        ticks={isPercent ? undefined : bounds.ticks}
        numberFormat={isPercent ? '0%' : (valueAxis?.numFmt ?? primary.numberFormat)}
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
        <CategoryTick
          key={index}
          x={60 + (index / count) * 500}
          label={categories[index] ?? String(index + 1)}
          slotWidth={500 / Math.max(1, count)}
          stride={categoryTickStride(categories, primary.values.length, 500 / Math.max(1, count))}
          index={index}
        />
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
  if (!primary) return <></>
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
            stroke={isSel ? '#107C41' : (lineStroke(series, seriesIndex) ?? 'none')}
            strokeWidth={
              isSel ? Math.max(4, (series.lineWidth ?? 2.5) + 1.5) : (series.lineWidth ?? 2.5)
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
  xAxis,
  scatterStyle,
  categoryFormat,
  onElement,
  selectedEl,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly axisTitles?: ChartAxisTitles
  readonly dataLabels?: ChartDataLabels
  readonly gridlines?: boolean | undefined
  readonly valueAxis?: ChartValueAxis
  readonly xAxis?:
    | {
        min?: number | undefined
        max?: number | undefined
        majorUnit?: number | undefined
        numFmt?: string | undefined
        majorGridlines: boolean
      }
    | undefined
  readonly scatterStyle?: string | undefined
  readonly categoryFormat?: string | undefined
} & ChartElementProps): React.JSX.Element {
  const points = seriesList.map((series) => {
    const xValues = series.categories.map((value, index) => {
      const parsed = Number.parseFloat(value)
      return Number.isFinite(parsed) ? parsed : index
    })
    // Blank cache slots are 0 fillers, not data — Excel plots no point.
    const blankSet = new Set(series.blanks ?? [])
    return { series, xValues, blankSet }
  })
  const allX = points.flatMap((entry) =>
    entry.xValues.filter((_, index) => !entry.blankSet.has(index)),
  )
  const allY = points.flatMap((entry) =>
    entry.series.values.filter((_, index) => !entry.blankSet.has(index)),
  )
  if (allY.length === 0) return <></>
  const boundsX = scatterAxisBounds(allX, xAxis)
  const boundsY = scatterAxisBounds(allY, valueAxis)
  // lineMarker / smoothMarker / line / smooth connect the points; a series
  // whose spPr line is an explicit noFill stays marker-only.
  const styleWantsLines = scatterStyle !== undefined && scatterStyle !== 'marker'
  const xFormat = xAxis?.numFmt ?? categoryFormat
  const plotX = (value: number): number =>
    60 + Math.max(0, Math.min(1, (value - boundsX.min) / (boundsX.max - boundsX.min))) * 520
  const plotY = (value: number): number =>
    280 - Math.max(0, Math.min(1, (value - boundsY.min) / (boundsY.max - boundsY.min))) * 240
  return (
    <svg className="chart-svg" viewBox="0 0 600 320" role="img">
      <VerticalAxis
        minimum={boundsY.min}
        maximum={boundsY.max}
        ticks={boundsY.ticks}
        numberFormat={valueAxis?.numFmt ?? seriesList[0]?.numberFormat}
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
        <g key={index}>
          {xAxis?.majorGridlines && (
            <line
              x1={plotX(tick)}
              y1="40"
              x2={plotX(tick)}
              y2="280"
              stroke="#e3e3e3"
              strokeWidth="1"
            />
          )}
          <text
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
            {formatScatterTick(tick, xFormat)}
          </text>
        </g>
      ))}
      {points.map(({ series, xValues, blankSet }, seriesIndex) => (
        <g key={seriesIndex}>
          {styleWantsLines && series.lineColor !== 'none' && series.values.length > 1 && (
            <polyline
              points={series.values
                .map((value, index) =>
                  blankSet.has(index) ? null : `${plotX(xValues[index] ?? 0)},${plotY(value)}`,
                )
                .filter((point) => point !== null)
                .join(' ')}
              fill="none"
              stroke={
                series.lineColor && series.lineColor !== 'none'
                  ? series.lineColor
                  : seriesColor(series, seriesIndex)
              }
              strokeWidth={series.lineWidth ?? 2}
            />
          )}
          {series.marker !== 'none' &&
            series.values.map((value, index) =>
              blankSet.has(index) ? null : (
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
              ),
            )}
          {dataLabels === 'value' &&
            series.values.map((value, index) =>
              blankSet.has(index) ? null : (
                <text
                  key={`label-${index}`}
                  x={plotX(xValues[index] ?? 0)}
                  y={plotY(value) - 8}
                  textAnchor="middle"
                  className="axis-label"
                >
                  {formatAxisValue(value, series.numberFormat)}
                </text>
              ),
            )}
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
  /// dLbls' own numFmt only — the source cells' format is not a percent.
  percentFormat?: string | undefined,
): PieSliceLabel[] {
  const { cx, cy, r, inner, offsets } = geometry
  const labels: PieSliceLabel[] = []
  let cursor = 0
  for (const [index, value] of values.entries()) {
    const share = Math.max(0, value) / total
    const mid = (cursor + share / 2) * 2 * Math.PI
    cursor += share
    // Excel labels every visible slice; only slivers too thin to aim a
    // leader line at go unlabeled.
    if (share < 0.005) continue
    const percentText = formatPiePercent(share, percentFormat)
    const valueText = formatPieValue(value, formatCode)
    const lines =
      mode === 'category-value-percent'
        ? [truncateLabel(categories[index] ?? '', 12), valueText, percentText]
        : mode === 'category-percent'
          ? [truncateLabel(categories[index] ?? '', 12), percentText]
          : [mode === 'value' ? valueText : percentText]
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
          // Labels without their own numFmt inherit the source cells' format.
          dataLabelFormat ?? series.numberFormat,
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
          {categories.slice(0, 12).map((category, index) => (
            <span key={`${category}-${index}`}>
              <i style={{ background: pieSliceColor(series, index) }} />
              {truncateLabel(category, 12)}
            </span>
          ))}
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
