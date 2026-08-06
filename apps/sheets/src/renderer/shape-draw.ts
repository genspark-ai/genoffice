/**
 * Excel-parity shape draw mode for the grid: picking a shape in the ribbon
 * gallery arms a crosshair over the sheet; a single click inserts the
 * predefined 1x1 inch shape at the click point, a drag draws a custom-size one
 * (Shift constrains to a square, Esc cancels). The gesture rectangle is
 * resolved into a twoCellAnchor by walking real row/column sizes from the
 * viewport's top-left cell.
 */
import { shapeClipCss } from '@genoffice/ui'
import type { UniverRuntime } from './univer-state'
import { EMU_PER_PIXEL, walkMarker, type AnchorMarker } from './WorkbookVisuals'
import type { WorkbookVisualObject } from '../shared/desktop-api'

/** Excel's predefined single-click insert size: 1x1 inch (96 CSS px). */
const DEFAULT_SHAPE_PX = 96
const CLICK_THRESHOLD_PX = 3
const XLSX_MAX_COLUMN = 16383
const XLSX_MAX_ROW = 1048575

export interface DrawRectPx {
  x: number
  y: number
  w: number
  h: number
}

/** Drag rectangle in viewport px; Shift locks a square grown in the drag direction. */
export function resolveDrawRect(
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  shift: boolean,
): DrawRectPx {
  let dx = x2 - x1
  let dy = y2 - y1
  if (shift) {
    const side = Math.max(Math.abs(dx), Math.abs(dy))
    dx = dx < 0 ? -side : side
    dy = dy < 0 ? -side : side
  }
  return {
    x: Math.min(x1, x1 + dx),
    y: Math.min(y1, y1 + dy),
    w: Math.abs(dx),
    h: Math.abs(dy),
  }
}

/** Marker pair → xlsx twoCellAnchor fields. */
function anchorFromMarkers(
  fromX: AnchorMarker,
  fromY: AnchorMarker,
  toX: AnchorMarker,
  toY: AnchorMarker,
): WorkbookVisualObject['anchor'] {
  return {
    fromRow: fromY.index,
    fromColumn: fromX.index,
    fromRowOffset: Math.round(fromY.offset * EMU_PER_PIXEL),
    fromColumnOffset: Math.round(fromX.offset * EMU_PER_PIXEL),
    toRow: toY.index,
    toColumn: toX.index,
    toRowOffset: Math.round(toY.offset * EMU_PER_PIXEL),
    toColumnOffset: Math.round(toX.offset * EMU_PER_PIXEL),
  }
}

/**
 * Screen rectangle → twoCellAnchor. The viewport's top-left visible cell gives
 * an on-screen reference (its DOMRect already reflects partial scroll and
 * frozen panes); real row/column sizes are walked from there.
 */
export function rectToAnchor(
  runtime: UniverRuntime,
  rect: DrawRectPx,
): WorkbookVisualObject['anchor'] | null {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!workbook || !worksheet) return null
  // getCellRect is relative to the grid's render surface — the largest canvas
  // in the container (the container itself has extra chrome above the grid)
  const hostEl = document.getElementById('univer-container')
  if (!hostEl) return null
  let surface: DOMRect | null = null
  for (const canvas of hostEl.querySelectorAll('canvas')) {
    const r = canvas.getBoundingClientRect()
    if (!surface || r.width * r.height > surface.width * surface.height) surface = r
  }
  if (!surface) return null
  let visible: { startRow: number; startColumn: number } | null
  let cellX: number
  let cellY: number
  try {
    visible = worksheet.getVisibleRange()
    if (!visible) return null
    const cellRect = worksheet.getRange(visible.startRow, visible.startColumn, 1, 1).getCellRect()
    cellX = surface.x + cellRect.x
    cellY = surface.y + cellRect.y
  } catch {
    return null
  }
  const zoom = worksheet.getZoom() || 1
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
  const originX: AnchorMarker = { index: visible.startColumn, offset: 0 }
  const originY: AnchorMarker = { index: visible.startRow, offset: 0 }
  const fromX = walkMarker(originX, (rect.x - cellX) / zoom, columnWidth, XLSX_MAX_COLUMN)
  const fromY = walkMarker(originY, (rect.y - cellY) / zoom, rowHeight, XLSX_MAX_ROW)
  const toX = walkMarker(fromX, Math.max(rect.w, 1) / zoom, columnWidth, XLSX_MAX_COLUMN)
  const toY = walkMarker(fromY, Math.max(rect.h, 1) / zoom, rowHeight, XLSX_MAX_ROW)
  return anchorFromMarkers(fromX, fromY, toX, toY)
}

/**
 * The active-mode cancel lives on window (not module scope): dev HMR swaps the
 * module while a mode is armed, and the stale instance's listeners must still
 * be torn down by the replacement before it arms its own.
 */
const CANCEL_KEY = '__sheetShapeDrawCancel'
// Lazy: this module is imported by node-env test suites where window is absent
const cancelStore = (): Record<string, (() => void) | undefined> =>
  window as unknown as Record<string, (() => void) | undefined>

/** Arm the crosshair draw mode over the grid; commit receives the drawn anchor. */
export function startSheetShapeDraw(
  runtime: UniverRuntime,
  prst: string,
  commit: (anchor: WorkbookVisualObject['anchor']) => void,
): void {
  cancelStore()[CANCEL_KEY]?.()
  const host = document.getElementById('univer-container')
  if (!host) return
  host.classList.add('sheet-shape-drawing')

  let start: { x: number; y: number } | null = null
  let cur: { x: number; y: number } | null = null
  let shift = false
  let ghost: HTMLDivElement | null = null

  const updateGhost = () => {
    if (!start || !cur) return
    if (Math.hypot(cur.x - start.x, cur.y - start.y) <= CLICK_THRESHOLD_PX) return
    const r = resolveDrawRect(start.x, start.y, cur.x, cur.y, shift)
    if (!ghost) {
      ghost = document.createElement('div')
      ghost.style.position = 'fixed'
      ghost.style.zIndex = '9999'
      ghost.style.pointerEvents = 'none'
      // Ghost of the default light-blue shape the gesture will insert
      ghost.style.background = 'rgba(221,235,247,0.75)'
      ghost.style.border = '1px solid #9db8d4'
      ghost.style.boxSizing = 'border-box'
      document.body.appendChild(ghost)
    }
    const clip = shapeClipCss(prst, r.w, r.h)
    ghost.style.clipPath = clip?.clipPath ?? ''
    ghost.style.borderRadius = clip?.borderRadius ?? ''
    ghost.style.left = `${r.x}px`
    ghost.style.top = `${r.y}px`
    ghost.style.width = `${r.w}px`
    ghost.style.height = `${r.h}px`
  }

  const onMouseDown = (e: MouseEvent) => {
    if (e.button !== 0) return
    e.preventDefault()
    e.stopPropagation()
    start = { x: e.clientX, y: e.clientY }
    cur = { ...start }
    shift = e.shiftKey
    window.addEventListener('mousemove', onWinMove, true)
    window.addEventListener('mouseup', onWinUp, true)
  }
  const onWinMove = (e: MouseEvent) => {
    if (!start) return
    cur = { x: e.clientX, y: e.clientY }
    shift = e.shiftKey
    updateGhost()
  }
  const onWinUp = (e: MouseEvent) => {
    if (!start) return
    const s = start
    const isClick = Math.hypot(e.clientX - s.x, e.clientY - s.y) <= CLICK_THRESHOLD_PX
    const rect = isClick
      ? { x: s.x, y: s.y, w: DEFAULT_SHAPE_PX, h: DEFAULT_SHAPE_PX }
      : resolveDrawRect(s.x, s.y, e.clientX, e.clientY, e.shiftKey)
    cleanup()
    const anchor = rectToAnchor(runtime, rect)
    if (anchor) commit(anchor)
  }
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape') {
      cleanup()
      return
    }
    if (e.key === 'Shift' && start) {
      shift = e.type === 'keydown'
      updateGhost()
    }
  }

  const cleanup = () => {
    delete cancelStore()[CANCEL_KEY]
    host.classList.remove('sheet-shape-drawing')
    host.removeEventListener('mousedown', onMouseDown, true)
    window.removeEventListener('mousemove', onWinMove, true)
    window.removeEventListener('mouseup', onWinUp, true)
    window.removeEventListener('keydown', onKey, true)
    window.removeEventListener('keyup', onKey, true)
    ghost?.remove()
    ghost = null
    start = null
  }

  host.addEventListener('mousedown', onMouseDown, true)
  window.addEventListener('keydown', onKey, true)
  window.addEventListener('keyup', onKey, true)
  cancelStore()[CANCEL_KEY] = cleanup
}
