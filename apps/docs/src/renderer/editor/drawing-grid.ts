/**
 * The drawing grid of Format > Document > Document Grid > Drawing Grid:
 * what View > Gridlines draws and what a dragged picture, shape or text box
 * snaps to. The grid starts at the first page's margin corner (or its own
 * origin) and runs down the page canvas.
 */
import type { DrawingGrid } from '@genoffice/docx-engine'

const TWIPS_PER_PX = 15
/** an eighth of an inch: the spacing a file without one is drawn with */
export const DEFAULT_GRID_SPACING_TWIPS = 180
export const DEFAULT_GRID_DISPLAY_EVERY = 2
export const DEFAULT_GRID_ORIGIN_TWIPS = { horizontal: 1800, vertical: 1440 }
const MIN_SPACING_TWIPS = 15
const SNAP_KEY = 'aidocs.snapToDrawingGrid'

export interface GridGeometry {
  /** px between grid points */
  stepX: number
  stepY: number
  /** a vertical gridline every n horizontal steps; 0 draws none */
  everyX: number
  /** a horizontal gridline every n vertical steps */
  everyY: number
  /** px from the page edge, when the grid does not start at the margins */
  origin: { x: number; y: number } | null
}

const spacingOf = (twips: number | undefined): number =>
  (twips !== undefined && twips >= MIN_SPACING_TWIPS ? twips : DEFAULT_GRID_SPACING_TWIPS) /
  TWIPS_PER_PX

export function gridGeometry(grid: DrawingGrid | null | undefined): GridGeometry {
  return {
    stepX: spacingOf(grid?.horizontalSpacing),
    stepY: spacingOf(grid?.verticalSpacing),
    everyX: Math.max(0, grid?.displayHorizontalEvery ?? DEFAULT_GRID_DISPLAY_EVERY),
    // horizontal lines have no off state of their own
    everyY: Math.max(1, grid?.displayVerticalEvery ?? DEFAULT_GRID_DISPLAY_EVERY),
    origin: grid?.ownOrigin
      ? {
          x: (grid.horizontalOrigin ?? 0) / TWIPS_PER_PX,
          y: (grid.verticalOrigin ?? 0) / TWIPS_PER_PX,
        }
      : null,
  }
}

const px = (n: number) => `${Math.round(n * 1000) / 1000}px`

/** the declarations View > Gridlines reads on .doc-page for a document that states a grid */
export function gridlinesCss(grid: DrawingGrid | null | undefined): string {
  if (!grid) return ''
  const g = gridGeometry(grid)
  return [
    `--docs-grid-x:${px(g.stepX * Math.max(1, g.everyX))}`,
    `--docs-grid-y:${px(g.stepY * g.everyY)}`,
    ...(g.everyX === 0 ? ['--docs-grid-columns:none'] : []),
    ...(g.origin
      ? ['--docs-grid-box:padding-box', `--docs-grid-at:${px(g.origin.x)} ${px(g.origin.y)}`]
      : []),
  ].join('; ')
}

/** the grid point nearest to `value` on an axis that starts at `origin` */
export const snapToStep = (value: number, origin: number, step: number): number =>
  step > 0 ? origin + Math.round((value - origin) / step) * step : value

let active = gridGeometry(null)

export function setActiveDrawingGrid(grid: DrawingGrid | null | undefined): void {
  active = gridGeometry(grid)
}

/** "Snap objects to grid": kept on this computer, like Word keeps it out of the file */
export function snapsToGrid(): boolean {
  try {
    return localStorage.getItem(SNAP_KEY) === '1'
  } catch {
    return false
  }
}

export function setSnapsToGrid(on: boolean): void {
  try {
    localStorage.setItem(SNAP_KEY, on ? '1' : '0')
  } catch {
    // private mode: the setting lasts for the session's reads only
  }
}

interface ClientGrid {
  x: number
  y: number
  stepX: number
  stepY: number
}

/** the grid in client px on the page canvas that holds `el` */
function clientGrid(el: Element): ClientGrid | null {
  const page = el.closest('.doc-page')
  if (!(page instanceof HTMLElement)) return null
  const zoomEl = page.closest('.doc-zoom')
  const zoom = zoomEl ? parseFloat(getComputedStyle(zoomEl).zoom || '1') || 1 : 1
  const rect = page.getBoundingClientRect()
  const style = getComputedStyle(page)
  const inset = (edge: 'Left' | 'Top') =>
    active.origin
      ? (edge === 'Left' ? page.clientLeft : page.clientTop) +
        (edge === 'Left' ? active.origin.x : active.origin.y)
      : (edge === 'Left' ? page.clientLeft : page.clientTop) +
        (parseFloat(style[`padding${edge}`]) || 0)
  return {
    x: rect.left + inset('Left') * zoom,
    y: rect.top + inset('Top') * zoom,
    stepX: active.stepX * zoom,
    stepY: active.stepY * zoom,
  }
}

type DragEvent = Pick<MouseEvent, 'clientX' | 'clientY'>

/** the pointer of `e`, moved onto the nearest grid point */
export function snappedPoint(el: Element, e: DragEvent): { x: number; y: number } {
  const grid = snapsToGrid() ? clientGrid(el) : null
  if (!grid) return { x: e.clientX, y: e.clientY }
  return {
    x: snapToStep(e.clientX, grid.x, grid.stepX),
    y: snapToStep(e.clientY, grid.y, grid.stepY),
  }
}

/**
 * The pointer of a move drag that began at `start` with the object's corner
 * at `corner`: where the pointer has to be for the corner to sit on the grid.
 */
export function snappedDragPointer(
  el: Element,
  start: { x: number; y: number },
  corner: { left: number; top: number },
  e: DragEvent,
): { x: number; y: number } {
  const grid = snapsToGrid() ? clientGrid(el) : null
  if (!grid) return { x: e.clientX, y: e.clientY }
  const left = corner.left + (e.clientX - start.x)
  const top = corner.top + (e.clientY - start.y)
  return {
    x: e.clientX + snapToStep(left, grid.x, grid.stepX) - left,
    y: e.clientY + snapToStep(top, grid.y, grid.stepY) - top,
  }
}
