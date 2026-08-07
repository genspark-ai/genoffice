/**
 * Geometry for the shape draw mode (PowerPoint/WPS parity): after picking a shape in
 * the ribbon gallery the cursor becomes a crosshair — a single click inserts the
 * PowerPoint default size (1x1 inch square), a drag draws a custom rectangle, and
 * holding Shift constrains shapes to a square / lines to 45-degree increments.
 * Coordinates are slide px in the FIT_WIDTH viewport (96px = 1 inch).
 */
import { PX_PER_INCH } from './app-constants'

/** Inserted box + optional mirroring (lines only: connectors render top-left → bottom-right, flips restore the drag direction). */
export interface DrawRect {
  x: number
  y: number
  w: number
  h: number
  flipH?: boolean
  flipV?: boolean
}

const STRAIGHT_LINE_KINDS = new Set(['line', 'lineArrow', 'lineArrowDouble'])
const LINE_KINDS = new Set([...STRAIGHT_LINE_KINDS, 'lineBent', 'lineCurved'])

export function isLineDrawKind(kind: string): boolean {
  return LINE_KINDS.has(kind)
}

export function isStraightLineKind(kind: string): boolean {
  return STRAIGHT_LINE_KINDS.has(kind)
}

/** Single-click insert size: PowerPoint's predefined 1x1 inch (straight lines: 1 inch horizontal stroke). */
export function defaultDrawSize(kind: string): { w: number; h: number } {
  if (isStraightLineKind(kind)) return { w: PX_PER_INCH, h: 0 }
  return { w: PX_PER_INCH, h: PX_PER_INCH }
}

/**
 * Resolve a drag gesture (anchor x1/y1 → pointer x2/y2) into the box to insert.
 * Shapes: Shift locks a square on the larger delta, growing in the drag direction.
 * Lines: Shift snaps the segment to the nearest 45°; the box spans the endpoints and
 * flipH/flipV mark drags that go left/up so the connector (and its arrowheads) keeps
 * the drawn direction.
 */
export function resolveDrawRect(
  kind: string,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  shift: boolean,
): DrawRect {
  let dx = x2 - x1
  let dy = y2 - y1
  if (isLineDrawKind(kind)) {
    if (shift) {
      const len = Math.hypot(dx, dy)
      const angle = (Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * Math.PI) / 4
      dx = Math.round(len * Math.cos(angle))
      dy = Math.round(len * Math.sin(angle))
      // Kill float noise on the axis-aligned snaps so horizontal/vertical are exact
      if (Math.abs(dx) < 1) dx = 0
      if (Math.abs(dy) < 1) dy = 0
    }
    const rect: DrawRect = {
      x: Math.min(x1, x1 + dx),
      y: Math.min(y1, y1 + dy),
      w: Math.abs(dx),
      h: Math.abs(dy),
    }
    if (dx < 0) rect.flipH = true
    if (dy < 0) rect.flipV = true
    return rect
  }
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
