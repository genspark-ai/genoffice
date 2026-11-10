export interface Point {
  x: number
  y: number
}

export interface ViewportBounds {
  left: number
  top: number
  right: number
  bottom: number
}

export interface SelectionAskPosition {
  left: number
  top: number
}

export interface SelectionAskAnchor {
  pointer: Point
  bounds: ViewportBounds
}

export const SELECTION_DRAG_THRESHOLD_PX = 4

/** Ignore pointer jitter so an ordinary cell click never summons Ask AI. */
export function isSelectionDrag(
  start: Point,
  current: Point,
  threshold = SELECTION_DRAG_THRESHOLD_PX,
): boolean {
  return Math.hypot(current.x - start.x, current.y - start.y) >= threshold
}

/** Keep the trigger beside the drag endpoint and inside the grid viewport. */
export function selectionAskPosition(
  pointer: Point,
  bounds: ViewportBounds,
  width: number,
  height: number,
  gap = 8,
): SelectionAskPosition {
  const inset = 4
  const minLeft = bounds.left + inset
  const maxLeft = Math.max(minLeft, bounds.right - width - inset)
  const minTop = bounds.top + inset
  const maxTop = Math.max(minTop, bounds.bottom - height - inset)
  const preferredLeft = pointer.x + gap
  const preferredTop = pointer.y + gap
  const flippedLeft = pointer.x - width - gap
  const flippedTop = pointer.y - height - gap

  return {
    left: Math.min(
      Math.max(
        preferredLeft + width <= bounds.right - inset ? preferredLeft : flippedLeft,
        minLeft,
      ),
      maxLeft,
    ),
    top: Math.min(
      Math.max(preferredTop + height <= bounds.bottom - inset ? preferredTop : flippedTop, minTop),
      maxTop,
    ),
  }
}
