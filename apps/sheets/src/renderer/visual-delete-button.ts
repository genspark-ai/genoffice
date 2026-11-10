/**
 * Visibility and placement rules for the floating "delete this visual" button
 * on pictures, shapes and charts.
 *
 * Excel has no on-object delete affordance — a selected visual goes away via
 * Delete/Backspace or the context menu. Ours stays as a discoverability aid,
 * but only for the selected visual and only OUTSIDE its frame: the previous
 * always-mounted 20x20 button sat inside every visual's top-right corner,
 * covered small pictures and icons entirely, and turned a plain click on
 * them into a delete.
 */

import { BooleanNumber } from '@univerjs/core'

export interface BoxRect {
  readonly left: number
  readonly top: number
  readonly right: number
  readonly bottom: number
}

export const VISUAL_DELETE_BUTTON_SIZE = 20
/** Clearance between the selection frame and the button. */
export const VISUAL_DELETE_BUTTON_GAP = 4

export interface VisualDeleteButtonState {
  readonly selected: boolean
  readonly textEditing: boolean
}

/**
 * Only the selected visual carries the button, and never while its text is
 * being edited. Hovering an unselected visual does not count.
 */
export function shouldShowVisualDeleteButton(state: VisualDeleteButtonState): boolean {
  return state.selected && !state.textEditing
}

export type VisualDeleteButtonPlacement = 'above' | 'below' | 'inside'

export interface VisualDeleteButtonPosition {
  readonly left: number
  readonly top: number
  readonly placement: VisualDeleteButtonPlacement
}

/**
 * Anchors the button above the frame's top-right corner, outside the box.
 * `bounds` is the scrollable grid area (below the column header and any
 * frozen rows): when it leaves no room above, the button flips below the
 * bottom-right corner; when neither side fits (a frame taller than the
 * viewport) it falls back to just inside the visible top-right corner.
 * Horizontally it is clamped into the bounds so a visual at the right edge
 * keeps its button reachable.
 */
export function visualDeleteButtonPosition(
  frame: BoxRect,
  bounds: BoxRect,
  size = VISUAL_DELETE_BUTTON_SIZE,
  gap = VISUAL_DELETE_BUTTON_GAP,
): VisualDeleteButtonPosition {
  const above = frame.top - gap - size
  const below = frame.bottom + gap
  let placement: VisualDeleteButtonPlacement
  let top: number
  if (above >= bounds.top) {
    placement = 'above'
    top = above
  } else if (below + size <= bounds.bottom) {
    placement = 'below'
    top = below
  } else {
    placement = 'inside'
    top = Math.max(frame.top, bounds.top) + gap
  }
  const maxLeft = Math.max(bounds.left, bounds.right - size)
  const left = Math.min(Math.max(frame.right - size, bounds.left), maxLeft)
  return { left, top, placement }
}

/// The slice of Univer's FWorksheet needed to find the top of the scrolling
/// grid area (column header plus frozen rows, on screen).
export interface GridInsetSource {
  getZoom(): number
  getFreeze(): { readonly ySplit: number; readonly startRow: number }
  getRowHeight(row: number): number
  getSheet(): {
    getConfig(): {
      readonly columnHeader: { readonly height: number; readonly hidden?: BooleanNumber }
    }
    getRowVisible(row: number): boolean
  }
}

/**
 * On-screen distance from the top of the sheet canvas to the top of the
 * scrolling grid: the column header and the frozen rows, scaled by zoom.
 * Univer clamps float DOM to this line, so a visual scrolled under it is
 * exactly the case the "no room above" rule has to catch.
 */
export function gridTopInset(worksheet: GridInsetSource): number {
  const zoom = worksheet.getZoom() || 1
  const sheet = worksheet.getSheet()
  const header = sheet.getConfig().columnHeader
  let inset = header.hidden === BooleanNumber.TRUE ? 0 : header.height
  const freeze = worksheet.getFreeze()
  for (let row = freeze.startRow - freeze.ySplit; row < freeze.startRow; row += 1) {
    if (row >= 0 && sheet.getRowVisible(row)) inset += worksheet.getRowHeight(row)
  }
  return inset * zoom
}
