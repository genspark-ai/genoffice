/**
 * Pure geometry shared by the tab strip renderer (tear-off gesture, dock
 * insertion indicator) and the main process (native window-drag docking).
 * No Electron or DOM imports so both sides — and unit tests — can use it.
 */

/** rendered height of the shell's tab strip (apps/shell/src/renderer/src/tabbar.css) */
export const TAB_STRIP_HEIGHT = 40

/**
 * How far (px) the pointer must leave the strip's vertical band before a
 * drag-to-reorder becomes a tear-off. Chrome uses a similar dead zone so a
 * wobbly horizontal drag never accidentally detaches.
 */
export const TEAR_OFF_SLACK = 28

/**
 * Vertical tolerance (px) around the strip when a detached window is dragged
 * back over it: the cursor sits in the dragged window's title bar, so a few
 * pixels of forgiveness keep the target easy to hit.
 */
export const DOCK_BAND_SLACK = 8

/** Hover dwell (ms) over the strip before a natively dragged window docks. */
export const DOCK_DWELL_MS = 180

export interface Rect {
  x: number
  y: number
  width: number
  height: number
}

/** true when `y` is more than `slack` px outside the [top, bottom] band */
export function isBeyondBand(y: number, top: number, bottom: number, slack: number): boolean {
  return y < top - slack || y > bottom + slack
}

export function pointInRect(x: number, y: number, rect: Rect): boolean {
  return x >= rect.x && x < rect.x + rect.width && y >= rect.y && y < rect.y + rect.height
}

/**
 * The screen rectangle a detached window's cursor must hover to dock into
 * the shell: the tab strip at the top of the shell window's content area,
 * grown by `slack` above and below.
 */
export function dockBand(shellContentBounds: Rect, slack: number = DOCK_BAND_SLACK): Rect {
  return {
    x: shellContentBounds.x,
    y: shellContentBounds.y - slack,
    width: shellContentBounds.width,
    height: TAB_STRIP_HEIGHT + 2 * slack,
  }
}

export interface TabRect {
  left: number
  width: number
}

/**
 * Insertion slot for a tab dropped at horizontal position `x`, given the
 * strip's current tab rectangles in order. Slot 0 is Home (pinned), so the
 * result is clamped to [1, rects.length]: `x` left of a tab's midpoint means
 * "before it", past every midpoint means "append".
 */
export function insertionIndexForX(rects: readonly TabRect[], x: number): number {
  for (let i = 1; i < rects.length; i++) {
    if (x < rects[i].left + rects[i].width / 2) return i
  }
  return Math.max(1, rects.length)
}
