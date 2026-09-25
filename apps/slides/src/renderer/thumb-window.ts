/**
 * Thumbnail-rail windowing (issue #763).
 *
 * The rail used to mount a Konva Stage — one DOM canvas — for every slide, so a
 * 409-slide deck painted 409 of them on first paint, each materializing its
 * slide's picture/effect canvases at full slide scale even though a thumbnail is
 * ~126 px wide. Only the rows that intersect the viewport are mounted now; the
 * rows outside it are represented by spacers of the same height, so the scroll
 * geometry and the drop targets stay exactly where they were.
 *
 * Everything here is pure so the maths can be tested without a browser.
 */

/** `.thumb` border: 2px on each side (styles.css) */
export const THUMB_BORDER_PX = 4
/** `.thumb` margin-bottom (styles.css) */
export const THUMB_GAP_PX = 10
/** Rows rendered beyond the viewport, in px, so fast scrolling never shows a blank rail */
export const DEFAULT_OVERSCAN_PX = 600
/** Below this many slides the rail renders everything: windowing a short deck buys nothing and every extra code path is risk */
export const WINDOWING_MIN_SLIDES = 60

/** Height of one thumbnail row: the scaled stage plus its border and margin. */
export function thumbRowHeight(
  slide: { widthPx: number; heightPx: number },
  width: number,
): number {
  const scale = width / Math.max(1, slide.widthPx)
  return Math.max(1, Math.round(slide.heightPx * scale)) + THUMB_BORDER_PX + THUMB_GAP_PX
}

/** `offsets[i]` is the top of row i; the last entry is the total height. */
export function rowOffsets(heights: number[]): number[] {
  const offsets = new Array<number>(heights.length + 1)
  offsets[0] = 0
  for (let i = 0; i < heights.length; i++) offsets[i + 1] = offsets[i] + heights[i]
  return offsets
}

/**
 * Half-open range of rows intersecting `[scrollTop - overscan, scrollTop + viewportHeight + overscan]`.
 * `end` is exclusive; an empty list yields an empty range at 0.
 */
export function visibleRowRange(
  offsets: number[],
  scrollTop: number,
  viewportHeight: number,
  overscan = DEFAULT_OVERSCAN_PX,
): { start: number; end: number } {
  const count = Math.max(0, offsets.length - 1)
  if (count === 0) return { start: 0, end: 0 }
  const top = Math.max(0, scrollTop - overscan)
  const bottom = scrollTop + Math.max(0, viewportHeight) + overscan
  let start = 0
  // linear scan is fine: the rail is scrolled, not searched
  while (start < count && offsets[start + 1] <= top) start += 1
  let end = start
  while (end < count && offsets[end] < bottom) end += 1
  return { start, end }
}

/**
 * Row index whose insertion point is nearest `contentY` (content coordinates, i.e.
 * list scroll offset included). Used for the rail's blank-space context menu when
 * the far rows are not mounted and therefore have no element to measure.
 */
export function insertionIndexAt(offsets: number[], contentY: number): number {
  const count = Math.max(0, offsets.length - 1)
  for (let i = 0; i < count; i++) {
    const mid = (offsets[i] + offsets[i + 1]) / 2
    if (contentY < mid) return i
  }
  return count
}
