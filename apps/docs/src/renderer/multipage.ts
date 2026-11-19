/**
 * Word's Multiple Pages layout: in Print Layout the pages flow into as many
 * columns as the pane is wide for the current zoom (two pages side by side at
 * ~50%, three at ~33%, ...). The column count is a function of zoom × pane
 * width, never a separate view mode: zooming back in collapses the grid to the
 * usual single column.
 */
import { GAP_BAND } from './editor/pagination-gaps'

/** horizontal space between page columns (= the inter-page band of the canvas) */
export const MULTIPAGE_GUTTER = GAP_BAND
/** scroller padding reserved on both sides, same as the width-fit zoom */
export const MULTIPAGE_PANE_PAD = 48

/** Columns that fit: `paneW` in CSS px, `slotW` in unzoomed canvas px, zoom in % */
export function multipageColumns(paneW: number, slotW: number, zoom: number): number {
  if (!(paneW > 0) || !(slotW > 0) || !(zoom > 0)) return 1
  const avail = (paneW - MULTIPAGE_PANE_PAD) / (zoom / 100)
  return Math.max(1, Math.floor((avail + MULTIPAGE_GUTTER) / (slotW + MULTIPAGE_GUTTER)))
}

/** Zoom (%) at which exactly `cols` pages fit across the pane */
export function multipageZoomForColumns(paneW: number, slotW: number, cols: number): number {
  const n = Math.max(1, Math.floor(cols))
  return ((paneW - MULTIPAGE_PANE_PAD) / (n * slotW + (n - 1) * MULTIPAGE_GUTTER)) * 100
}

export interface MultipageSlot {
  left: number
  top: number
  width: number
  height: number
}

export interface MultipageLayout {
  cols: number
  slots: MultipageSlot[]
  width: number
  height: number
}

/**
 * Row-major grid of page slots (unzoomed canvas px). Pages in a row share the
 * row's tallest height; every slot is `slotW` wide (differing paper sizes
 * center within it like the canvas centers them on the wrap).
 */
export function layoutMultipage(
  pageHeights: readonly number[],
  cols: number,
  slotW: number,
): MultipageLayout {
  const n = Math.max(1, Math.floor(cols))
  const slots: MultipageSlot[] = []
  let top = 0
  for (let r = 0; r * n < pageHeights.length; r++) {
    const row = pageHeights.slice(r * n, r * n + n)
    const rowH = Math.max(...row)
    row.forEach((h, c) => {
      slots.push({ left: c * (slotW + MULTIPAGE_GUTTER), top, width: slotW, height: h })
    })
    top += rowH + MULTIPAGE_GUTTER
  }
  const usedCols = Math.min(n, pageHeights.length)
  return {
    cols: n,
    slots,
    width: Math.max(0, usedCols * slotW + (usedCols - 1) * MULTIPAGE_GUTTER),
    height: Math.max(0, top - MULTIPAGE_GUTTER),
  }
}

/** index of the slot containing canvas point (x, y), or the nearest row's last slot */
export function multipageSlotAt(layout: MultipageLayout, x: number, y: number): number {
  const { slots } = layout
  if (slots.length === 0) return 0
  let best = 0
  for (let i = 0; i < slots.length; i++) {
    const s = slots[i]
    if (y >= s.top && y < s.top + s.height + MULTIPAGE_GUTTER) {
      best = i
      if (x < s.left + s.width + MULTIPAGE_GUTTER) return i
    }
  }
  return best
}
