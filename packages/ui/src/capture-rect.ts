/**
 * Locating the page element for Electron's `capturePage`, shared by every
 * editor that can render itself for the agent's `view_page` tool. Only the
 * selector differs between them (docs `.doc-page`, slides `.stage-rel`), so
 * the geometry lives here rather than being copied per app.
 */

export interface CaptureRect {
  x: number
  y: number
  width: number
  height: number
}

/**
 * Where the first element matching `selector` sits, in the viewport-relative
 * coordinates `capturePage` expects.
 *
 * Clipped to the visible area: a page taller than the window would otherwise
 * ask for a rect running off the bottom edge, and capturePage answers that
 * with an empty image. Returns undefined when there is nothing to capture —
 * no such element, or it is scrolled fully out of view — which the caller
 * reports rather than sending a blank image to the model.
 */
export function captureRect(selector: string): CaptureRect | undefined {
  const el = document.querySelector(selector)
  if (!el) return undefined
  const r = el.getBoundingClientRect()
  const x = Math.max(0, r.left)
  const y = Math.max(0, r.top)
  const width = Math.min(r.right, window.innerWidth) - x
  const height = Math.min(r.bottom, window.innerHeight) - y
  return width >= 1 && height >= 1 ? { x, y, width, height } : undefined
}
