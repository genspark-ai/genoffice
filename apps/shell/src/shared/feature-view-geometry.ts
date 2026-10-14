/**
 * Geometry and address rules for the native view slot (K5).
 *
 * Kept free of electron so it can be unit-tested: the main process asks for a
 * clamped rectangle, and the renderer asks for a host name to show while the
 * view attaches. Which addresses are allowed lives with the contract, in
 * packages/agent-core (isAllowedViewUrl), because that is where a slot is judged.
 */

export interface ViewRect {
  x: number
  y: number
  width: number
  height: number
}

export interface ViewBounds extends ViewRect {
  // the window's content area, used as the clamp
  limitWidth: number
  limitHeight: number
}

/** A view smaller than this is not worth mounting (the panel is collapsed). */
export const MIN_VIEW_SIZE = 80

function whole(value: unknown): number {
  return typeof value === 'number' && Number.isFinite(value) ? Math.round(value) : 0
}

/**
 * Fit a renderer-reported rectangle into the window's content area. Anything
 * that falls outside, or that is too small to be a panel, comes back as null —
 * the caller then simply does not mount a view.
 */
export function clampViewRect(
  rect: unknown,
  area: { width: number; height: number },
): ViewRect | null {
  if (!rect || typeof rect !== 'object') return null
  const raw = rect as Record<string, unknown>
  const x = Math.max(0, whole(raw.x))
  const y = Math.max(0, whole(raw.y))
  const width = Math.min(whole(raw.width), Math.max(0, whole(area.width) - x))
  const height = Math.min(whole(raw.height), Math.max(0, whole(area.height) - y))
  if (width < MIN_VIEW_SIZE || height < MIN_VIEW_SIZE) return null
  return { x, y, width, height }
}
