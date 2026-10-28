// Shared responsive width rule for the apps' AI side panels, replacing the
// copy-pasted per-app constants: the panel auto-grows and auto-shrinks with
// the window but never drops below the default width. The shell window's
// minimum width is 720, so the 60% cap is at least 432 there — the default
// floor is always satisfiable.

/** Default AI panel width, and the floor the panel never shrinks below */
export const AI_PANEL_WIDTH_DEFAULT = 360

/** Absolute ceiling for the panel width, however wide the window gets */
export const AI_PANEL_WIDTH_MAX = 720

/** Hard cap: the panel never exceeds 60% of the window or 720px */
export function aiPanelWidthCap(windowWidth: number): number {
  const w = saneWindowWidth(windowWidth)
  return Math.max(AI_PANEL_WIDTH_DEFAULT, Math.min(AI_PANEL_WIDTH_MAX, Math.round(w * 0.6)))
}

/** Width when the user has not dragged the divider: grows with the window
 * (~22%), floored at the default, capped by aiPanelWidthCap */
export function aiPanelAutoWidth(windowWidth: number): number {
  const w = saneWindowWidth(windowWidth)
  const cap = aiPanelWidthCap(w)
  return Math.min(cap, Math.max(AI_PANEL_WIDTH_DEFAULT, Math.round(w * 0.22)))
}

/** Clamp a user-dragged width: never below the default, never above the cap */
export function aiPanelClampWidth(width: number, windowWidth: number): number {
  const cap = aiPanelWidthCap(windowWidth)
  return Math.min(cap, Math.max(AI_PANEL_WIDTH_DEFAULT, width))
}

/** NaN, negative or infinite window widths count as a zero-size window, so every rule falls back to the default width */
function saneWindowWidth(windowWidth: number): number {
  return Number.isFinite(windowWidth) && windowWidth > 0 ? windowWidth : 0
}
