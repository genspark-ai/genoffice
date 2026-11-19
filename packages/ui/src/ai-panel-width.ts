/**
 * Resizable AI-panel width: the same clamp/load/save contract every app's
 * AiPanel carried by hand (docs/slides/pdf/markdown/html/sheets). The storage
 * key stays per-app so existing persisted widths migrate unchanged; slides
 * keeps its width session-only and uses only the clamp.
 */

export const AI_PANEL_WIDTH_DEFAULT = 360
export const AI_PANEL_WIDTH_MIN = 280

/** Hard ceiling, matching the widest panel any app allowed. */
export const AI_PANEL_WIDTH_MAX = 720

export function maxAiPanelWidth(): number {
  // The viewport can be transiently tiny (a WebContentsView is 0×0 until the
  // shell lays it out), so never let the ceiling drop below the minimum
  return Math.max(
    AI_PANEL_WIDTH_MIN,
    Math.min(AI_PANEL_WIDTH_MAX, Math.round(window.innerWidth * 0.6)),
  )
}

export function clampAiPanelWidth(w: number): number {
  return Math.min(Math.max(w, AI_PANEL_WIDTH_MIN), maxAiPanelWidth())
}

/**
 * Restores the persisted preference, or null when none is stored. Static
 * bounds only — clamping against the window here would bake a transiently
 * small viewport into the restored preference.
 */
export function loadAiPanelWidth(key: string): number | null {
  const saved = Number(localStorage.getItem(key))
  return Number.isFinite(saved) && saved > 0
    ? Math.min(Math.max(saved, AI_PANEL_WIDTH_MIN), AI_PANEL_WIDTH_MAX)
    : null
}

export function saveAiPanelWidth(key: string, width: number): void {
  localStorage.setItem(key, String(Math.round(width)))
}
