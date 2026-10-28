import { aiPanelAutoWidth, aiPanelClampWidth } from '@genoffice/ui'

/** localStorage key of the user-dragged panel width; a null preference means auto width */
export const PANEL_WIDTH_KEY = 'markdown-ai-panel-width'

/**
 * The dragged width exactly as saved. Static bounds only — clamping against
 * the window here would bake a transiently small viewport into the restored
 * preference; the display width derives it per render instead.
 */
export function loadSavedPanelWidth(): number | null {
  const saved = Number(localStorage.getItem(PANEL_WIDTH_KEY))
  return Number.isFinite(saved) && saved > 0 ? saved : null
}

/**
 * Width actually shown for the current window: the dragged preference clamped
 * into it (legacy values below the floor migrate up via the clamp), or — when
 * the user never dragged — the auto width that grows with the window.
 */
export function displayPanelWidth(preferred: number | null, windowWidth: number): number {
  return preferred === null
    ? aiPanelAutoWidth(windowWidth)
    : aiPanelClampWidth(preferred, windowWidth)
}
