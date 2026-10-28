import { aiPanelAutoWidth, aiPanelClampWidth } from '@genoffice/ui'

/** localStorage key of the user's dragged AI panel width (written by the panel resizer) */
export const PANEL_WIDTH_KEY = 'sheets-ai-panel-width'

/** The user's dragged width, or null when they never dragged (the panel then
 * tracks the window). Static sanity only: clamping against the live window
 * happens at display time, so a transiently tiny viewport never bakes into the
 * restored preference; legacy values below the 360 floor migrate up through
 * aiPanelClampWidth when displayed. */
export function loadPanelWidth(storage: Storage = localStorage): number | null {
  const saved = Number(storage.getItem(PANEL_WIDTH_KEY))
  return Number.isFinite(saved) && saved > 0 ? saved : null
}

/** Display width of the panel column: the dragged preference clamped to the
 * live window (never below the 360 floor), or the window-tracking auto width
 * when the user has no preference. */
export function panelDisplayWidth(preferred: number | null, windowWidth: number): number {
  return preferred === null
    ? aiPanelAutoWidth(windowWidth)
    : aiPanelClampWidth(preferred, windowWidth)
}
