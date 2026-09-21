import { notchStep, type ZoomWheelIntent } from '@genoffice/ui'

export const DOCS_ZOOM_MIN = 50
export const DOCS_ZOOM_MAX = 200

/** Word steps 10 percentage points per wheel notch; a pinch stays continuous. */
export function nextDocsZoom(
  zoom: number,
  intent: Exclude<ZoomWheelIntent, null>,
  deltaY: number,
): number {
  if (intent === 'pinch') {
    return Math.min(DOCS_ZOOM_MAX, Math.max(DOCS_ZOOM_MIN, zoom - deltaY * 0.6))
  }
  const dir = intent === 'zoom-in' ? 1 : -1
  return Math.round(notchStep(zoom / 100, dir, DOCS_ZOOM_MIN / 100, DOCS_ZOOM_MAX / 100) * 100)
}
