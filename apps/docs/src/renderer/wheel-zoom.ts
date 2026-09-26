import { clampZoom, notchStep, type ZoomWheelIntent } from '@genoffice/ui'

export const DOCS_ZOOM_MIN = 10
export const DOCS_ZOOM_MAX = 500

export function clampDocsZoom(zoom: number): number {
  return clampZoom(zoom, DOCS_ZOOM_MIN, DOCS_ZOOM_MAX)
}

/** Word's Zoom In / Zoom Out buttons and shortcuts: ten whole points per press. */
export function stepDocsZoom(zoom: number, dir: 1 | -1): number {
  return clampDocsZoom(Math.round(zoom) + dir * 10)
}

/** Word steps 10 percentage points per wheel notch; a pinch stays continuous. */
export function nextDocsZoom(
  zoom: number,
  intent: Exclude<ZoomWheelIntent, null>,
  deltaY: number,
): number {
  if (intent === 'pinch') return clampDocsZoom(zoom - deltaY * 0.6)
  const dir = intent === 'zoom-in' ? 1 : -1
  return Math.round(notchStep(zoom / 100, dir, DOCS_ZOOM_MIN / 100, DOCS_ZOOM_MAX / 100) * 100)
}

export type ZoomFitMode = 'width' | 'text' | 'page'
