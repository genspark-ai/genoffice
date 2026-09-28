/** Excel's zoom range; Univer's facade clamps to the same 0.1-4 ratio. */
export const SHEET_ZOOM_MIN = 10
export const SHEET_ZOOM_MAX = 400
export const ZOOM_STEP = 10

export function clampZoomPercent(percent: number): number {
  if (!Number.isFinite(percent)) return 100
  return Math.min(SHEET_ZOOM_MAX, Math.max(SHEET_ZOOM_MIN, Math.round(percent)))
}

export function clampZoomRatio(ratio: number): number {
  return clampZoomPercent(ratio * 100) / 100
}

export function stepZoomPercent(percent: number, direction: 1 | -1): number {
  return clampZoomPercent(percent + direction * ZOOM_STEP)
}
