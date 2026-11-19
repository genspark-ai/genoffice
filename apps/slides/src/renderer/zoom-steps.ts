// PowerPoint zoom stepping: +/- walk a preset ladder, a Ctrl+wheel notch and the
// status-bar buttons move 10 percentage points, everything clamps to 10-400%.

import { clampZoom as clampZoomTo, notchStep as notchStepTo } from '@genoffice/ui'

export const ZOOM_MIN = 0.1
export const ZOOM_MAX = 4
export const ZOOM_PRESETS = [0.1, 0.25, 0.33, 0.5, 0.66, 0.75, 1, 1.5, 2, 3, 4] as const

/** Fit values land on arbitrary fractions; anything within half a percent counts as equal. */
const EPS = 0.005

export function clampZoom(z: number): number {
  return clampZoomTo(z, ZOOM_MIN, ZOOM_MAX)
}

export function nextPreset(z: number): number {
  return ZOOM_PRESETS.find((p) => p > z + EPS) ?? ZOOM_MAX
}

export function prevPreset(z: number): number {
  for (let i = ZOOM_PRESETS.length - 1; i >= 0; i--) {
    if (ZOOM_PRESETS[i] < z - EPS) return ZOOM_PRESETS[i]
  }
  return ZOOM_MIN
}

export function notchStep(z: number, dir: 1 | -1): number {
  return notchStepTo(z, dir, ZOOM_MIN, ZOOM_MAX)
}
