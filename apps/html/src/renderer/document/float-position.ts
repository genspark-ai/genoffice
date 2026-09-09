import type { ElementRect } from '../preview/inspector-protocol'

export interface FloatPosition {
  left: number
  top: number
  /** the bar sits under the element because there is no room above it */
  below: boolean
}

export interface FloatLayout {
  /** preview zoom in percent */
  zoom: number
  /** preview host origin relative to the stage the bar is positioned in */
  offsetX: number
  offsetY: number
  stageWidth: number
  barWidth: number
  barHeight: number
}

const GAP = 6
const EDGE = 4

/** floating toolbar anchored just above the selected element (frame coordinates → stage coordinates) */
export function floatPosition(rect: ElementRect, layout: FloatLayout): FloatPosition {
  const z = layout.zoom / 100
  const maxLeft = Math.max(EDGE, layout.stageWidth - layout.barWidth - EDGE)
  const left = Math.min(maxLeft, Math.max(EDGE, layout.offsetX + rect.x * z))
  const above = layout.offsetY + rect.y * z - layout.barHeight - GAP
  if (above >= EDGE) return { left, top: above, below: false }
  return { left, top: layout.offsetY + (rect.y + rect.height) * z + GAP, below: true }
}

/** `a: b; c: d` → declarations for a set_style op (invalid pieces dropped) */
export function parseDeclarations(css: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const part of css.split(';')) {
    const i = part.indexOf(':')
    if (i < 0) continue
    const prop = part.slice(0, i).trim().toLowerCase()
    const val = part.slice(i + 1).trim()
    if (/^-?[a-z][a-z0-9-]*$/.test(prop) && val) out[prop] = val
  }
  return out
}
