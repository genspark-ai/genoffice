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

/** Frame-supplied numbers can be NaN/Infinity from a compromised preview frame. */
function finite(n: unknown, fallback: number): number {
  return typeof n === 'number' && Number.isFinite(n) ? n : fallback
}

/** floating toolbar anchored just above the selected element (frame coordinates → stage coordinates) */
export function floatPosition(rect: ElementRect, layout: FloatLayout): FloatPosition {
  const rx = finite(rect.x, 0)
  const ry = finite(rect.y, 0)
  const rh = finite(rect.height, 0)
  const zoom = finite(layout.zoom, 100)
  const offsetX = finite(layout.offsetX, 0)
  const offsetY = finite(layout.offsetY, 0)
  const stageWidth = finite(layout.stageWidth, 800)
  const barWidth = finite(layout.barWidth, 200)
  const barHeight = finite(layout.barHeight, 40)
  const z = zoom / 100
  const maxLeft = Math.max(EDGE, stageWidth - barWidth - EDGE)
  const left = Math.min(maxLeft, Math.max(EDGE, offsetX + rx * z))
  const above = offsetY + ry * z - barHeight - GAP
  if (above >= EDGE) return { left, top: above, below: false }
  return { left, top: offsetY + (ry + rh) * z + GAP, below: true }
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
