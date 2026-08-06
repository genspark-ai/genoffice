/**
 * Geometry-driven gallery previews: shape cells render the same preset geometry
 * as the inserted OOXML shape, SmartArt thumbs reuse the engine's layout math,
 * so preview and insert result always match.
 */
import { isPillPreset, presetPath, presetPolygon } from '@genoffice/pptx-render'
import {
  layoutShapes,
  type SmartArtLayout,
} from '../../../../../packages/pptx-engine/src/smartart-layout'

const R = (v: number) => Math.round(v * 100) / 100

function polygonPathD(pts: number[]): string {
  const parts: string[] = []
  for (let i = 0; i < pts.length; i += 2) {
    parts.push(`${i === 0 ? 'M' : 'L'} ${R(pts[i]!)} ${R(pts[i + 1]!)}`)
  }
  return parts.join(' ') + ' Z'
}

function roundRectPathD(w: number, h: number, r: number): string {
  return (
    `M ${R(r)} 0 L ${R(w - r)} 0 A ${R(r)} ${R(r)} 0 0 1 ${R(w)} ${R(r)} L ${R(w)} ${R(h - r)} ` +
    `A ${R(r)} ${R(r)} 0 0 1 ${R(w - r)} ${R(h)} L ${R(r)} ${R(h)} A ${R(r)} ${R(r)} 0 0 1 0 ${R(h - r)} ` +
    `L 0 ${R(r)} A ${R(r)} ${R(r)} 0 0 1 ${R(r)} 0 Z`
  )
}

/**
 * Preset name → SVG path data (local w×h box), same resolution order as the
 * canvas renderer: polygon channel → path channel → pill → native primitives.
 * Returns null for presets none of the channels cover.
 */
/** Open V arrowhead stroke at (x2,y2), pointing away from (x1,y1). */
function arrowHeadD(x1: number, y1: number, x2: number, y2: number, len: number): string {
  const dx = x2 - x1
  const dy = y2 - y1
  const l = Math.hypot(dx, dy) || 1
  const ux = dx / l
  const uy = dy / l
  const bx = x2 - ux * len
  const by = y2 - uy * len
  const wl = len * 0.6
  return `M ${R(bx - uy * wl)} ${R(by + ux * wl)} L ${R(x2)} ${R(y2)} L ${R(bx + uy * wl)} ${R(by - ux * wl)}`
}

export function shapePreviewPath(prst: string, w: number, h: number): string | null {
  if (prst === 'line' || prst === 'lineArrow' || prst === 'lineArrowDouble') {
    const parts = [`M 0 ${R(h)} L ${R(w)} 0`]
    const len = Math.min(w, h) * 0.35
    if (prst !== 'line') parts.push(arrowHeadD(0, h, w, 0, len))
    if (prst === 'lineArrowDouble') parts.push(arrowHeadD(w, 0, 0, h, len))
    return parts.join(' ')
  }
  if (prst === 'lineBent') {
    // Elbow connector preview: bentConnector3 default shape (adj1 = 50%)
    return `M 0 ${R(h)} L ${R(w / 2)} ${R(h)} L ${R(w / 2)} 0 L ${R(w)} 0`
  }
  if (prst === 'lineCurved') {
    // Curved connector preview: S curve between opposite corners
    return `M 0 ${R(h)} C ${R(w * 0.6)} ${R(h)} ${R(w * 0.4)} 0 ${R(w)} 0`
  }
  const poly = presetPolygon(prst, w, h)
  if (poly) return polygonPathD(poly)
  const path = presetPath(prst, w, h)
  if (path) return [path.path, path.fillPath, path.strokePath].filter(Boolean).join(' ')
  // Alternate Process is a rounded rectangle, not a pill like Terminator
  if (prst === 'roundRect' || prst === 'flowChartAlternateProcess')
    return roundRectPathD(w, h, Math.min(w, h) * 0.16667)
  if (isPillPreset(prst)) return roundRectPathD(w, h, Math.min(w, h) / 2)
  if (prst === 'ellipse') {
    const rx = w / 2
    const ry = h / 2
    return `M 0 ${R(ry)} A ${R(rx)} ${R(ry)} 0 1 1 ${R(w)} ${R(ry)} A ${R(rx)} ${R(ry)} 0 1 1 0 ${R(ry)} Z`
  }
  if (prst === 'rect' || prst === 'flowChartProcess')
    return `M 0 0 L ${R(w)} 0 L ${R(w)} ${R(h)} L 0 ${R(h)} Z`
  return null
}

/** Preview box for a prst: flowchart nodes are flat (also disambiguates them from diamond/ellipse). */
export function shapePreviewBox(prst: string, size: number): { w: number; h: number } {
  return prst.startsWith('flowChart') ? { w: size, h: size * 0.62 } : { w: size, h: size }
}

export function ShapePreview({ prst, size = 18 }: { prst: string; size?: number }) {
  const { w, h } = shapePreviewBox(prst, size)
  const d = shapePreviewPath(prst, w, h)
  return (
    <svg
      width={size}
      height={size}
      viewBox={`-1 ${-1 - (size - h) / 2} ${size + 2} ${size + 2}`}
      aria-hidden
    >
      {d ? (
        <path d={d} fill="none" stroke="currentColor" strokeWidth={1.2} strokeLinejoin="round" />
      ) : null}
    </svg>
  )
}

/** Same virtual canvas scale as real insertion (EMU) so layout ratios/minimums behave identically. */
const SA_CX = 4800000
const SA_CY = 3000000
/** Node counts mirror SMARTART_GALLERY defaultItems */
const SA_ITEMS: Record<SmartArtLayout, number> = {
  list: 3,
  process: 4,
  cycle: 4,
  hierarchy: 4,
  pyramid: 3,
  matrix: 4,
  venn: 3,
}

export function smartArtPreviewShapes(
  layout: SmartArtLayout,
): Array<{ d: string; color: string; x: number; y: number; alpha?: number }> {
  const items = Array.from({ length: SA_ITEMS[layout] }, () => '')
  return layoutShapes(layout, items, SA_CX, SA_CY).map((s) => ({
    d: shapePreviewPath(s.prst, s.box.cx, s.box.cy) ?? '',
    color: `#${s.color}`,
    x: s.box.x,
    y: s.box.y,
    alpha: s.alpha,
  }))
}

export function SmartArtPreview({
  layout,
  width = 44,
}: {
  layout: SmartArtLayout
  width?: number
}) {
  return (
    <svg
      width={width}
      height={(width * SA_CY) / SA_CX}
      viewBox={`0 0 ${SA_CX} ${SA_CY}`}
      aria-hidden
    >
      {smartArtPreviewShapes(layout).map((s, i) => (
        <path
          key={i}
          d={s.d}
          transform={`translate(${s.x} ${s.y})`}
          fill={s.color}
          fillOpacity={s.alpha}
        />
      ))}
    </svg>
  )
}
