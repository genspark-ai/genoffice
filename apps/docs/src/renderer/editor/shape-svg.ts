/**
 * Preset shape geometry for docs: gallery previews and textbox shape
 * backgrounds share the geometry the saved OOXML prstGeom describes.
 * The shape visual is a data-URI SVG background so the box DOM stays a plain
 * div (textbox sub-editors replaceChildren() the box on mount).
 */
import {
  isPillPreset,
  presetPath,
  presetPolygon,
} from '../../../../../packages/pptx-render/src/preset-geometry'

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

interface ShapePaths {
  /** fill + stroke */
  main?: string
  /** fill only */
  fillOnly?: string
  /** stroke only (braces, arcs) */
  strokeOnly?: string
}

/** Open V arrowhead at (x2,y2), pointing away from (x1,y1). */
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

/** Straight line kinds: saved with cy=0 (Word's horizontal line), drawn level. */
export function isStraightLineKind(prst: string | undefined): boolean {
  return prst === 'line' || prst === 'lineArrow' || prst === 'lineArrowDouble'
}

/** Level line at the box's vertical center (the in-document rendering). */
function straightLinePaths(prst: string, w: number, h: number): ShapePaths {
  const y = h / 2
  const parts = [`M 0 ${R(y)} L ${R(w)} ${R(y)}`]
  const len = Math.min(16, Math.max(7, w * 0.09))
  if (prst !== 'line') parts.push(arrowHeadD(0, y, w, y, len))
  if (prst === 'lineArrowDouble') parts.push(arrowHeadD(w, y, 0, y, len))
  return { strokeOnly: parts.join(' ') }
}

/** Stroke-only line/connector kinds (gallery preview traces the diagonal). */
function linePaths(prst: string, w: number, h: number): ShapePaths | null {
  if (isStraightLineKind(prst)) {
    const parts = [`M 0 0 L ${R(w)} ${R(h)}`]
    const len = Math.min(16, Math.max(7, Math.hypot(w, h) * 0.09))
    if (prst !== 'line') parts.push(arrowHeadD(0, 0, w, h, len))
    if (prst === 'lineArrowDouble') parts.push(arrowHeadD(w, h, 0, 0, len))
    return { strokeOnly: parts.join(' ') }
  }
  if (prst === 'lineBent')
    return { strokeOnly: `M 0 0 L ${R(w / 2)} 0 L ${R(w / 2)} ${R(h)} L ${R(w)} ${R(h)}` }
  if (prst === 'lineCurved')
    return { strokeOnly: `M 0 0 C ${R(w * 0.6)} 0 ${R(w * 0.4)} ${R(h)} ${R(w)} ${R(h)}` }
  return null
}

/** Preset → path channels in a w×h box; null when the geometry is unknown. */
export function shapePaths(prst: string, w: number, h: number): ShapePaths | null {
  const line = linePaths(prst, w, h)
  if (line) return line
  const poly = presetPolygon(prst, w, h)
  if (poly) return { main: polygonPathD(poly) }
  const path = presetPath(prst, w, h)
  if (path) return { main: path.path, fillOnly: path.fillPath, strokeOnly: path.strokePath }
  // Alternate Process is a rounded rectangle, not a pill like Terminator
  if (prst === 'roundRect' || prst === 'flowChartAlternateProcess')
    return { main: roundRectPathD(w, h, Math.min(w, h) * 0.16667) }
  if (isPillPreset(prst)) return { main: roundRectPathD(w, h, Math.min(w, h) / 2) }
  if (prst === 'ellipse') {
    const rx = w / 2
    const ry = h / 2
    return {
      main: `M 0 ${R(ry)} A ${R(rx)} ${R(ry)} 0 1 1 ${R(w)} ${R(ry)} A ${R(rx)} ${R(ry)} 0 1 1 0 ${R(ry)} Z`,
    }
  }
  if (prst === 'rect' || prst === 'flowChartProcess')
    return { main: `M 0 0 L ${R(w)} 0 L ${R(w)} ${R(h)} L 0 ${R(h)} Z` }
  return null
}

/** Stroke-only outline for gallery preview cells. */
export function shapePreviewPathD(prst: string, w: number, h: number): string | null {
  const paths = shapePaths(prst, w, h)
  if (!paths) return null
  return [paths.main, paths.fillOnly, paths.strokeOnly].filter(Boolean).join(' ')
}

/**
 * Shape visual as a CSS background-image url() (data-URI SVG at the box's
 * pixel size, insets included so the stroke isn't clipped at the edges).
 */
export function shapeBackgroundImage(
  prst: string,
  w: number,
  h: number,
  fillHex?: string,
  borderHex?: string,
): string | null {
  const paths = isStraightLineKind(prst)
    ? straightLinePaths(prst, Math.max(8, w - 2), Math.max(8, h - 2))
    : shapePaths(prst, Math.max(8, w - 2), Math.max(8, h - 2))
  if (!paths) return null
  const fill = fillHex ? `#${fillHex}` : 'none'
  const stroke = borderHex ? `#${borderHex}` : 'none'
  const parts: string[] = []
  if (paths.main)
    parts.push(`<path d="${paths.main}" fill="${fill}" stroke="${stroke}" stroke-width="2"/>`)
  if (paths.fillOnly) parts.push(`<path d="${paths.fillOnly}" fill="${fill}" stroke="none"/>`)
  if (paths.strokeOnly)
    parts.push(
      `<path d="${paths.strokeOnly}" fill="none" ` +
        `stroke="${stroke === 'none' ? fill : stroke}" stroke-width="2.5"/>`,
    )
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="-1 -1 ${w} ${h}">` +
    parts.join('') +
    '</svg>'
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`
}

/** Same visual as full CSS background properties (for style strings). */
export function shapeBackgroundCss(
  prst: string,
  w: number,
  h: number,
  fillHex?: string,
  borderHex?: string,
): string | null {
  const image = shapeBackgroundImage(prst, w, h, fillHex, borderHex)
  if (!image) return null
  return `background-image:${image};background-size:100% 100%;background-repeat:no-repeat`
}
