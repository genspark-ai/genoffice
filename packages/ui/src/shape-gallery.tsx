/**
 * Shared shape-gallery definitions for the office apps: one list of preset
 * shapes (grouped like PowerPoint's Shapes dropdown) and one geometry source
 * for previews and CSS clips, so every app's gallery shows the same shapes
 * with the same silhouettes. Label keys are the shared ribbonShape* i18n ids;
 * each app resolves them through its own strings table.
 */
// Deep import: preset-geometry is dependency-free, keeping the heavy render/engine
// sources out of every consumer app's type graph
import { isPillPreset, presetPath, presetPolygon } from '@genoffice/pptx-render/preset-geometry'

export interface ShapeGalleryShape {
  prst: string
  labelKey: string
}

export interface ShapeGalleryGroup {
  groupKey: string
  shapes: ShapeGalleryShape[]
}

const s = (prst: string, labelKey: string): ShapeGalleryShape => ({ prst, labelKey })

/** The canonical gallery (the full 104-preset cross-app set): groups → shapes, in display order. */
export const SHAPE_GALLERY_GROUPS: ShapeGalleryGroup[] = [
  {
    groupKey: 'ribbonShapeGroupLines',
    shapes: [
      s('line', 'ribbonShapeLine'),
      s('lineArrow', 'ribbonShapeLineArrow'),
      s('lineArrowDouble', 'ribbonShapeLineArrowDouble'),
      s('lineBent', 'ribbonShapeLineBent'),
      s('lineCurved', 'ribbonShapeLineCurved'),
    ],
  },
  {
    groupKey: 'ribbonShapeGroupRects',
    shapes: [
      s('rect', 'ribbonShapeRect'),
      s('roundRect', 'ribbonShapeRoundRect'),
      s('snip1Rect', 'ribbonShapeSnip1Rect'),
      s('snip2SameRect', 'ribbonShapeSnip2SameRect'),
      s('snip2DiagRect', 'ribbonShapeSnip2DiagRect'),
      s('snipRoundRect', 'ribbonShapeSnipRoundRect'),
      s('round1Rect', 'ribbonShapeRound1Rect'),
      s('round2SameRect', 'ribbonShapeRound2SameRect'),
      s('round2DiagRect', 'ribbonShapeRound2DiagRect'),
    ],
  },
  {
    groupKey: 'ribbonShapeGroupBasic',
    shapes: [
      s('ellipse', 'ribbonShapeEllipse'),
      s('triangle', 'ribbonShapeTriangle'),
      s('rtTriangle', 'ribbonShapeRtTriangle'),
      s('parallelogram', 'ribbonShapeParallelogram'),
      s('trapezoid', 'ribbonShapeTrapezoid'),
      s('diamond', 'ribbonShapeDiamond'),
      s('pentagon', 'ribbonShapePentagon'),
      s('hexagon', 'ribbonShapeHexagon'),
      s('octagon', 'ribbonShapeOctagon'),
      s('plus', 'ribbonShapePlus'),
      s('mathPlus', 'ribbonShapeMathPlus'),
      s('pie', 'ribbonShapePie'),
      s('chord', 'ribbonShapeChord'),
      s('teardrop', 'ribbonShapeTeardrop'),
      s('frame', 'ribbonShapeFrame'),
      s('halfFrame', 'ribbonShapeHalfFrame'),
      s('corner', 'ribbonShapeCorner'),
      s('diagStripe', 'ribbonShapeDiagStripe'),
      s('donut', 'ribbonShapeDonut'),
      s('noSmoking', 'ribbonShapeNoSmoking'),
      s('blockArc', 'ribbonShapeBlockArc'),
      s('foldedCorner', 'ribbonShapeFoldedCorner'),
      s('bevel', 'ribbonShapeBevel'),
      s('cube', 'ribbonShapeCube'),
      s('can', 'ribbonShapeCan'),
      s('lightningBolt', 'ribbonShapeLightningBolt'),
      s('heart', 'ribbonShapeHeart'),
      s('sun', 'ribbonShapeSun'),
      s('moon', 'ribbonShapeMoon'),
      s('cloud', 'ribbonShapeCloud'),
      s('arc', 'ribbonShapeArc'),
      s('plaque', 'ribbonShapePlaque'),
      s('smileyFace', 'ribbonShapeSmileyFace'),
      s('leftBracket', 'ribbonShapeLeftBracket'),
      s('rightBracket', 'ribbonShapeRightBracket'),
      s('leftBrace', 'ribbonShapeLeftBrace'),
      s('rightBrace', 'ribbonShapeRightBrace'),
    ],
  },
  {
    groupKey: 'ribbonShapeGroupArrows',
    shapes: [
      s('rightArrow', 'ribbonShapeRightArrow'),
      s('leftArrow', 'ribbonShapeLeftArrow'),
      s('upArrow', 'ribbonShapeUpArrow'),
      s('downArrow', 'ribbonShapeDownArrow'),
      s('leftRightArrow', 'ribbonShapeLeftRightArrow'),
      s('upDownArrow', 'ribbonShapeUpDownArrow'),
      s('quadArrow', 'ribbonShapeQuadArrow'),
      s('bentArrow', 'ribbonShapeBentArrow'),
      s('uturnArrow', 'ribbonShapeUturnArrow'),
      s('curvedRightArrow', 'ribbonShapeCurvedRightArrow'),
      s('stripedRightArrow', 'ribbonShapeStripedRightArrow'),
      s('notchedRightArrow', 'ribbonShapeNotchedRightArrow'),
      s('chevron', 'ribbonShapeChevron'),
      s('homePlate', 'ribbonShapeHomePlate'),
    ],
  },
  {
    groupKey: 'ribbonShapeGroupStars',
    shapes: [
      s('star4', 'ribbonShapeStar4'),
      s('star5', 'ribbonShapeStar5'),
      s('star6', 'ribbonShapeStar6'),
      s('star7', 'ribbonShapeStar7'),
      s('star8', 'ribbonShapeStar8'),
      s('star10', 'ribbonShapeStar10'),
      s('star12', 'ribbonShapeStar12'),
      s('star16', 'ribbonShapeStar16'),
      s('star24', 'ribbonShapeStar24'),
      s('star32', 'ribbonShapeStar32'),
      s('irregularSeal1', 'ribbonShapeIrregularSeal1'),
      s('irregularSeal2', 'ribbonShapeIrregularSeal2'),
      s('ribbon', 'ribbonShapeRibbon'),
      s('ribbon2', 'ribbonShapeRibbon2'),
      s('wave', 'ribbonShapeWave'),
      s('doubleWave', 'ribbonShapeDoubleWave'),
    ],
  },
  {
    groupKey: 'ribbonShapeGroupFlowchart',
    shapes: [
      s('flowChartProcess', 'ribbonShapeFlowProcess'),
      s('flowChartAlternateProcess', 'ribbonShapeFlowAlternateProcess'),
      s('flowChartDecision', 'ribbonShapeFlowDecision'),
      s('flowChartPredefinedProcess', 'ribbonShapeFlowPredefinedProcess'),
      s('flowChartInternalStorage', 'ribbonShapeFlowInternalStorage'),
      s('flowChartDocument', 'ribbonShapeFlowDocument'),
      s('flowChartMultidocument', 'ribbonShapeFlowMultidocument'),
      s('flowChartTerminator', 'ribbonShapeFlowTerminator'),
      s('flowChartPreparation', 'ribbonShapeFlowPreparation'),
      s('flowChartManualInput', 'ribbonShapeFlowManualInput'),
      s('flowChartManualOperation', 'ribbonShapeFlowManualOperation'),
      s('flowChartConnector', 'ribbonShapeFlowConnector'),
      s('flowChartOffpageConnector', 'ribbonShapeFlowOffpageConnector'),
      s('flowChartMagneticDisk', 'ribbonShapeFlowMagneticDisk'),
      s('flowChartMagneticDrum', 'ribbonShapeFlowMagneticDrum'),
      s('flowChartDisplay', 'ribbonShapeFlowDisplay'),
      s('flowChartDelay', 'ribbonShapeFlowDelay'),
      s('flowChartOr', 'ribbonShapeFlowOr'),
      s('flowChartCollate', 'ribbonShapeFlowCollate'),
      s('flowChartSort', 'ribbonShapeFlowSort'),
      s('flowChartExtract', 'ribbonShapeFlowExtract'),
      s('flowChartMerge', 'ribbonShapeFlowMerge'),
      s('flowChartSummingJunction', 'ribbonShapeFlowSummingJunction'),
      s('flowChartPunchedTape', 'ribbonShapeFlowPunchedTape'),
    ],
  },
  {
    groupKey: 'ribbonShapeGroupCallouts',
    shapes: [
      s('wedgeRectCallout', 'ribbonShapeWedgeRectCallout'),
      s('wedgeRoundRectCallout', 'ribbonShapeWedgeRoundRectCallout'),
      s('wedgeEllipseCallout', 'ribbonShapeWedgeEllipseCallout'),
      s('cloudCallout', 'ribbonShapeCloudCallout'),
    ],
  },
]

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

/**
 * Preset name → SVG path data (local w×h box), same resolution order as the
 * slides canvas renderer: polygon channel → path channel → pill → native
 * primitives. Returns null for presets none of the channels cover.
 */
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

/** Gallery cell: outline rendering of the preset geometry (matches the inserted shape). */
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

/**
 * CSS clip for rendering a filled prstGeom shape as a plain DOM box (docs
 * pages): polygon presets clip with size-independent percentages; path-channel
 * presets need the box size for an absolute path() clip. roundRect/ellipse
 * stay border-radius so CSS borders keep working. Returns undefined when the
 * geometry channels don't cover the preset (callers fall back to a plain box).
 */
export function shapeClipCss(
  prst: string,
  wPx?: number,
  hPx?: number,
): { clipPath?: string; borderRadius?: string } | undefined {
  if (prst === 'rect') return undefined
  if (prst === 'roundRect') return { borderRadius: '12%' }
  if (prst === 'ellipse') return { borderRadius: '50%' }
  if (isPillPreset(prst)) return { borderRadius: '9999px' }
  const poly = presetPolygon(prst, 100, 100)
  if (poly) {
    const pts: string[] = []
    for (let i = 0; i < poly.length; i += 2) pts.push(`${R(poly[i]!)}% ${R(poly[i + 1]!)}%`)
    return { clipPath: `polygon(${pts.join(', ')})` }
  }
  const w = wPx ?? 100
  const h = hPx ?? 100
  const path = presetPath(prst, w, h)
  if (path) {
    const d = path.fillPath ?? path.path
    if (d) return { clipPath: `path('${d}')` }
  }
  return undefined
}
