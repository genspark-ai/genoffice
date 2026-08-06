/**
 * Insert-tab shape gallery: grouped presets + geometry-driven previews.
 * The preview traces the same geometry as the box background and the saved
 * OOXML prstGeom, so the picked cell matches the inserted shape.
 */
import { shapePreviewPathD } from '../editor/shape-svg'
import type { StringKey } from '../i18n/locale'

export interface ShapeDef {
  /** OOXML prstGeom name (persisted verbatim, also drives the preview) */
  prst: string
  labelKey: StringKey
}

export interface ShapeGroup {
  labelKey: StringKey
  shapes: ShapeDef[]
}

const s = (prst: string, labelKey: StringKey): ShapeDef => ({ prst, labelKey })

/** Display kinds inserted as stroke-only lines/connectors (no text body). */
export function isLineKind(prst: string): boolean {
  return prst.startsWith('line')
}

export const SHAPE_GALLERY: ShapeGroup[] = [
  {
    labelKey: 'ribbonShapeGroupLines',
    shapes: [
      s('line', 'ribbonShapeLine'),
      s('lineArrow', 'ribbonShapeLineArrow'),
      s('lineArrowDouble', 'ribbonShapeLineArrowDouble'),
      s('lineBent', 'ribbonShapeLineBent'),
      s('lineCurved', 'ribbonShapeLineCurved'),
    ],
  },
  {
    labelKey: 'ribbonShapeGroupRects',
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
    labelKey: 'ribbonShapeGroupBasic',
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
    labelKey: 'ribbonShapeGroupArrows',
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
    labelKey: 'ribbonShapeGroupStars',
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
    labelKey: 'ribbonShapeGroupFlowchart',
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
    labelKey: 'ribbonShapeGroupCallouts',
    shapes: [
      s('wedgeRectCallout', 'ribbonShapeWedgeRectCallout'),
      s('wedgeRoundRectCallout', 'ribbonShapeWedgeRoundRectCallout'),
      s('wedgeEllipseCallout', 'ribbonShapeWedgeEllipseCallout'),
      s('cloudCallout', 'ribbonShapeCloudCallout'),
    ],
  },
]

const SHAPE_BY_PRST = new Map(SHAPE_GALLERY.flatMap((g) => g.shapes.map((d) => [d.prst, d])))

export function shapeLabelKey(prst: string): StringKey {
  return SHAPE_BY_PRST.get(prst)?.labelKey ?? 'ribbonShapeRect'
}

/** Flowchart nodes preview flat (also disambiguates them from diamond/ellipse). */
export function ShapePreview({ prst, size = 20 }: { prst: string; size?: number }) {
  const h = prst.startsWith('flowChart') ? size * 0.62 : size
  const d = shapePreviewPathD(prst, size, h)
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
