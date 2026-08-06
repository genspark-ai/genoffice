/**
 * Insert-tab shape gallery: grouped presets + geometry-driven previews.
 * Preview cells trace the same preset geometry as the canvas renderer and the
 * saved OOXML prstGeom, so the picked cell matches the inserted shape.
 */
import { useEffect, useRef, useState } from 'react'

import {
  isPillPreset,
  presetPath,
  presetPolygon,
} from '../../../../packages/pptx-render/src/preset-geometry'
import { t, type StringKey } from './i18n/locale'

export interface ShapeDef {
  /** OOXML prstGeom name (persisted verbatim, also drives the preview) */
  prst: string
  label: string
}

const shape = (prst: string, key: StringKey): ShapeDef => ({
  prst,
  get label() {
    return t(key)
  },
})

const shapeGroup = (key: StringKey, shapes: ShapeDef[]): { group: string; shapes: ShapeDef[] } => ({
  get group() {
    return t(key)
  },
  shapes,
})

export const SHAPE_GALLERY: Array<{ group: string; shapes: ShapeDef[] }> = [
  shapeGroup('appShapeGroupRects', [
    shape('rect', 'appShapeRect'),
    shape('roundRect', 'appShapeRoundRect'),
    shape('snip1Rect', 'appShapeSnip1Rect'),
    shape('snip2SameRect', 'appShapeSnip2SameRect'),
    shape('snip2DiagRect', 'appShapeSnip2DiagRect'),
    shape('snipRoundRect', 'appShapeSnipRoundRect'),
    shape('round1Rect', 'appShapeRound1Rect'),
    shape('round2SameRect', 'appShapeRound2SameRect'),
    shape('round2DiagRect', 'appShapeRound2DiagRect'),
  ]),
  shapeGroup('appShapeGroupBasic', [
    shape('ellipse', 'appShapeEllipse'),
    shape('triangle', 'appShapeTriangle'),
    shape('rtTriangle', 'appShapeRtTriangle'),
    shape('parallelogram', 'appShapeParallelogram'),
    shape('trapezoid', 'appShapeTrapezoid'),
    shape('diamond', 'appShapeDiamond'),
    shape('pentagon', 'appShapePentagon'),
    shape('hexagon', 'appShapeHexagon'),
    shape('octagon', 'appShapeOctagon'),
    shape('plus', 'appShapePlus'),
    shape('mathPlus', 'appShapeMathPlus'),
    shape('pie', 'appShapePie'),
    shape('chord', 'appShapeChord'),
    shape('teardrop', 'appShapeTeardrop'),
    shape('frame', 'appShapeFrame'),
    shape('halfFrame', 'appShapeHalfFrame'),
    shape('corner', 'appShapeCorner'),
    shape('diagStripe', 'appShapeDiagStripe'),
    shape('donut', 'appShapeDonut'),
    shape('noSmoking', 'appShapeNoSmoking'),
    shape('blockArc', 'appShapeBlockArc'),
    shape('foldedCorner', 'appShapeFoldedCorner'),
    shape('bevel', 'appShapeBevel'),
    shape('cube', 'appShapeCube'),
    shape('can', 'appShapeCan'),
    shape('lightningBolt', 'appShapeLightningBolt'),
    shape('heart', 'appShapeHeart'),
    shape('sun', 'appShapeSun'),
    shape('moon', 'appShapeMoon'),
    shape('cloud', 'appShapeCloud'),
    shape('arc', 'appShapeArc'),
    shape('plaque', 'appShapePlaque'),
    shape('smileyFace', 'appShapeSmileyFace'),
    shape('leftBracket', 'appShapeLeftBracket'),
    shape('rightBracket', 'appShapeRightBracket'),
    shape('leftBrace', 'appShapeLeftBrace'),
    shape('rightBrace', 'appShapeRightBrace'),
  ]),
  shapeGroup('appShapeGroupArrows', [
    shape('rightArrow', 'appShapeRightArrow'),
    shape('leftArrow', 'appShapeLeftArrow'),
    shape('upArrow', 'appShapeUpArrow'),
    shape('downArrow', 'appShapeDownArrow'),
    shape('leftRightArrow', 'appShapeLeftRightArrow'),
    shape('upDownArrow', 'appShapeUpDownArrow'),
    shape('quadArrow', 'appShapeQuadArrow'),
    shape('bentArrow', 'appShapeBentArrow'),
    shape('uturnArrow', 'appShapeUturnArrow'),
    shape('curvedRightArrow', 'appShapeCurvedRightArrow'),
    shape('stripedRightArrow', 'appShapeStripedRightArrow'),
    shape('notchedRightArrow', 'appShapeNotchedRightArrow'),
    shape('chevron', 'appShapeChevron'),
    shape('homePlate', 'appShapeHomePlate'),
  ]),
  shapeGroup('appShapeGroupStars', [
    shape('star4', 'appShapeStar4'),
    shape('star5', 'appShapeStar5'),
    shape('star6', 'appShapeStar6'),
    shape('star7', 'appShapeStar7'),
    shape('star8', 'appShapeStar8'),
    shape('star10', 'appShapeStar10'),
    shape('star12', 'appShapeStar12'),
    shape('star16', 'appShapeStar16'),
    shape('star24', 'appShapeStar24'),
    shape('star32', 'appShapeStar32'),
    shape('irregularSeal1', 'appShapeIrregularSeal1'),
    shape('irregularSeal2', 'appShapeIrregularSeal2'),
    shape('ribbon', 'appShapeRibbon'),
    shape('ribbon2', 'appShapeRibbon2'),
    shape('wave', 'appShapeWave'),
    shape('doubleWave', 'appShapeDoubleWave'),
  ]),
  shapeGroup('appShapeGroupFlowchart', [
    shape('flowChartProcess', 'appShapeFlowProcess'),
    shape('flowChartAlternateProcess', 'appShapeFlowAlternateProcess'),
    shape('flowChartDecision', 'appShapeFlowDecision'),
    shape('flowChartPredefinedProcess', 'appShapeFlowPredefinedProcess'),
    shape('flowChartInternalStorage', 'appShapeFlowInternalStorage'),
    shape('flowChartDocument', 'appShapeFlowDocument'),
    shape('flowChartMultidocument', 'appShapeFlowMultidocument'),
    shape('flowChartTerminator', 'appShapeFlowTerminator'),
    shape('flowChartPreparation', 'appShapeFlowPreparation'),
    shape('flowChartManualInput', 'appShapeFlowManualInput'),
    shape('flowChartManualOperation', 'appShapeFlowManualOperation'),
    shape('flowChartConnector', 'appShapeFlowConnector'),
    shape('flowChartOffpageConnector', 'appShapeFlowOffpageConnector'),
    shape('flowChartMagneticDisk', 'appShapeFlowMagneticDisk'),
    shape('flowChartMagneticDrum', 'appShapeFlowMagneticDrum'),
    shape('flowChartDisplay', 'appShapeFlowDisplay'),
    shape('flowChartDelay', 'appShapeFlowDelay'),
    shape('flowChartOr', 'appShapeFlowOr'),
    shape('flowChartCollate', 'appShapeFlowCollate'),
    shape('flowChartSort', 'appShapeFlowSort'),
    shape('flowChartExtract', 'appShapeFlowExtract'),
    shape('flowChartMerge', 'appShapeFlowMerge'),
    shape('flowChartSummingJunction', 'appShapeFlowSummingJunction'),
    shape('flowChartPunchedTape', 'appShapeFlowPunchedTape'),
  ]),
  shapeGroup('appShapeGroupCallouts', [
    shape('wedgeRectCallout', 'appShapeWedgeRectCallout'),
    shape('wedgeRoundRectCallout', 'appShapeWedgeRoundRectCallout'),
    shape('wedgeEllipseCallout', 'appShapeWedgeEllipseCallout'),
    shape('cloudCallout', 'appShapeCloudCallout'),
  ]),
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

/** Preset name → stroke-only preview path (same channel order as the canvas). */
function shapePreviewPath(prst: string, w: number, h: number): string | null {
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

/** Flowchart nodes preview flat (also disambiguates them from diamond/ellipse). */
function previewBox(prst: string, size: number): { w: number; h: number } {
  return prst.startsWith('flowChart') ? { w: size, h: size * 0.62 } : { w: size, h: size }
}

export function ShapePreview({ prst, size = 20 }: { prst: string; size?: number }) {
  const { w, h } = previewBox(prst, size)
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

/** Cover-trigger dropdown with the grouped shape grid (Insert tab). */
export function ShapeGalleryMenu({
  label,
  onPick,
}: {
  readonly label: string
  readonly onPick: (prst: string) => void
}): React.JSX.Element {
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    if (!open) return
    const onDown = (event: MouseEvent): void => {
      if (!wrapRef.current?.contains(event.target as Node)) setOpen(false)
    }
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') setOpen(false)
    }
    window.addEventListener('mousedown', onDown)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('mousedown', onDown)
      window.removeEventListener('keydown', onKey)
    }
  }, [open])
  return (
    <div ref={wrapRef} className="menu-select menu-select-cover">
      <button
        type="button"
        className="cover-select"
        aria-label={label}
        aria-haspopup="listbox"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
      />
      {open && (
        <div className="menu-select-drop shape-gallery-drop" role="listbox" aria-label={label}>
          {SHAPE_GALLERY.map((group) => (
            <div key={group.group} className="shape-gallery-group">
              <div className="shape-gallery-title">{group.group}</div>
              <div className="shape-gallery-grid">
                {group.shapes.map((s) => (
                  <button
                    type="button"
                    key={s.prst}
                    role="option"
                    title={s.label}
                    aria-label={s.label}
                    onClick={() => {
                      setOpen(false)
                      onPick(s.prst)
                    }}
                  >
                    <ShapePreview prst={s.prst} />
                  </button>
                ))}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}
