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

/** Diagonal / flip state of a straight connector with a real vertical extent. */
export interface LineRenderOpts {
  diag?: boolean
  flipH?: boolean
  flipV?: boolean
}

/**
 * Straight connector inside its box: level at the vertical center by default;
 * corner-to-corner when the shape has a real vertical extent (flips pick the
 * diagonal — flipV runs bottom-left → top-right). flipH also reverses level
 * lines so a head-only arrow (drawn at the path end) points the right way.
 */
function straightLinePaths(prst: string, w: number, h: number, line?: LineRenderOpts): ShapePaths {
  const x1 = line?.flipH ? w : 0
  const x2 = w - x1
  let y1 = h / 2
  let y2 = y1
  if (line?.diag) {
    y1 = line.flipV ? h : 0
    y2 = h - y1
  }
  const parts = [`M ${R(x1)} ${R(y1)} L ${R(x2)} ${R(y2)}`]
  const len = Math.min(16, Math.max(7, Math.hypot(x2 - x1, y2 - y1) * 0.09))
  if (prst !== 'line') parts.push(arrowHeadD(x1, y1, x2, y2, len))
  if (prst === 'lineArrowDouble') parts.push(arrowHeadD(x2, y2, x1, y1, len))
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
  line?: LineRenderOpts,
  strokeWidthPx?: number,
): string | null {
  const paths = isStraightLineKind(prst)
    ? straightLinePaths(prst, Math.max(8, w - 2), Math.max(8, h - 2), line)
    : shapePaths(prst, Math.max(8, w - 2), Math.max(8, h - 2))
  if (!paths) return null
  return pathsBackgroundImage(paths, w, h, fillHex, borderHex, strokeWidthPx)
}

function pathsBackgroundImage(
  paths: ShapePaths,
  w: number,
  h: number,
  fillHex?: string,
  borderHex?: string,
  strokeWidthPx?: number,
): string {
  const fill = fillHex ? `#${fillHex}` : 'none'
  const stroke = borderHex ? `#${borderHex}` : 'none'
  // authored outline width (a:ln w in px); absent keeps the historical 2 / 2.5
  const strokeWidth = strokeWidthPx === undefined ? undefined : R(Math.max(0.5, strokeWidthPx))
  const parts: string[] = []
  if (paths.main)
    parts.push(
      `<path d="${paths.main}" fill="${fill}" stroke="${stroke}" stroke-width="${strokeWidth ?? 2}"/>`,
    )
  if (paths.fillOnly) parts.push(`<path d="${paths.fillOnly}" fill="${fill}" stroke="none"/>`)
  if (paths.strokeOnly)
    parts.push(
      `<path d="${paths.strokeOnly}" fill="none" ` +
        `stroke="${stroke === 'none' ? fill : stroke}" stroke-width="${strokeWidth ?? 2.5}"/>`,
    )
  // a wide stroke needs more than the historical 1px inset, or the SVG viewport
  // clips it at the edges of the box: grow the viewBox symmetrically around it
  const inset = strokeWidth === undefined ? 1 : Math.max(1, R(strokeWidth / 2 + 0.5))
  const viewW = strokeWidth === undefined ? w : R(w + 2 * (inset - 1))
  const viewH = strokeWidth === undefined ? h : R(h + 2 * (inset - 1))
  // printToPDF honors the intrinsic ratio over background-size 100% 100%: a
  // box whose live height differs from the measured one shrank uniformly
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="${-inset} ${-inset} ${viewW} ${viewH}" preserveAspectRatio="none">` +
    parts.join('') +
    '</svg>'
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`
}

/** Normalized a:custGeom path channels (TextboxDisplay.pathData, coords 0..1). */
export interface CustGeomPathData {
  path?: string
  fillPath?: string
  strokePath?: string
}

/** Scale a normalized path string ("M x y L x y … Z") into a w×h pixel box. */
function scaleNormPathD(d: string, w: number, h: number): string {
  let axis = 0
  return d
    .split(' ')
    .map((tok) => {
      const n = Number(tok)
      if (!Number.isFinite(n)) {
        axis = 0
        return tok
      }
      return String(R(n * (axis++ % 2 === 0 ? w : h)))
    })
    .join(' ')
}

/** custGeom visual as full CSS background properties (same pipeline as presets). */
export function custGeomBackgroundCss(
  geom: CustGeomPathData,
  w: number,
  h: number,
  fillHex?: string,
  borderHex?: string,
  strokeWidthPx?: number,
): string {
  const sw = Math.max(8, w - 2)
  const sh = Math.max(8, h - 2)
  const paths: ShapePaths = {
    main: geom.path ? scaleNormPathD(geom.path, sw, sh) : undefined,
    fillOnly: geom.fillPath ? scaleNormPathD(geom.fillPath, sw, sh) : undefined,
    strokeOnly: geom.strokePath ? scaleNormPathD(geom.strokePath, sw, sh) : undefined,
  }
  const image = pathsBackgroundImage(paths, w, h, fillHex, borderHex, strokeWidthPx)
  return `background-image:${image};background-size:100% 100%;background-repeat:no-repeat`
}

/** Same visual as full CSS background properties (for style strings). */
export function shapeBackgroundCss(
  prst: string,
  w: number,
  h: number,
  fillHex?: string,
  borderHex?: string,
  line?: LineRenderOpts,
  strokeWidthPx?: number,
): string | null {
  const image = shapeBackgroundImage(prst, w, h, fillHex, borderHex, line, strokeWidthPx)
  if (!image) return null
  return `background-image:${image};background-size:100% 100%;background-repeat:no-repeat`
}

/** inset = cw·w + ch·h + cs·min(w,h) */
type InsetCoef = readonly [cw: number, ch: number, cs: number]
type RectCoef = readonly [l: InsetCoef, t: InsetCoef, r: InsetCoef, b: InsetCoef]

const Z: InsetCoef = [0, 0, 0]
const W = (c: number): InsetCoef => [c, 0, 0]
const H = (c: number): InsetCoef => [0, c, 0]
const S = (c: number): InsetCoef => [0, 0, c]
/** inscribed rect of the ellipse: (1 - 1/√2) / 2 per side */
const K = (1 - Math.SQRT1_2) / 2
const ELLIPSE: RectCoef = [W(K), H(K), W(K), H(K)]
const ROUND = S(0.04882)
const SNIP = S(1 / 12)

/**
 * OOXML presetShapeDefinitions text rectangles at the default adjust values,
 * as insets from the shape box. Every rect that is linear in w, h and
 * min(w,h) lives here; the few with a pin()/branch term are computed below.
 */
const TEXT_RECT_INSETS: Record<string, RectCoef> = {
  ellipse: ELLIPSE,
  chord: ELLIPSE,
  donut: ELLIPSE,
  noSmoking: ELLIPSE,
  smileyFace: ELLIPSE,
  teardrop: ELLIPSE,
  flowChartConnector: ELLIPSE,
  flowChartOr: ELLIPSE,
  flowChartSummingJunction: ELLIPSE,
  wedgeEllipseCallout: ELLIPSE,
  flowChartDelay: [Z, H(K), W(K), H(K)],
  flowChartTerminator: [W(0.04713), H(K), W(0.04713), H(K)],
  roundRect: [ROUND, ROUND, ROUND, ROUND],
  flowChartAlternateProcess: [ROUND, ROUND, ROUND, ROUND],
  wedgeRoundRectCallout: [ROUND, ROUND, ROUND, ROUND],
  round2DiagRect: [ROUND, ROUND, ROUND, ROUND],
  round2SameRect: [ROUND, ROUND, ROUND, Z],
  round1Rect: [Z, Z, ROUND, Z],
  snipRoundRect: [ROUND, ROUND, SNIP, Z],
  snip1Rect: [Z, SNIP, SNIP, Z],
  snip2SameRect: [SNIP, SNIP, SNIP, Z],
  snip2DiagRect: [SNIP, SNIP, SNIP, SNIP],
  horizontalScroll: [S(0.125), S(0.125), S(0.0625), S(0.125)],
  verticalScroll: [S(0.125), S(0.125), S(0.125), S(0.0625)],
  bevel: [S(0.125), S(0.125), S(0.125), S(0.125)],
  frame: [S(0.125), S(0.125), S(0.125), S(0.125)],
  plaque: [S(0.11785), S(0.11785), S(0.11785), S(0.11785)],
  octagon: [S(K), S(K), S(K), S(K)],
  can: [Z, S(0.25), Z, S(0.125)],
  cube: [Z, S(0.25), S(0.25), Z],
  foldedCorner: [Z, Z, Z, S(1 / 6)],
  diamond: [W(0.25), H(0.25), W(0.25), H(0.25)],
  flowChartDecision: [W(0.25), H(0.25), W(0.25), H(0.25)],
  flowChartSort: [W(0.25), H(0.25), W(0.25), H(0.25)],
  flowChartCollate: [W(0.25), H(0.25), W(0.25), H(0.25)],
  triangle: [W(0.25), H(0.5), W(0.25), Z],
  flowChartExtract: [W(0.25), H(0.5), W(0.25), Z],
  flowChartMerge: [W(0.25), Z, W(0.25), H(0.5)],
  rtTriangle: [W(1 / 12), H(7 / 12), W(5 / 12), H(1 / 12)],
  pentagon: [W(0.19098), H(0.23607), W(0.19098), Z],
  homePlate: [Z, Z, S(0.25), Z],
  heart: [W(1 / 6), H(0.25), W(1 / 6), H(1 / 3)],
  wave: [Z, H(0.25), Z, H(0.25)],
  doubleWave: [Z, H(0.125), Z, H(0.125)],
  ribbon: [W(0.25), H(1 / 6), W(0.25), Z],
  ribbon2: [W(0.25), Z, W(0.25), H(1 / 6)],
  diagStripe: [Z, Z, W(0.25), H(0.25)],
  rightArrow: [Z, H(0.25), S(0.25), H(0.25)],
  leftArrow: [S(0.25), H(0.25), Z, H(0.25)],
  leftRightArrow: [S(0.25), H(0.25), S(0.25), H(0.25)],
  notchedRightArrow: [S(0.25), H(0.25), S(0.25), H(0.25)],
  stripedRightArrow: [S(0.15625), H(0.25), S(0.25), H(0.25)],
  upArrow: [W(0.25), S(0.25), W(0.25), Z],
  downArrow: [W(0.25), Z, W(0.25), S(0.25)],
  upDownArrow: [W(0.25), S(0.25), W(0.25), S(0.25)],
  quadArrow: [S(0.1125), [0, 0.5, -0.1125], S(0.1125), [0, 0.5, -0.1125]],
  leftArrowCallout: [W(0.35023), Z, Z, Z],
  rightArrowCallout: [Z, Z, W(0.35023), Z],
  upArrowCallout: [Z, H(0.35023), Z, Z],
  downArrowCallout: [Z, Z, Z, H(0.35023)],
  star4: [W(0.41161), H(0.41161), W(0.41161), H(0.41161)],
  star5: [W(0.30902), H(0.38197), W(0.30902), H(0.23607)],
  star6: [W(1 / 6), H(0.25), W(1 / 6), H(0.25)],
  star7: [W(0.22252), H(0.19806), W(0.22252), H(0.24698)],
  star8: [W(0.15355), H(0.15355), W(0.15355), H(0.15355)],
  star10: [W(0.13819), H(0.25), W(0.13819), H(0.25)],
  star12: [W(0.23483), H(0.23483), W(0.23483), H(0.23483)],
  star16: [W(0.23483), H(0.23483), W(0.23483), H(0.23483)],
  star24: [W(0.23483), H(0.23483), W(0.23483), H(0.23483)],
  star32: [W(0.23483), H(0.23483), W(0.23483), H(0.23483)],
  sun: [W(0.32323), H(0.32323), W(0.32323), H(0.32323)],
  moon: [W(K), H(0.23949), W(0.5), H(0.23949)],
  cloud: [W(0.13782), H(0.15102), W(0.20894), H(0.19736)],
  cloudCallout: [W(0.13782), H(0.15102), W(0.20894), H(0.19736)],
  irregularSeal1: [W(0.21421), H(0.29259), W(0.22676), H(0.35477)],
  irregularSeal2: [W(0.2487), H(0.29546), W(0.32222), H(0.26227)],
  lightningBolt: [W(0.40542), H(0.34431), W(0.35569), H(0.33903)],
  gear6: [
    [0.14536, 0, 0.10639],
    [0, 0.14754, 0.10574],
    [0.14536, 0, 0.10639],
    [0, 0.14754, 0.10574],
  ],
  gear9: [
    [0.12631, 0, 0.07474],
    [0, 0.16781, 0.06644],
    [0.12631, 0, 0.07474],
    [0, 0.18967, 0.06207],
  ],
  circularArrow: [
    [K, 0, 0.04419],
    [0, K, 0.04419],
    [K, 0, 0.04419],
    [0, K, 0.04419],
  ],
  leftCircularArrow: [
    [K, 0, 0.04419],
    [0, K, 0.04419],
    [K, 0, 0.04419],
    [0, K, 0.04419],
  ],
  mathEqual: [W(0.13255), H(0.206), W(0.13255), H(0.206)],
  mathNotEqual: [W(0.13255), H(0.206), W(0.13255), H(0.206)],
  mathPlus: [W(0.13255), [0, 0.5, -0.1176], W(0.13255), [0, 0.5, -0.1176]],
  arc: [W(0.5), Z, Z, H(0.5)],
  blockArc: [Z, Z, Z, H(0.5)],
  // the spec's pie rect swaps t/r (known erratum): use the inscribed rect
  pie: ELLIPSE,
  leftBrace: [W(1 - 0.35355), S(0.02441), Z, S(0.02441)],
  rightBrace: [Z, S(0.02441), W(1 - 0.35355), S(0.02441)],
  leftBracket: [W(0.29289), S(0.02441), Z, S(0.02441)],
  rightBracket: [Z, S(0.02441), W(0.29289), S(0.02441)],
  flowChartDisplay: [W(1 / 6), Z, W(1 / 6), Z],
  flowChartMagneticDrum: [W(1 / 6), Z, W(1 / 3), Z],
  flowChartManualOperation: [W(0.2), Z, W(0.2), Z],
  flowChartPreparation: [W(0.2), Z, W(0.2), Z],
  flowChartPredefinedProcess: [W(0.125), Z, W(0.125), Z],
  flowChartInternalStorage: [W(0.125), H(0.125), Z, Z],
  flowChartManualInput: [Z, H(0.2), Z, Z],
  flowChartPunchedTape: [Z, H(0.2), Z, H(0.2)],
  flowChartMagneticDisk: [Z, H(1 / 3), Z, H(1 / 6)],
  flowChartOffpageConnector: [Z, Z, Z, H(0.2)],
  flowChartDocument: [Z, Z, Z, H(0.19806)],
  flowChartMultidocument: [Z, H(0.17014), W(0.13912), H(0.03787)],
}

/** text rects whose default-adjust formula pins on the aspect ratio */
function pinnedTextInsets(
  prst: string,
  w: number,
  h: number,
): { l: number; t: number; r: number; b: number } | null {
  const ss = Math.min(w, h)
  switch (prst) {
    case 'hexagon': {
      const maxAdj = (50000 * w) / ss
      const a = Math.min(25000, maxAdj)
      const q1 = -maxAdj / 2
      const q2 = a + q1
      const q6 = (a + (q2 > 0 ? q1 : 0)) / q1
      const q8 = (q2 > 0 ? 4 : 2) - q6 * (q2 > 0 ? 3 : 2)
      return { l: (w * q8) / 24, t: (h * q8) / 24, r: (w * q8) / 24, b: (h * q8) / 24 }
    }
    case 'parallelogram': {
      const maxAdj = (100000 * w) / ss
      const q2 = (1 + (5 * Math.min(25000, maxAdj)) / maxAdj) / 12
      return { l: q2 * w, t: q2 * h, r: q2 * w, b: q2 * h }
    }
    case 'trapezoid': {
      const maxAdj = (50000 * w) / ss
      const f = Math.min(25000, maxAdj) / maxAdj
      return { l: (w / 3) * f, t: (h / 3) * f, r: (w / 3) * f, b: 0 }
    }
    case 'plus': {
      const x1 = ss / 4
      return w > h ? { l: 0, t: x1, r: 0, b: x1 } : { l: x1, t: 0, r: x1, b: 0 }
    }
    case 'chevron': {
      const x1 = (ss * Math.min(50000, (100000 * w) / ss)) / 100000
      const inset = w - 2 * x1 > 0 ? x1 : 0
      return { l: inset, t: 0, r: inset, b: 0 }
    }
    case 'corner': {
      const x1 = (ss * Math.min(50000, (100000 * w) / ss)) / 100000
      const dy1 = (ss * Math.min(50000, (100000 * h) / ss)) / 100000
      return w > h ? { l: 0, t: h - dy1, r: 0, b: 0 } : { l: 0, t: 0, r: w - x1, b: 0 }
    }
    default:
      return null
  }
}

/** Geometry text-rect insets (px) inside a w×h preset shape; Word keeps the
 *  text inside this rect, bodyPr insets apply within it. Null = the full box.
 *  Document geometry, not chrome — never themed. */
export function shapeTextInsetsPx(
  prst: string,
  w: number,
  h: number,
): { l: number; t: number; r: number; b: number } | null {
  if (!(w > 0) || !(h > 0)) return null
  const coef = TEXT_RECT_INSETS[prst]
  if (!coef) return pinnedTextInsets(prst, w, h)
  const ss = Math.min(w, h)
  const at = ([cw, ch, cs]: InsetCoef) => Math.max(0, cw * w + ch * h + cs * ss)
  return { l: at(coef[0]), t: at(coef[1]), r: at(coef[2]), b: at(coef[3]) }
}
