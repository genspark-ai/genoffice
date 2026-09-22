/**
 * Vector export backend: a RenderSlide → one SVG document per page.
 *
 * Mirrors NodeBody/SlideThumb (the Konva backend) element by element so the
 * PDF export shares the render tree's layout, but keeps text as real glyph
 * runs: printToPDF turns `<text>` into PDF text operators with embedded font
 * subsets, so the exported PDF is searchable, copyable and convertible instead
 * of a page-sized bitmap. Effects SVG cannot express (WordArt warp) fall back
 * to a per-node raster supplied by the caller.
 */
import type {
  ArrowEndRender,
  CellBevelRender,
  ChartRenderNode,
  ChipRenderNode,
  GroupRenderNode,
  PictureRenderNode,
  RenderFill,
  RenderGlow,
  RenderNode,
  RenderReflection,
  RenderShadow,
  RenderSlide,
  RenderStroke,
  RenderTextLayout,
  ShapeRenderNode,
  TableCellRender,
  TableRenderNode,
} from '@genoffice/pptx-render'
import { extrusionFrontFace } from '@genoffice/pptx-render'
import {
  anchoredTileCanvas,
  averageColor,
  cropToKonva,
  featheredImage,
  featheredShapeCanvas,
  flatColorImage,
  glyphToDraw,
  hasBlipEffects,
  insetFillTile,
  isDegenerateImage,
  isOverlayShadow,
  konvaBaselineDrop,
  linearGradientDirection,
  linearRampStops,
  normalizeColor,
  pathGradientCanvas,
  patternCanvas,
  presetToShapeKind,
  processedImage,
  processedImageKey,
  radialCircleGeometry,
  shapeShadowOverlay,
  smoothTension,
  type GlyphDraw,
  type ShadowGeom,
  displayFontFamily,
} from './konva-adapter'

/** Raster of one top-level node in slide px (data URL), for effects SVG cannot express. */
export interface NodeRaster {
  href: string
  x: number
  y: number
  w: number
  h: number
}

export interface SvgExportOptions {
  /** Rasterize a node the vector backend cannot draw; null keeps it out of the page. */
  rasterizeNode?: (node: RenderNode) => Promise<NodeRaster | null>
  /** Def id prefix: every page shares one print document, so ids must not repeat across pages */
  idPrefix?: string
}

type Origin = { x: number; y: number }
type Bounds = { x: number; y: number; w: number; h: number }
type ImageSource = HTMLImageElement | HTMLCanvasElement

const XLINK = 'http://www.w3.org/1999/xlink'
const SVG_NS = 'http://www.w3.org/2000/svg'
const CHART_FONT = 'Calibri, Carlito, Arial, sans-serif'

/** 3 decimals: enough for print, keeps the markup small. */
function n(v: number): string {
  if (!Number.isFinite(v)) return '0'
  const r = Math.round(v * 1000) / 1000
  return r === 0 ? '0' : String(r)
}

function esc(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

// XML 1.0 rejects most C0 controls; tabs/newlines are kept (they are glyph-less anyway)
// eslint-disable-next-line no-control-regex
const XML_INVALID = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g

function escText(s: string): string {
  return esc(s.replace(XML_INVALID, ''))
}

function attrs(a: Record<string, string | number | undefined>): string {
  let out = ''
  for (const [k, v] of Object.entries(a)) {
    if (v === undefined || v === '') continue
    out += ` ${k}="${typeof v === 'number' ? n(v) : esc(v)}"`
  }
  return out
}

function el(tag: string, a: Record<string, string | number | undefined>, inner = ''): string {
  return inner ? `<${tag}${attrs(a)}>${inner}</${tag}>` : `<${tag}${attrs(a)}/>`
}

/** rgba(r,g,b,a) → separate color/opacity (SVG stop-color predates alpha colors in some PDF backends). */
function splitAlpha(color: string): { color: string; opacity?: number } {
  const c = normalizeColor(color)
  const m = /^rgba\((\d+),(\d+),(\d+),([0-9.]+)\)$/.exec(c)
  if (c === 'transparent') return { color: '#000000', opacity: 0 }
  if (!m) return { color: c }
  return { color: `rgb(${m[1]},${m[2]},${m[3]})`, opacity: Number(m[4]) }
}

const hrefCache = new WeakMap<object, string>()

/** data: URL for an image source; non-data sources (blob:, protocol URLs) are re-encoded. */
function imageHref(src: ImageSource): string | null {
  const hit = hrefCache.get(src)
  if (hit) return hit
  let out: string | null = null
  if (src instanceof HTMLCanvasElement) {
    if (src.width && src.height) out = src.toDataURL('image/png')
  } else {
    const s = src.currentSrc || src.src
    if (s.startsWith('data:')) out = s
    else if (src.naturalWidth && src.naturalHeight) {
      const cv = document.createElement('canvas')
      cv.width = src.naturalWidth
      cv.height = src.naturalHeight
      cv.getContext('2d')?.drawImage(src, 0, 0)
      out = cv.toDataURL('image/png')
    }
  }
  if (out) hrefCache.set(src, out)
  return out
}

function srcSize(src: ImageSource): { w: number; h: number } {
  return src instanceof HTMLCanvasElement
    ? { w: src.width, h: src.height }
    : { w: src.naturalWidth || src.width, h: src.naturalHeight || src.height }
}

/** Konva rotate/flip container: pivot on the box center, matching boxPivotProps. */
function boxTransform(box: RenderNode['box']): string {
  const sx = box.flipH ? -1 : 1
  const sy = box.flipV ? -1 : 1
  if (!box.rotationDeg && sx === 1 && sy === 1) return `translate(${n(box.x)} ${n(box.y)})`
  return (
    `translate(${n(box.x + box.w / 2)} ${n(box.y + box.h / 2)})` +
    (box.rotationDeg ? ` rotate(${n(box.rotationDeg)})` : '') +
    (sx !== 1 || sy !== 1 ? ` scale(${sx} ${sy})` : '') +
    ` translate(${n(-box.w / 2)} ${n(-box.h / 2)})`
  )
}

/** Axis-aligned bounds of a rotated/flipped box in its parent's coordinates. */
export function rotatedBounds(box: RenderNode['box']): Bounds {
  if (!box.rotationDeg) return { x: box.x, y: box.y, w: box.w, h: box.h }
  const r = (box.rotationDeg * Math.PI) / 180
  const c = Math.abs(Math.cos(r))
  const s = Math.abs(Math.sin(r))
  const w = box.w * c + box.h * s
  const h = box.w * s + box.h * c
  return { x: box.x + box.w / 2 - w / 2, y: box.y + box.h / 2 - h / 2, w, h }
}

/** Konva Line tension → cubic/quadratic segments (same control points as Konva's expandPoints). */
function controlPoints(
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  t: number,
): number[] {
  const d01 = Math.hypot(x1 - x0, y1 - y0)
  const d12 = Math.hypot(x2 - x1, y2 - y1)
  const fa = (t * d01) / (d01 + d12)
  const fb = (t * d12) / (d01 + d12)
  return [x1 - fa * (x2 - x0), y1 - fa * (y2 - y0), x1 + fb * (x2 - x0), y1 + fb * (y2 - y0)]
}

function expandPoints(p: number[], t: number): number[] {
  const out: number[] = []
  for (let i = 2; i < p.length - 2; i += 2) {
    const cp = controlPoints(p[i - 2]!, p[i - 1]!, p[i]!, p[i + 1]!, p[i + 2]!, p[i + 3]!, t)
    if (Number.isNaN(cp[0])) continue
    out.push(cp[0]!, cp[1]!, p[i]!, p[i + 1]!, cp[2]!, cp[3]!)
  }
  return out
}

export function polylinePath(points: number[], tension: number, closed: boolean): string {
  if (points.length < 4) return ''
  let d = `M${n(points[0]!)} ${n(points[1]!)}`
  if (tension !== 0 && points.length > 4) {
    if (closed) {
      const len = points.length
      const first = controlPoints(
        points[len - 2]!,
        points[len - 1]!,
        points[0]!,
        points[1]!,
        points[2]!,
        points[3]!,
        tension,
      )
      const last = controlPoints(
        points[len - 4]!,
        points[len - 3]!,
        points[len - 2]!,
        points[len - 1]!,
        points[0]!,
        points[1]!,
        tension,
      )
      const tp = [first[2]!, first[3]!]
        .concat(expandPoints(points, tension))
        .concat([
          last[0]!,
          last[1]!,
          points[len - 2]!,
          points[len - 1]!,
          last[2]!,
          last[3]!,
          first[0]!,
          first[1]!,
          points[0]!,
          points[1]!,
        ])
      for (let i = 0; i + 5 < tp.length; i += 6)
        d += `C${n(tp[i]!)} ${n(tp[i + 1]!)} ${n(tp[i + 2]!)} ${n(tp[i + 3]!)} ${n(tp[i + 4]!)} ${n(tp[i + 5]!)}`
    } else {
      const tp = expandPoints(points, tension)
      const len = tp.length
      d += `Q${n(tp[0]!)} ${n(tp[1]!)} ${n(tp[2]!)} ${n(tp[3]!)}`
      for (let i = 4; i + 5 < len; i += 6)
        d += `C${n(tp[i]!)} ${n(tp[i + 1]!)} ${n(tp[i + 2]!)} ${n(tp[i + 3]!)} ${n(tp[i + 4]!)} ${n(tp[i + 5]!)}`
      d += `Q${n(tp[len - 2]!)} ${n(tp[len - 1]!)} ${n(points[points.length - 2]!)} ${n(points[points.length - 1]!)}`
    }
  } else {
    for (let i = 2; i < points.length; i += 2) d += `L${n(points[i]!)} ${n(points[i + 1]!)}`
  }
  return closed ? d + 'Z' : d
}

/** Konva Arc (12 o'clock = -90°, clockwise sweep) as an annular sector path. */
export function arcPath(
  cx: number,
  cy: number,
  innerR: number,
  outerR: number,
  startDeg: number,
  sweepDeg: number,
): string {
  const sweep = Math.max(-360, Math.min(360, sweepDeg))
  const full = Math.abs(sweep) >= 360
  const a0 = (startDeg * Math.PI) / 180
  const a1 = ((startDeg + (full ? sweep / 2 : sweep)) * Math.PI) / 180
  const a2 = ((startDeg + sweep) * Math.PI) / 180
  const pt = (r: number, a: number) => `${n(cx + r * Math.cos(a))} ${n(cy + r * Math.sin(a))}`
  const dir = sweep >= 0 ? 1 : 0
  const large = !full && Math.abs(sweep) > 180 ? 1 : 0
  let d = `M${pt(outerR, a0)}`
  if (full)
    d += `A${n(outerR)} ${n(outerR)} 0 0 ${dir} ${pt(outerR, a1)}A${n(outerR)} ${n(outerR)} 0 0 ${dir} ${pt(outerR, a2)}`
  else d += `A${n(outerR)} ${n(outerR)} 0 ${large} ${dir} ${pt(outerR, a2)}`
  if (innerR > 0) {
    d += `L${pt(innerR, a2)}`
    if (full)
      d += `A${n(innerR)} ${n(innerR)} 0 0 ${1 - dir} ${pt(innerR, a1)}A${n(innerR)} ${n(innerR)} 0 0 ${1 - dir} ${pt(innerR, a0)}`
    else d += `A${n(innerR)} ${n(innerR)} 0 ${large} ${1 - dir} ${pt(innerR, a0)}`
  } else d += `L${n(cx)} ${n(cy)}`
  return d + 'Z'
}

function shadeHex(color: string, f: number): string {
  const m = /^#([0-9a-f]{6})/i.exec(color)
  if (!m) return color
  const ch = (i: number) =>
    Math.round(parseInt(m[1]!.slice(i, i + 2), 16) * f)
      .toString(16)
      .padStart(2, '0')
  return `#${ch(0)}${ch(2)}${ch(4)}`
}

/** Fill/stroke paint attributes (color or a url(#def) reference plus opacity). */
interface Paint {
  paint: string
  opacity?: number
}

class PageBuilder {
  private defs: string[] = []
  private ids = 0
  readonly fontFamilies = new Set<string>()

  constructor(
    readonly images: Map<string, HTMLImageElement>,
    readonly rasters: Map<string, NodeRaster>,
    private readonly idPrefix: string,
  ) {}

  id(kind: string): string {
    return `${this.idPrefix}${kind}${++this.ids}`
  }

  def(markup: string): void {
    this.defs.push(markup)
  }

  defsMarkup(): string {
    return this.defs.length ? `<defs>${this.defs.join('')}</defs>` : ''
  }

  gradientStops(stops: Array<{ pos: number; color: string }>): string {
    const ramp = linearRampStops(stops)
    let out = ''
    for (let i = 0; i + 1 < ramp.length; i += 2) {
      const off = Math.max(0, Math.min(1, ramp[i] as number))
      const c = splitAlpha(ramp[i + 1] as string)
      out += el('stop', { offset: off, 'stop-color': c.color, 'stop-opacity': c.opacity })
    }
    return out
  }

  linearGradient(
    stops: Array<{ pos: number; color: string }>,
    angleDeg: number,
    scaled: boolean | undefined,
    w: number,
    h: number,
    o: Origin,
  ): string {
    const [dx, dy] = linearGradientDirection(angleDeg, scaled, w, h)
    const cx = o.x + w / 2
    const cy = o.y + h / 2
    const len = Math.abs(dx) * w + Math.abs(dy) * h
    const id = this.id('lg')
    this.def(
      el(
        'linearGradient',
        {
          id,
          gradientUnits: 'userSpaceOnUse',
          x1: cx - (dx * len) / 2,
          y1: cy - (dy * len) / 2,
          x2: cx + (dx * len) / 2,
          y2: cy + (dy * len) / 2,
        },
        this.gradientStops(stops),
      ),
    )
    return `url(#${id})`
  }

  /** A shape-sized image used as a fill: no-repeat pattern anchored at the shape origin. */
  imagePattern(src: ImageSource, w: number, h: number, o: Origin, repeat = false): string | null {
    const href = imageHref(src)
    if (!href) return null
    const id = this.id('pt')
    const size = srcSize(src)
    const tileW = repeat ? size.w : w
    const tileH = repeat ? size.h : h
    this.def(
      el(
        'pattern',
        { id, patternUnits: 'userSpaceOnUse', x: o.x, y: o.y, width: tileW, height: tileH },
        el('image', {
          href,
          width: tileW,
          height: tileH,
          preserveAspectRatio: 'none',
        }),
      ),
    )
    return `url(#${id})`
  }

  /** RenderFill → paint. null = nothing to paint. */
  fillPaint(
    fill: RenderFill | undefined,
    w: number,
    h: number,
    o: Origin = { x: 0, y: 0 },
  ): Paint | null {
    if (!fill || fill.kind === 'none') return null
    switch (fill.kind) {
      case 'solid':
        return { paint: normalizeColor(fill.color) }
      case 'pattern': {
        const cv = patternCanvas(fill, w, h, { x: o.x, y: o.y })
        const p = this.imagePattern(cv, w, h, o)
        return p ? { paint: p } : null
      }
      case 'gradient': {
        if (fill.radial) {
          if (fill.path === 'rect' || fill.path === 'shape') {
            const cv = pathGradientCanvas(
              fill.path,
              fill.stops,
              w,
              h,
              fill.center?.x ?? 0.5,
              fill.center?.y ?? 0.5,
            )
            if (cv) {
              const p = this.imagePattern(cv, w, h, o)
              if (p) return { paint: p }
            }
          }
          const { cx, cy, r } = radialCircleGeometry(w, h, fill.center, fill.tileRect)
          const id = this.id('rg')
          this.def(
            el(
              'radialGradient',
              { id, gradientUnits: 'userSpaceOnUse', cx: o.x + cx, cy: o.y + cy, r },
              this.gradientStops(fill.stops),
            ),
          )
          return { paint: `url(#${id})` }
        }
        return { paint: this.linearGradient(fill.stops, fill.angleDeg, fill.scaled, w, h, o) }
      }
      case 'image': {
        const img = fill.dataUrl ? this.images.get(fill.dataUrl) : undefined
        if (!img) return null
        const src = processedImage(img, fill.dataUrl ?? '', fill) as ImageSource
        const srcKey = processedImageKey(fill.dataUrl ?? '', fill)
        const opacity = fill.alpha != null ? { opacity: fill.alpha } : {}
        if (fill.mode === 'tile') {
          const tile = fill.tile ? anchoredTileCanvas(src, srcKey, w, h, fill.tile) : src
          const p = this.imagePattern(tile, w, h, o, !fill.tile)
          return p ? { paint: p, ...opacity } : null
        }
        if (isDegenerateImage(img)) return { paint: averageColor(src, srcKey), ...opacity }
        const tile = fill.fillRect ? insetFillTile(src, srcKey, fill.fillRect) : src
        const p = this.imagePattern(tile, w, h, o)
        return p ? { paint: p, ...opacity } : null
      }
      default:
        return null
    }
  }

  fillAttrs(paint: Paint | null): Record<string, string | number | undefined> {
    if (!paint) return { fill: 'none' }
    return { fill: paint.paint, 'fill-opacity': paint.opacity }
  }

  strokeAttrs(
    stroke: RenderStroke | undefined,
    size?: { w: number; h: number },
    o: Origin = { x: 0, y: 0 },
  ): Record<string, string | number | undefined> {
    if (!stroke) return {}
    const paint =
      stroke.gradient && size
        ? this.linearGradient(
            stroke.gradient.stops,
            stroke.gradient.angleDeg,
            stroke.gradient.scaled,
            size.w,
            size.h,
            o,
          )
        : normalizeColor(stroke.color)
    // only defined keys: callers spread these over their own defaults (e.g. a round join)
    return {
      stroke: paint,
      'stroke-width': stroke.widthPx,
      ...(stroke.dash?.length ? { 'stroke-dasharray': stroke.dash.map(n).join(' ') } : {}),
      ...(stroke.cap ? { 'stroke-linecap': stroke.cap } : {}),
      ...(stroke.join ? { 'stroke-linejoin': stroke.join } : {}),
    }
  }

  /**
   * Outer shadow / glow as a drop-shadow filter. Canvas shadowBlur b ≈ gaussian σ = b/2.
   * The region is explicit (userSpaceOnUse): percentage regions collapse on hairline geometry.
   */
  shadowFilter(
    rawShadow: RenderShadow | undefined,
    glow: RenderGlow | undefined,
    region: Bounds,
  ): string | undefined {
    const shadow = isOverlayShadow(rawShadow) ? undefined : rawShadow
    const s = shadow
      ? {
          color: shadow.color,
          blur: shadow.blurPx,
          dx: shadow.offsetX,
          dy: shadow.offsetY,
        }
      : glow
        ? { color: glow.color, blur: glow.blurPx, dx: 0, dy: 0 }
        : null
    if (!s) return undefined
    const c = splitAlpha(s.color)
    const pad = s.blur * 3 + Math.max(Math.abs(s.dx), Math.abs(s.dy)) + 4
    const id = this.id('sh')
    this.def(
      el(
        'filter',
        {
          id,
          filterUnits: 'userSpaceOnUse',
          x: region.x - pad,
          y: region.y - pad,
          width: region.w + 2 * pad,
          height: region.h + 2 * pad,
        },
        el('feDropShadow', {
          dx: s.dx,
          dy: s.dy,
          stdDeviation: s.blur / 2,
          'flood-color': c.color,
          'flood-opacity': c.opacity,
        }),
      ),
    )
    return `url(#${id})`
  }

  blurFilter(blurPx: number, region: Bounds): string {
    const id = this.id('bl')
    const pad = blurPx * 3 + 2
    this.def(
      el(
        'filter',
        {
          id,
          filterUnits: 'userSpaceOnUse',
          x: region.x - pad,
          y: region.y - pad,
          width: region.w + 2 * pad,
          height: region.h + 2 * pad,
        },
        el('feGaussianBlur', { stdDeviation: blurPx / 2 }),
      ),
    )
    return `url(#${id})`
  }

  clipPath(inner: string): string {
    const id = this.id('cp')
    this.def(el('clipPath', { id }, inner))
    return `url(#${id})`
  }

  image(src: ImageSource, a: Record<string, string | number | undefined>): string {
    const href = imageHref(src)
    if (!href) return ''
    return el('image', { href, preserveAspectRatio: 'none', ...a })
  }
}

// ── geometry ────────────────────────────────────────────────────────────

/** The closed outline of a shape/picture clip as an element (fill/stroke attrs supplied). */
function outlineElement(
  geom: { pathData?: string; polygonPoints?: number[]; cornerRadiusPx?: number; ellipse?: boolean },
  w: number,
  h: number,
  a: Record<string, string | number | undefined>,
): string {
  if (geom.pathData) return el('path', { d: geom.pathData, 'stroke-linejoin': 'round', ...a })
  if (geom.polygonPoints)
    return el('polygon', {
      points: geom.polygonPoints.map(n).join(' '),
      'stroke-linejoin': 'round',
      ...a,
    })
  if (geom.ellipse) return el('ellipse', { cx: w / 2, cy: h / 2, rx: w / 2, ry: h / 2, ...a })
  // Konva clamps cornerRadius to half the short side
  const r = Math.min(geom.cornerRadiusPx ?? 0, Math.min(w, h) / 2)
  return el('rect', { width: w, height: h, rx: r || undefined, ry: r || undefined, ...a })
}

function shapeGeom(shape: ShapeRenderNode): {
  pathData?: string
  polygonPoints?: number[]
  cornerRadiusPx?: number
  ellipse?: boolean
} {
  if (shape.fillPathData || shape.pathData)
    return { pathData: shape.fillPathData ?? shape.pathData }
  if (shape.polygonPoints) return { polygonPoints: shape.polygonPoints }
  const kind = presetToShapeKind(shape.presetGeometry)
  if (kind === 'ellipse') return { ellipse: true }
  const rounded = kind === 'roundRect' || shape.cornerRadiusPx != null
  return {
    cornerRadiusPx: rounded
      ? (shape.cornerRadiusPx ?? Math.min(shape.box.w, shape.box.h) * 0.167)
      : 0,
  }
}

function shadowGeom(shape: ShapeRenderNode): ShadowGeom {
  const { w, h } = shape.box
  if (shape.fillPathData || shape.pathData)
    return { kind: 'path', data: (shape.fillPathData ?? shape.pathData)! }
  if (shape.polygonPoints) return { kind: 'polygon', points: shape.polygonPoints }
  if (presetToShapeKind(shape.presetGeometry) === 'ellipse') return { kind: 'ellipse', w, h }
  return {
    kind: 'rect',
    w,
    h,
    cornerRadius:
      presetToShapeKind(shape.presetGeometry) === 'roundRect'
        ? (shape.cornerRadiusPx ?? Math.min(w, h) * 0.167)
        : (shape.cornerRadiusPx ?? 0),
  }
}

function overlayImage(
  b: PageBuilder,
  ov: { canvas: HTMLCanvasElement; x: number; y: number; w: number; h: number } | null,
  opacity?: number,
): string {
  if (!ov) return ''
  return b.image(ov.canvas, { x: ov.x, y: ov.y, width: ov.w, height: ov.h, opacity })
}

// ── arrows ──────────────────────────────────────────────────────────────

function arrowHead(end: ArrowEndRender, pts: number[], atStart: boolean, color: string): string {
  const nPts = pts.length / 2
  if (nPts < 2) return ''
  let angle: number
  let tipX: number
  let tipY: number
  if (atStart) {
    angle = Math.atan2(pts[1]! - pts[3]!, pts[0]! - pts[2]!) * (180 / Math.PI)
    tipX = pts[0]!
    tipY = pts[1]!
  } else {
    const lx = pts[(nPts - 1) * 2]!
    const ly = pts[(nPts - 1) * 2 + 1]!
    angle = Math.atan2(ly - pts[(nPts - 2) * 2 + 1]!, lx - pts[(nPts - 2) * 2]!) * (180 / Math.PI)
    tipX = lx
    tipY = ly
  }
  const w = end.widthPx
  const l = end.lengthPx
  const halfW = w / 2
  const tf = `translate(${n(tipX)} ${n(tipY)}) rotate(${n(angle)})`
  if (end.type === 'oval')
    return el('ellipse', { rx: l / 2, ry: halfW, fill: color, transform: tf })
  if (end.type === 'diamond')
    return el('polygon', {
      points: [0, 0, -l / 2, -halfW, -l, 0, -l / 2, halfW].map(n).join(' '),
      fill: color,
      transform: tf,
    })
  if (end.type === 'stealth')
    return el('polygon', {
      points: [0, 0, -l, -halfW, -l * 0.65, 0, -l, halfW].map(n).join(' '),
      fill: color,
      transform: tf,
    })
  return el('polyline', {
    points: [-l, -halfW, 0, 0, -l, halfW].map(n).join(' '),
    fill: 'none',
    stroke: color,
    'stroke-width': Math.max(1, w * 0.12),
    'stroke-linecap': 'round',
    'stroke-linejoin': 'round',
    transform: tf,
  })
}

/** Konva Arrow's built-in triangle pointer (filled and stroked with the line color). */
function konvaPointer(
  pts: number[],
  atStart: boolean,
  length: number,
  width: number,
  color: string,
  sw: number,
): string {
  const nPts = pts.length
  if (nPts < 4) return ''
  const [tx, ty, px, py] = atStart
    ? [pts[0]!, pts[1]!, pts[2]!, pts[3]!]
    : [pts[nPts - 2]!, pts[nPts - 1]!, pts[nPts - 4]!, pts[nPts - 3]!]
  const angle = Math.atan2(ty - py, tx - px) * (180 / Math.PI)
  return el('polygon', {
    points: [0, 0, -length, width / 2, -length, -width / 2].map(n).join(' '),
    fill: color,
    stroke: color,
    'stroke-width': sw,
    'stroke-linejoin': 'round',
    transform: `translate(${n(tx)} ${n(ty)}) rotate(${n(angle)})`,
  })
}

// ── text ────────────────────────────────────────────────────────────────

/** One Konva-Text-equivalent glyph run as SVG text (x0/y0 = Konva node origin = run top-left). */
function glyphText(
  b: PageBuilder,
  g: GlyphDraw,
  run: { baselineY: number; widthPx: number; text: string; justifyExtraPx?: number },
  x0: number,
  y0: number,
  extra: Record<string, string | number | undefined> = {},
): string {
  const bold = g.fontStyle.includes('bold')
  const italic = g.fontStyle.includes('italic')
  // Konva's top-anchored y lands the alphabetic baseline on baselineY for resolved faces
  // and konvaBaselineDrop below it for fallback-drawn ones (the engine's legacy 0.8em rule)
  const baseline = run.baselineY + konvaBaselineDrop(g.fontFamily, g.fontSize, bold, italic)
  const dy = baseline - g.y
  b.fontFamilies.add(g.fontFamily)
  const trailing = g.letterSpacing ?? 0
  const chars = [...run.text]
  const fitWidth =
    chars.length >= 2 && run.text.trim() !== '' && run.justifyExtraPx == null
      ? Math.max(run.widthPx - trailing, 0)
      : 0
  const rotation = g.rotation ?? 0
  const a: Record<string, string | number | undefined> = {
    x: g.direction === 'rtl' ? x0 + run.widthPx : x0,
    y: y0 + dy,
    'font-family': g.fontFamily,
    'font-size': g.fontSize,
    'font-weight': bold ? 700 : undefined,
    'font-style': italic ? 'italic' : undefined,
    'letter-spacing': trailing && Math.abs(trailing) > 1e-3 ? trailing : undefined,
    'text-decoration': g.textDecoration || undefined,
    direction: g.direction,
    'unicode-bidi': g.direction ? 'bidi-override' : undefined,
    textLength: fitWidth > 0 ? fitWidth : undefined,
    lengthAdjust: fitWidth > 0 ? 'spacingAndGlyphs' : undefined,
    'xml:space': 'preserve',
    transform: rotation ? `rotate(${n(rotation)} ${n(x0)} ${n(y0)})` : undefined,
    ...extra,
  }
  if (g.fillPriority === 'linear-gradient' && g.fillLinearGradientColorStops && !extra.fill) {
    const id = b.id('tg')
    const s = g.fillLinearGradientStartPoint!
    const e = g.fillLinearGradientEndPoint!
    const stops = g.fillLinearGradientColorStops
    let stopMarkup = ''
    for (let i = 0; i + 1 < stops.length; i += 2) {
      const c = splitAlpha(stops[i + 1] as string)
      stopMarkup += el('stop', {
        offset: Math.max(0, Math.min(1, stops[i] as number)),
        'stop-color': c.color,
        'stop-opacity': c.opacity,
      })
    }
    b.def(
      el(
        'linearGradient',
        {
          id,
          gradientUnits: 'userSpaceOnUse',
          x1: x0 + s.x,
          y1: y0 + s.y,
          x2: x0 + e.x,
          y2: y0 + e.y,
        },
        stopMarkup,
      ),
    )
    a.fill = `url(#${id})`
  } else if (!extra.fill) a.fill = g.fill
  if (g.stroke) {
    a.stroke = g.stroke
    a['stroke-width'] = g.strokeWidth
    a['paint-order'] = 'stroke'
    a['stroke-linejoin'] = 'round'
  }
  if (g.shadowEnabled && g.shadowColor) {
    a.filter = b.shadowFilter(
      {
        color: g.shadowColor,
        blurPx: g.shadowBlur ?? 0,
        offsetX: g.shadowOffsetX ?? 0,
        offsetY: g.shadowOffsetY ?? 0,
      },
      undefined,
      { x: x0, y: y0, w: run.widthPx, h: g.fontSize * 1.5 },
    )
  }
  return el('text', a, escText(run.text))
}

/** Highlights, extrusion layers, glyph runs and reflections of one laid-out text body. */
function textBody(
  b: PageBuilder,
  text: RenderTextLayout | undefined,
  hideCell?: boolean,
  dx = 0,
  dy = 0,
): string {
  if (!text || hideCell) return ''
  const il = dx + text.insets.l
  const it = dy + text.insets.t
  let highlights = ''
  let extrusion = ''
  let body = ''
  let reflections = ''
  for (const line of text.lines) {
    for (const run of line.runs) {
      const g = glyphToDraw(run)
      const x0 = il + g.x
      const y0 = it + g.y
      if (run.highlight && !text.vert)
        highlights += el('rect', {
          x: il + run.x,
          y: it + line.top,
          width: run.widthPx,
          height: line.height,
          fill: normalizeColor(run.highlight),
        })
      if (g.image) {
        const img = b.images.get(g.image)
        if (img)
          body += b.image(img, {
            x: x0,
            y: y0,
            width: g.imageW ?? g.fontSize,
            height: g.imageH ?? g.fontSize,
          })
        continue
      }
      if (text.extrusion) {
        for (const k of [1, 0.5])
          extrusion += glyphText(
            b,
            g,
            run,
            x0 + text.extrusion.dx * k,
            y0 + text.extrusion.dy * k,
            { fill: normalizeColor(shadeHex(text.extrusion.color, 0.7)) },
          )
      }
      body += glyphText(
        b,
        g,
        run,
        x0,
        y0,
        text.extrusion ? { fill: normalizeColor(shadeHex(text.extrusion.color, 0.35)) } : {},
      )
      if (g.reflection) {
        // Konva: a copy top-anchored 1.8em below, scaleY −1 about its own origin
        const yp = y0 + g.fontSize * 1.8
        reflections += `<g transform="translate(0 ${n(2 * yp)}) scale(1 -1)" opacity="0.15">${glyphText(
          b,
          { ...g, shadowEnabled: false, stroke: undefined, textDecoration: '' },
          run,
          x0,
          yp,
        )}</g>`
      }
    }
  }
  return highlights + extrusion + body + reflections
}

// ── nodes ───────────────────────────────────────────────────────────────

interface FlipParity {
  h: boolean
  v: boolean
}

function emitPicture(b: PageBuilder, pic: PictureRenderNode): string {
  const { box } = pic
  const rawImg = pic.dataUrl ? b.images.get(pic.dataUrl) : undefined
  const srcKey = processedImageKey(pic.dataUrl ?? '', pic)
  const procImg =
    rawImg && hasBlipEffects(pic)
      ? (processedImage(rawImg, pic.dataUrl ?? '', pic) as ImageSource)
      : rawImg
  const tiny = !!procImg && isDegenerateImage(procImg)
  const img =
    procImg && tiny ? (flatColorImage(procImg as HTMLImageElement, srcKey) as ImageSource) : procImg
  const clip = pic.clip
  const cropProps = tiny || !rawImg ? {} : cropToKonva(pic, rawImg)
  const inset = 'x' in cropProps
  const shadowOv = shapeShadowOverlay(
    pic.shadow,
    { kind: 'rect', w: box.w, h: box.h, cornerRadius: clip?.cornerRadiusPx },
    box.w,
    box.h,
  )
  const opacity = pic.opacity != null ? pic.opacity : undefined
  const region = { x: 0, y: 0, w: box.w, h: box.h }

  let image = ''
  if (img) {
    const drawn: ImageSource =
      pic.softEdgePx && srcSize(img).w
        ? (featheredImage(
            img as HTMLImageElement,
            srcKey,
            pic.softEdgePx * (srcSize(img).w / Math.max(box.w, 1)),
          ) as ImageSource)
        : img
    const dest = inset
      ? { x: cropProps.x!, y: cropProps.y!, w: cropProps.width!, h: cropProps.height! }
      : { x: 0, y: 0, w: box.w, h: box.h }
    const plain: Record<string, string | number | undefined> = clip
      ? {}
      : {
          ...(inset ? {} : b.strokeAttrs(pic.stroke)),
          filter: b.shadowFilter(pic.shadow, pic.glow, region),
        }
    const crop = 'crop' in cropProps ? cropProps.crop : undefined
    if (crop && !tiny) {
      // Konva crop: a source sub-rect stretched into the destination box
      const href = imageHref(drawn)
      const size = srcSize(drawn)
      if (href) {
        const scale = size.w / Math.max(srcSize(img).w, 1)
        image = `<svg${attrs({
          x: dest.x,
          y: dest.y,
          width: dest.w,
          height: dest.h,
          viewBox: `${n(crop.x * scale)} ${n(crop.y * scale)} ${n(crop.width * scale)} ${n(crop.height * scale)}`,
          preserveAspectRatio: 'none',
          overflow: 'hidden',
          opacity,
        })}>${el('image', { href, width: size.w, height: size.h, preserveAspectRatio: 'none' })}</svg>`
        if (plain.filter || plain.stroke)
          image =
            el('rect', {
              x: dest.x,
              y: dest.y,
              width: dest.w,
              height: dest.h,
              fill: 'none',
              ...plain,
            }) + image
      }
    } else {
      image = b.image(drawn, {
        x: dest.x,
        y: dest.y,
        width: dest.w,
        height: dest.h,
        opacity,
        ...plain,
      })
    }
    if (inset && !clip)
      image =
        el('rect', { width: box.w, height: box.h, fill: 'none', ...b.strokeAttrs(pic.stroke) }) +
        image
  } else if (pic.dataUrl) {
    image = el('rect', {
      width: box.w,
      height: box.h,
      fill: '#eef',
      stroke: '#99f',
      'stroke-dasharray': '4 4',
    })
  }

  const out: string[] = []
  if (shadowOv?.under) out.push(overlayImage(b, shadowOv))
  if (clip && img) {
    const geom = {
      pathData: clip.pathData,
      polygonPoints: clip.polygonPoints,
      cornerRadiusPx: clip.cornerRadiusPx,
    }
    const outline = (a: Record<string, string | number | undefined>) =>
      outlineElement(geom, box.w, box.h, a)
    const backing = (a: Record<string, string | number | undefined>) =>
      inset
        ? el('rect', {
            x: cropProps.x,
            y: cropProps.y,
            width: cropProps.width,
            height: cropProps.height,
            ...a,
          })
        : outline(a)
    const filter = b.shadowFilter(pic.shadow, pic.glow, region)
    if (filter) out.push(backing({ fill: '#ffffff', filter, opacity }))
    const clipUrl = b.clipPath(outline({}))
    let inner = ''
    if (pic.fill) inner += outline(b.fillAttrs(b.fillPaint(pic.fill, box.w, box.h)))
    if (pic.bgColor && (pic.opacity ?? 1) >= 1) inner += backing({ fill: pic.bgColor })
    inner += image
    out.push(`<g clip-path="${clipUrl}">${inner}</g>`)
    if (pic.stroke) out.push(outline({ fill: 'none', ...b.strokeAttrs(pic.stroke) }))
  } else {
    if (pic.fill)
      out.push(
        el('rect', {
          width: box.w,
          height: box.h,
          ...b.fillAttrs(b.fillPaint(pic.fill, box.w, box.h)),
        }),
      )
    if (pic.bgColor && img && (pic.opacity ?? 1) >= 1)
      out.push(el('rect', { width: box.w, height: box.h, fill: pic.bgColor }))
    out.push(image)
  }
  if (shadowOv && !shadowOv.under) out.push(overlayImage(b, shadowOv))
  return out.join('')
}

function bevelBands(b: PageBuilder, cell: TableCellRender, bevel: CellBevelRender): string {
  const { x, y, w, h } = cell
  const bw = bevel.widthPx
  const bands: Array<{
    key: keyof CellBevelRender['edges']
    points: number[]
    from: Origin
    to: Origin
  }> = [
    {
      key: 't',
      points: [x, y, x + w, y, x + w - bw, y + bw, x + bw, y + bw],
      from: { x, y },
      to: { x, y: y + bw },
    },
    {
      key: 'r',
      points: [x + w, y, x + w, y + h, x + w - bw, y + h - bw, x + w - bw, y + bw],
      from: { x: x + w, y },
      to: { x: x + w - bw, y },
    },
    {
      key: 'b',
      points: [x, y + h, x + w, y + h, x + w - bw, y + h - bw, x + bw, y + h - bw],
      from: { x, y: y + h },
      to: { x, y: y + h - bw },
    },
    {
      key: 'l',
      points: [x, y, x, y + h, x + bw, y + h - bw, x + bw, y + bw],
      from: { x, y },
      to: { x: x + bw, y },
    },
  ]
  let out = ''
  for (const band of bands) {
    const id = b.id('bv')
    let stops = ''
    for (const s of bevel.edges[band.key]) {
      const c = splitAlpha(s.color)
      stops += el('stop', { offset: s.pos, 'stop-color': c.color, 'stop-opacity': c.opacity })
    }
    b.def(
      el(
        'linearGradient',
        {
          id,
          gradientUnits: 'userSpaceOnUse',
          x1: band.from.x,
          y1: band.from.y,
          x2: band.to.x,
          y2: band.to.y,
        },
        stops,
      ),
    )
    out += el('polygon', { points: band.points.map(n).join(' '), fill: `url(#${id})` })
  }
  return out
}

function emitTable(b: PageBuilder, table: TableRenderNode): string {
  const { box } = table
  const tblW = table.gridX[table.gridX.length - 1] ?? box.w
  const tblH = table.gridY[table.gridY.length - 1] ?? box.h
  const out: string[] = []
  if (table.bgFill)
    out.push(
      el('rect', {
        width: tblW,
        height: tblH,
        ...b.fillAttrs(b.fillPaint(table.bgFill, tblW, tblH)),
      }),
    )
  for (const cell of table.cells) {
    const bw = cell.bevel?.widthPx ?? 0
    out.push(
      el('rect', {
        x: cell.x + bw,
        y: cell.y + bw,
        width: cell.w - 2 * bw,
        height: cell.h - 2 * bw,
        ...b.fillAttrs(b.fillPaint(cell.fill, cell.w, cell.h, { x: cell.x, y: cell.y })),
      }),
    )
    if (cell.bevel) out.push(bevelBands(b, cell, cell.bevel))
    const line = (x1: number, y1: number, x2: number, y2: number, s: RenderStroke | undefined) =>
      s ? el('line', { x1, y1, x2, y2, ...b.strokeAttrs(s) }) : ''
    out.push(line(cell.x, cell.y, cell.x + cell.w, cell.y, cell.borders?.t))
    out.push(line(cell.x, cell.y + cell.h, cell.x + cell.w, cell.y + cell.h, cell.borders?.b))
    out.push(line(cell.x, cell.y, cell.x, cell.y + cell.h, cell.borders?.l))
    out.push(line(cell.x + cell.w, cell.y, cell.x + cell.w, cell.y + cell.h, cell.borders?.r))
    out.push(textBody(b, cell.text, false, cell.x, cell.y))
  }
  return out.join('')
}

function emitChart(b: PageBuilder, chart: ChartRenderNode): string {
  const out: string[] = []
  const { box } = chart
  if (chart.bgFill)
    out.push(
      el('rect', {
        width: box.w,
        height: box.h,
        ...b.fillAttrs(b.fillPaint(chart.bgFill, box.w, box.h)),
      }),
    )
  if (chart.plotRect) {
    const p = chart.plotRect
    out.push(
      el('rect', {
        x: p.x,
        y: p.y,
        width: p.w,
        height: p.h,
        ...(p.fill
          ? b.fillAttrs(b.fillPaint(p.fill, p.w, p.h, { x: p.x, y: p.y }))
          : { fill: 'none' }),
        ...(p.borderColor ? { stroke: p.borderColor, 'stroke-width': p.borderWidthPx ?? 1 } : {}),
      }),
    )
  }
  for (const wd of chart.wedges ?? [])
    out.push(
      el('path', {
        d: arcPath(wd.cx, wd.cy, wd.innerR, wd.outerR, wd.startDeg, wd.sweepDeg),
        fill: wd.noFill ? 'none' : wd.color,
        ...(wd.strokeWidthPx === 0
          ? {}
          : { stroke: wd.stroke ?? '#ffffff', 'stroke-width': wd.strokeWidthPx ?? 1 }),
      }),
    )
  for (const g of chart.gridLines)
    out.push(
      el('line', {
        x1: g.x1,
        y1: g.y1,
        x2: g.x2,
        y2: g.y2,
        stroke: g.color,
        'stroke-width': g.widthPx ?? 1,
        'stroke-dasharray': g.dash?.length ? g.dash.map(n).join(' ') : undefined,
      }),
    )
  for (const a of chart.axisLines)
    out.push(
      el('line', {
        x1: a.x1,
        y1: a.y1,
        x2: a.x2,
        y2: a.y2,
        stroke: a.color,
        'stroke-width': a.widthPx,
      }),
    )
  for (const p of chart.paths ?? [])
    out.push(
      el('path', {
        d: p.d,
        fill: p.fill,
        transform: p.dy ? `translate(0 ${n(p.dy)})` : undefined,
        ...(p.stroke ? { stroke: p.stroke, 'stroke-width': p.strokeWidthPx ?? 1 } : {}),
      }),
    )
  for (const bar of chart.bars)
    out.push(
      el('rect', {
        x: bar.x,
        y: bar.y,
        width: bar.w,
        height: bar.h,
        ...(bar.fill
          ? b.fillAttrs(b.fillPaint(bar.fill, bar.w, bar.h, { x: bar.x, y: bar.y }))
          : { fill: bar.color }),
      }),
    )
  for (const p of chart.polylines)
    out.push(
      el('path', {
        d: polylinePath(p.points, smoothTension(p.smooth), !!p.closed),
        fill: p.fill ?? 'none',
        stroke: p.color,
        'stroke-width': p.widthPx,
        'stroke-linecap': 'round',
        'stroke-linejoin': 'round',
        'stroke-dasharray': p.dash?.length ? p.dash.map(n).join(' ') : undefined,
      }),
    )
  for (const m of chart.markers) out.push(el('circle', { cx: m.x, cy: m.y, r: m.r, fill: m.color }))
  for (const s of chart.swatches)
    out.push(el('rect', { x: s.x, y: s.y, width: s.w, height: s.h, fill: s.color, rx: 1, ry: 1 }))
  for (const l of chart.labels) {
    const family = l.fontFamily ? displayFontFamily(l.fontFamily) : CHART_FONT
    b.fontFamilies.add(family)
    if (l.fill || l.stroke)
      out.push(
        el('rect', {
          x: l.x - l.fontSizePx * 0.3,
          y: l.y - l.fontSizePx * 0.1,
          width: (l.w ?? 0) + l.fontSizePx * 0.6,
          height: l.fontSizePx * 1.35,
          fill: l.fill ?? 'none',
          ...(l.stroke ? { stroke: l.stroke, 'stroke-width': 1 } : {}),
        }),
      )
    out.push(
      el(
        'text',
        {
          x: l.x,
          // Konva anchors single-line text on the em-box middle half a line below the top
          y: l.y + l.fontSizePx / 2,
          'dominant-baseline': 'central',
          'font-family': family,
          'font-size': l.fontSizePx,
          'font-weight': l.bold ? 700 : undefined,
          'font-style': l.italic ? 'italic' : undefined,
          fill: l.color,
          'xml:space': 'preserve',
          transform: l.rotationDeg ? `rotate(${n(l.rotationDeg)} ${n(l.x)} ${n(l.y)})` : undefined,
        },
        escText(l.text),
      ),
    )
  }
  if (chart.border)
    out.push(
      el('rect', {
        x: chart.border.widthPx / 2,
        y: chart.border.widthPx / 2,
        width: box.w - chart.border.widthPx,
        height: box.h - chart.border.widthPx,
        fill: 'none',
        stroke: chart.border.color,
        'stroke-width': chart.border.widthPx,
      }),
    )
  return out.join('')
}

function emitChip(b: PageBuilder, chip: ChipRenderNode): string {
  const { box } = chip
  b.fontFamilies.add('sans-serif')
  return (
    el('rect', {
      width: box.w,
      height: box.h,
      fill: '#f5f5f7',
      stroke: '#c7c7cc',
      'stroke-dasharray': '6 4',
      rx: 4,
      ry: 4,
    }) +
    el(
      'text',
      {
        x: box.w / 2,
        y: box.h / 2,
        'text-anchor': 'middle',
        'dominant-baseline': 'central',
        'font-size': Math.min(16, box.h / 2),
        'font-family': 'sans-serif',
        fill: '#8e8e93',
      },
      escText(`⧉ ${chip.label}`),
    )
  )
}

function emitLine(b: PageBuilder, shape: ShapeRenderNode): string {
  const line = shape.line!
  const stroke = b.strokeAttrs(shape.stroke, { w: shape.box.w, h: shape.box.h })
  const color = (stroke.stroke as string | undefined) ?? normalizeColor('#000000')
  const sw = (stroke['stroke-width'] as number | undefined) ?? 1
  const dash = stroke['stroke-dasharray']
  const { headEnd, tailEnd, bezier, points } = line
  const bounds = pointsBounds(points)
  const filter = b.shadowFilter(shape.shadow, shape.glow, bounds)
  const common = {
    fill: 'none',
    stroke: color,
    'stroke-width': sw,
    'stroke-dasharray': dash,
    'stroke-linecap': 'round',
    'stroke-linejoin': 'round',
    filter,
  }
  const out: string[] = []
  if (bezier && bezier.length) {
    let d = `M${n(points[0] ?? 0)} ${n(points[1] ?? 0)}`
    for (let i = 0; i + 5 < bezier.length; i += 6)
      d += `C${n(bezier[i]!)} ${n(bezier[i + 1]!)} ${n(bezier[i + 2]!)} ${n(bezier[i + 3]!)} ${n(bezier[i + 4]!)} ${n(bezier[i + 5]!)}`
    out.push(el('path', { d, ...common }))
    if (headEnd) out.push(arrowHead(headEnd, points, true, color))
    if (tailEnd) out.push(arrowHead(tailEnd, points, false, color))
    return out.join('')
  }
  const customHead = headEnd && headEnd.type !== 'arrow' && headEnd.type !== 'triangle'
  const customTail = tailEnd && tailEnd.type !== 'arrow' && tailEnd.type !== 'triangle'
  const konvaArrow = (headEnd || tailEnd) && !customHead && !customTail
  out.push(el('polyline', { points: points.map(n).join(' '), ...common }))
  if (konvaArrow) {
    const len = tailEnd
      ? Math.max(tailEnd.lengthPx, 6)
      : headEnd
        ? Math.max(headEnd.lengthPx, 6)
        : sw * 3.5
    const wid = tailEnd
      ? Math.max(tailEnd.widthPx, 5)
      : headEnd
        ? Math.max(headEnd.widthPx, 5)
        : sw * 3
    if (headEnd) out.push(konvaPointer(points, true, len, wid, color, sw))
    if (tailEnd) out.push(konvaPointer(points, false, len, wid, color, sw))
  } else {
    if (customHead && headEnd) out.push(arrowHead(headEnd, points, true, color))
    if (customTail && tailEnd) out.push(arrowHead(tailEnd, points, false, color))
  }
  return out.join('')
}

function pointsBounds(points: number[]): Bounds {
  let x0 = Infinity
  let y0 = Infinity
  let x1 = -Infinity
  let y1 = -Infinity
  for (let i = 0; i + 1 < points.length; i += 2) {
    x0 = Math.min(x0, points[i]!)
    x1 = Math.max(x1, points[i]!)
    y0 = Math.min(y0, points[i + 1]!)
    y1 = Math.max(y1, points[i + 1]!)
  }
  if (!Number.isFinite(x0)) return { x: 0, y: 0, w: 0, h: 0 }
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 }
}

function emitShape(b: PageBuilder, shape: ShapeRenderNode, flip: FlipParity): string {
  const { box } = shape
  if (shape.line) return emitLine(b, shape)
  if (shape.text?.txWarp && shape.text.lines.some((l) => l.runs.length))
    throw new Error('WordArt warp needs the raster fallback')

  const region = { x: 0, y: 0, w: box.w, h: box.h }
  const fill = b.fillPaint(shape.fill, box.w, box.h)
  const fillA = b.fillAttrs(fill)
  const strokeA = b.strokeAttrs(shape.stroke, { w: box.w, h: box.h })
  const filter = b.shadowFilter(shape.shadow, shape.glow, region)
  const geom = shapeGeom(shape)
  const out: string[] = []

  let geometry = ''
  if (shape.extrusion) {
    const front = extrusionFrontFace(shape.extrusion.faces)
    if (shape.extrusion.shadowPath && front && filter)
      geometry += el('path', {
        d: shape.extrusion.shadowPath,
        ...(front.front ? fillA : { fill: normalizeColor(front.color) }),
        filter,
      })
    for (const f of shape.extrusion.faces)
      geometry += el('path', {
        d: f.path,
        ...(f.front
          ? fillA
          : f.color === 'transparent'
            ? { fill: 'none' }
            : { fill: normalizeColor(f.color) }),
        ...(f.stroke
          ? { stroke: normalizeColor(f.stroke), 'stroke-width': f.strokeWidthPx ?? 1 }
          : {}),
        'stroke-linejoin': 'round',
      })
  } else if (shape.pathData || shape.fillPathData || shape.strokePathData) {
    if (shape.fillPathData) geometry += el('path', { d: shape.fillPathData, ...fillA, filter })
    if (shape.pathData)
      geometry += el('path', {
        d: shape.pathData,
        ...fillA,
        ...strokeA,
        'stroke-linejoin': strokeA['stroke-linejoin'] ?? 'round',
        filter,
      })
    if (shape.strokePathData)
      geometry += el('path', {
        d: shape.strokePathData,
        fill: 'none',
        ...strokeA,
        'stroke-linejoin': strokeA['stroke-linejoin'] ?? 'round',
      })
  } else if (
    !geom.polygonPoints &&
    !geom.ellipse &&
    shape.softEdgePx &&
    shape.fill.kind === 'solid' &&
    !shape.fillOverlay &&
    box.w >= 1 &&
    box.h >= 1
  ) {
    const feathered = featheredShapeCanvas(
      shape.fill.color,
      box.w,
      box.h,
      shape.softEdgePx,
      geom.cornerRadiusPx ?? 0,
      shape.stroke ? { color: shape.stroke.color, widthPx: shape.stroke.widthPx } : undefined,
    )
    geometry += b.image(feathered.canvas, {
      x: -feathered.pad,
      y: -feathered.pad,
      width: box.w + 2 * feathered.pad,
      height: box.h + 2 * feathered.pad,
      filter,
    })
  } else {
    geometry += outlineElement(geom, box.w, box.h, { ...fillA, ...strokeA, filter })
  }

  // fillOverlay: white underlay + base + multiply overlay (isolated blend approximation)
  let overlayUnder = ''
  let overlayGeom = ''
  if (shape.fillOverlay && !shape.extrusion) {
    const oFill = b.fillAttrs(b.fillPaint(shape.fillOverlay, box.w, box.h))
    overlayUnder = outlineElement(geom, box.w, box.h, { fill: '#ffffff' })
    overlayGeom = outlineElement(geom, box.w, box.h, { ...oFill, style: 'mix-blend-mode:multiply' })
  }

  let shadowUnder = ''
  let shadowOver = ''
  if (isOverlayShadow(shape.shadow) && (!shape.extrusion || shape.extrusion.flat)) {
    const ov = shapeShadowOverlay(shape.shadow, shadowGeom(shape), box.w, box.h)
    if (ov) {
      const img = overlayImage(b, ov)
      if (ov.under) shadowUnder = img
      else shadowOver = img
    }
  }

  out.push(shadowUnder, overlayUnder, geometry, overlayGeom, shadowOver)

  const text = textBody(b, shape.text)
  if (text) {
    // PowerPoint flips geometry only: counter-flip the text layer like NodeBody
    const fh = !!box.flipH !== flip.h
    const fv = !!box.flipV !== flip.v
    if (fh || fv)
      out.push(
        `<g transform="translate(${n(fh ? box.w : 0)} ${n(fv ? box.h : 0)}) scale(${fh ? -1 : 1} ${fv ? -1 : 1})">${text}</g>`,
      )
    else out.push(text)
  }
  return out.join('')
}

/** Node content in node-local coordinates (the caller wraps the box transform). */
function nodeBody(b: PageBuilder, node: RenderNode, flip: FlipParity): string {
  const refl =
    node.type === 'shape' || node.type === 'text' || node.type === 'picture'
      ? (node as ShapeRenderNode | PictureRenderNode).reflection
      : undefined
  if (refl) {
    const clone = { ...node, reflection: undefined } as RenderNode
    return reflectionCopy(b, clone, refl, flip) + nodeBody(b, clone, flip)
  }
  switch (node.type) {
    case 'group': {
      const g = node as GroupRenderNode
      const childFlip = { h: !!g.box.flipH !== flip.h, v: !!g.box.flipV !== flip.v }
      return g.children.map((c) => staticNode(b, c, childFlip)).join('')
    }
    case 'picture':
      return emitPicture(b, node as PictureRenderNode)
    case 'table':
      return emitTable(b, node as TableRenderNode)
    case 'chart':
      return emitChart(b, node as ChartRenderNode)
    case 'placeholder-chip':
      return emitChip(b, node as ChipRenderNode)
    default:
      return emitShape(b, node as ShapeRenderNode, flip)
  }
}

/** Faded, optionally blurred mirror of the node below it (ReflectionCopy). */
function reflectionCopy(
  b: PageBuilder,
  node: RenderNode,
  refl: RenderReflection,
  flip: FlipParity,
): string {
  const { w, h } = node.box
  if (w < 1 || h < 1) return ''
  const pad = Math.ceil(refl.blurPx) + 2
  const gradId = b.id('rf')
  const maskId = b.id('rm')
  // kept at the touching edge (local y = h), fully erased past the fade extent
  b.def(
    el(
      'linearGradient',
      {
        id: gradId,
        gradientUnits: 'userSpaceOnUse',
        x1: 0,
        y1: h,
        x2: 0,
        y2: h * (1 - Math.max(refl.endPos, 0.02)),
      },
      el('stop', { offset: 0, 'stop-color': '#fff' }) +
        el('stop', { offset: 1, 'stop-color': '#000' }),
    ),
  )
  b.def(
    el(
      'mask',
      {
        id: maskId,
        maskUnits: 'userSpaceOnUse',
        x: -pad,
        y: -pad,
        width: w + 2 * pad,
        height: h + 2 * pad,
      },
      el('rect', {
        x: -pad,
        y: -pad,
        width: w + 2 * pad,
        height: h + 2 * pad,
        fill: `url(#${gradId})`,
      }),
    ),
  )
  const filter = refl.blurPx ? b.blurFilter(refl.blurPx, { x: 0, y: 0, w, h }) : undefined
  const body = nodeBody(b, node, flip)
  return `<g${attrs({
    transform: `translate(0 ${n(2 * h + refl.distPx)}) scale(1 -1)`,
    opacity: refl.startAlpha,
    mask: `url(#${maskId})`,
    filter,
  })}>${body}</g>`
}

function staticNode(b: PageBuilder, node: RenderNode, flip: FlipParity): string {
  return `<g transform="${boxTransform(node.box)}">${nodeBody(b, node, flip)}</g>`
}

/** Nodes the vector backend hands to the raster fallback up front. */
function needsRaster(node: RenderNode): boolean {
  if (node.type === 'group') return (node as GroupRenderNode).children.some(needsRaster)
  if (node.type === 'shape' || node.type === 'text') {
    const t = (node as ShapeRenderNode).text
    return !!t?.txWarp && t.lines.some((l) => l.runs.length > 0)
  }
  return false
}

function rasterElement(r: NodeRaster): string {
  return el('image', {
    href: r.href,
    x: r.x,
    y: r.y,
    width: r.w,
    height: r.h,
    preserveAspectRatio: 'none',
  })
}

export interface SvgPage {
  svg: string
  /** Every font-family stack the page's text references (for @font-face embedding). */
  fontFamilies: Set<string>
}

/**
 * One slide → SVG. Nodes that fail to vectorize (or need effects SVG cannot express)
 * are rasterized through `rasterizeNode` and placed as page-space images.
 */
export async function renderSlideToSvg(
  slide: RenderSlide,
  images: Map<string, HTMLImageElement>,
  opts: SvgExportOptions = {},
): Promise<SvgPage> {
  const rasters = new Map<string, NodeRaster>()
  const rasterize = async (node: RenderNode): Promise<void> => {
    if (rasters.has(node.id) || !opts.rasterizeNode) return
    const r = await opts.rasterizeNode(node)
    if (r) rasters.set(node.id, r)
  }
  for (const node of slide.nodes) if (needsRaster(node)) await rasterize(node)

  const build = (): { b: PageBuilder; body: string; failed: RenderNode[] } => {
    const b = new PageBuilder(images, rasters, opts.idPrefix ?? '')
    const failed: RenderNode[] = []
    const bg = b.fillPaint(slide.background, slide.widthPx, slide.heightPx)
    let body = el('rect', {
      width: slide.widthPx,
      height: slide.heightPx,
      ...(bg ? b.fillAttrs(bg) : { fill: '#ffffff' }),
    })
    for (const node of slide.nodes) {
      const r = rasters.get(node.id)
      if (r) {
        body += rasterElement(r)
        continue
      }
      try {
        body += staticNode(b, node, { h: false, v: false })
      } catch {
        failed.push(node)
      }
    }
    return { b, body, failed }
  }

  let result = build()
  if (result.failed.length && opts.rasterizeNode) {
    for (const node of result.failed) await rasterize(node)
    result = build()
  }
  const svg =
    `<svg xmlns="${SVG_NS}" xmlns:xlink="${XLINK}" viewBox="0 0 ${n(slide.widthPx)} ${n(slide.heightPx)}" width="${n(slide.widthPx)}" height="${n(slide.heightPx)}">` +
    result.b.defsMarkup() +
    result.body +
    '</svg>'
  return { svg, fontFamilies: result.b.fontFamilies }
}
