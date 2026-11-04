import type { InkInfo, NewInkImage } from '@genoffice/docx-engine'

/**
 * Ink annotations (freehand drawing strokes) — editor-side model.
 *
 * Strokes live on a transparent overlay above the page. Each stroke is
 * anchored to a body paragraph: its points are stored relative to that
 * paragraph's top-left corner at 100% zoom, which maps 1:1 onto the engine's
 * wp:anchor offsets (positionH from column, positionV from paragraph). At
 * save time every stroke is rasterized into a transparent PNG floating
 * picture; the vector points travel along in wp:docPr descr so reopening the
 * file in our editor restores live, erasable strokes.
 */

export type InkPenKind = 'pen' | 'highlighter' | 'pencil'
export type InkTool = 'select' | InkPenKind | 'eraser'

export const isPenKind = (tool: InkTool): tool is InkPenKind =>
  tool === 'pen' || tool === 'highlighter' || tool === 'pencil'

export interface InkPoint {
  x: number
  y: number
}

export interface InkStroke {
  tool: InkPenKind
  /** hex without '#' */
  color: string
  /** stroke width in px at 100% zoom */
  width: number
  /** px relative to the anchor paragraph's top-left, 100% zoom */
  points: InkPoint[]
}

/**
 * One annotation on the overlay. Editor-drawn annotations carry `stroke`
 * (editable vectors); annotations restored from a document without a
 * readable payload carry `image` instead (movable/erasable as a whole).
 */
export interface InkAnnotation {
  id: string
  /** docxIndex of the anchor paragraph in the current parsed doc */
  anchorIndex: number
  stroke?: InkStroke
  image?: {
    dataUrl: string
    offsetXPx: number
    offsetYPx: number
    widthPx: number
    heightPx: number
  }
}

let inkIdSeq = 0
export const nextInkId = () => `ink${Date.now()}-${inkIdSeq++}`

// ---- payload (wp:docPr descr) round-trip ----

const PAYLOAD_VERSION = 1

export function encodeInkPayload(stroke: InkStroke): string {
  return JSON.stringify({
    v: PAYLOAD_VERSION,
    tool: stroke.tool,
    color: stroke.color,
    width: stroke.width,
    // round to 0.1px: descr must stay reasonably small on long strokes
    points: stroke.points.map((p) => [Math.round(p.x * 10) / 10, Math.round(p.y * 10) / 10]),
  })
}

export function decodeInkPayload(payload: string): InkStroke | null {
  try {
    const data = JSON.parse(payload) as {
      v?: number
      tool?: string
      color?: string
      width?: number
      points?: [number, number][]
    }
    if (data.v !== PAYLOAD_VERSION || !Array.isArray(data.points) || data.points.length === 0) {
      return null
    }
    return {
      tool: data.tool === 'highlighter' || data.tool === 'pencil' ? data.tool : 'pen',
      color: typeof data.color === 'string' ? data.color : '000000',
      width: typeof data.width === 'number' && data.width > 0 ? data.width : 2,
      points: data.points.map(([x, y]) => ({ x, y })),
    }
  } catch {
    return null
  }
}

/** Restore overlay annotations from a freshly parsed document. */
export function annotationsFromParsed(inks: InkInfo[]): InkAnnotation[] {
  return inks.map((ink) => {
    const stroke = ink.payload ? decodeInkPayload(ink.payload) : null
    if (stroke) return { id: nextInkId(), anchorIndex: ink.anchorIndex, stroke }
    return {
      id: nextInkId(),
      anchorIndex: ink.anchorIndex,
      image: {
        dataUrl: ink.dataUrl ?? '',
        offsetXPx: ink.offsetXPx,
        offsetYPx: ink.offsetYPx,
        widthPx: ink.widthPx,
        heightPx: ink.heightPx,
      },
    }
  })
}

// ---- geometry ----

export interface InkBounds {
  x: number
  y: number
  w: number
  h: number
}

/** bounding box of every rendered layer including line width (+1px antialias margin) */
export function strokeBounds(stroke: InkStroke): InkBounds {
  let minX = Infinity
  let minY = Infinity
  let maxX = -Infinity
  let maxY = -Infinity
  for (const layer of strokeLayers(stroke)) {
    const pad = layer.width / 2 + 1
    for (const p of layer.points) {
      if (p.x - pad < minX) minX = p.x - pad
      if (p.y - pad < minY) minY = p.y - pad
      if (p.x + pad > maxX) maxX = p.x + pad
      if (p.y + pad > maxY) maxY = p.y + pad
    }
  }
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY }
}

/** distance from a point to the segment ab */
function segmentDistance(p: InkPoint, a: InkPoint, b: InkPoint): number {
  const dx = b.x - a.x
  const dy = b.y - a.y
  const lenSq = dx * dx + dy * dy
  const t =
    lenSq === 0 ? 0 : Math.max(0, Math.min(1, ((p.x - a.x) * dx + (p.y - a.y) * dy) / lenSq))
  const cx = a.x + t * dx
  const cy = a.y + t * dy
  return Math.hypot(p.x - cx, p.y - cy)
}

/** true when the point is within `threshold` of any segment of the stroke */
export function strokeHitTest(stroke: InkStroke, point: InkPoint, threshold: number): boolean {
  const reach = threshold + stroke.width / 2
  const pts = stroke.points
  if (pts.length === 1) return Math.hypot(point.x - pts[0].x, point.y - pts[0].y) <= reach
  for (let i = 1; i < pts.length; i++) {
    if (segmentDistance(point, pts[i - 1], pts[i]) <= reach) return true
  }
  return false
}

/** SVG path data through the stroke points (simple polyline) */
export function strokePathD(points: InkPoint[]): string {
  if (points.length === 0) return ''
  if (points.length === 1) {
    // dot: tiny segment so stroke-linecap can render it
    return `M ${points[0].x} ${points[0].y} L ${points[0].x + 0.01} ${points[0].y}`
  }
  return points.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x} ${p.y}`).join(' ')
}

export const HIGHLIGHTER_ALPHA = 0.45

/** one polyline of a rendered stroke; pencil strokes fan out into several */
export interface InkLayer {
  points: InkPoint[]
  width: number
  alpha: number
}

/**
 * Jitter seed from the first point only: a live stroke grows by appending, so
 * the grain already drawn must not move when more points arrive.
 */
function jitterSeed(points: InkPoint[], salt: number): number {
  const p = points[0]
  return (Math.round(p.x * 10) * 31 + Math.round(p.y * 10) * 7 + salt * 1013) | 0
}

function mulberry(seed: number): () => number {
  let a = seed >>> 0
  return () => {
    a = (a + 0x6d2b79f5) | 0
    let t = Math.imul(a ^ (a >>> 15), 1 | a)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

/**
 * Pencil = a soft core line plus two thin grainy side lines wobbling across the
 * path; the sum reads as graphite instead of a vector pen.
 */
export function strokeLayers(stroke: InkStroke): InkLayer[] {
  if (stroke.tool === 'highlighter')
    return [{ points: stroke.points, width: stroke.width, alpha: HIGHLIGHTER_ALPHA }]
  if (stroke.tool === 'pen') return [{ points: stroke.points, width: stroke.width, alpha: 1 }]
  const spread = stroke.width * 0.45
  const side = (bias: number): InkPoint[] => {
    const rand = mulberry(jitterSeed(stroke.points, bias))
    return stroke.points.map((p, i) => {
      const prev = stroke.points[Math.max(0, i - 1)]
      const next = stroke.points[Math.min(stroke.points.length - 1, i + 1)]
      const len = Math.hypot(next.x - prev.x, next.y - prev.y) || 1
      const nx = -(next.y - prev.y) / len
      const ny = (next.x - prev.x) / len
      const d = (bias + (rand() - 0.5) * 0.8) * spread
      return { x: p.x + nx * d, y: p.y + ny * d }
    })
  }
  return [
    { points: stroke.points, width: stroke.width * 0.75, alpha: 0.7 },
    { points: side(1), width: stroke.width * 0.3, alpha: 0.4 },
    { points: side(-1), width: stroke.width * 0.3, alpha: 0.4 },
  ]
}

// ---- save-time rasterization ----

/** raster output of one annotation: transparent PNG + placement */
export interface InkRaster {
  base64: string
  widthPx: number
  heightPx: number
  offsetXPx: number
  offsetYPx: number
}

/** default renderer: draw the stroke on a transparent canvas at 2x scale */
export function rasterizeStroke(stroke: InkStroke): InkRaster {
  const bounds = strokeBounds(stroke)
  const scale = 2
  const canvas = document.createElement('canvas')
  canvas.width = Math.max(1, Math.ceil(bounds.w * scale))
  canvas.height = Math.max(1, Math.ceil(bounds.h * scale))
  const ctx = canvas.getContext('2d')!
  ctx.scale(scale, scale)
  ctx.translate(-bounds.x, -bounds.y)
  ctx.strokeStyle = `#${stroke.color}`
  ctx.lineCap = 'round'
  ctx.lineJoin = 'round'
  for (const layer of strokeLayers(stroke)) {
    ctx.lineWidth = layer.width
    ctx.globalAlpha = layer.alpha
    ctx.beginPath()
    const pts = layer.points
    ctx.moveTo(pts[0].x, pts[0].y)
    if (pts.length === 1) ctx.lineTo(pts[0].x + 0.01, pts[0].y)
    for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i].x, pts[i].y)
    ctx.stroke()
  }
  return {
    base64: canvas.toDataURL('image/png').split(',')[1],
    widthPx: Math.max(1, Math.ceil(bounds.w)),
    heightPx: Math.max(1, Math.ceil(bounds.h)),
    offsetXPx: bounds.x,
    offsetYPx: bounds.y,
  }
}

/**
 * Build the engine's ink list for a save. Annotations whose anchor paragraph
 * no longer exists in the save plan are dropped (their paragraph was deleted).
 */
export function buildInkImages(
  annotations: InkAnnotation[],
  saveBlockIndexByDocx: Map<number, number>,
  rasterize: (stroke: InkStroke) => InkRaster = rasterizeStroke,
): NewInkImage[] {
  const out: NewInkImage[] = []
  for (const ann of annotations) {
    const blockIndex = saveBlockIndexByDocx.get(ann.anchorIndex)
    if (blockIndex === undefined) continue
    if (ann.stroke) {
      const raster = rasterize(ann.stroke)
      out.push({ blockIndex, ...raster, payload: encodeInkPayload(ann.stroke) })
    } else if (ann.image && ann.image.dataUrl.startsWith('data:image/png;base64,')) {
      // restored foreign/opaque annotation: carry the original PNG through
      out.push({
        blockIndex,
        base64: ann.image.dataUrl.slice('data:image/png;base64,'.length),
        widthPx: ann.image.widthPx,
        heightPx: ann.image.heightPx,
        offsetXPx: ann.image.offsetXPx,
        offsetYPx: ann.image.offsetYPx,
      })
    }
  }
  return out
}
