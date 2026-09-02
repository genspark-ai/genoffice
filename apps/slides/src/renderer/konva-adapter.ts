/**
 * Thin Konva adapter — converts pptx-render's RenderTree (pure data) into primitives Konva can draw.
 *
 * All fidelity logic lives in pptx-render (coordinates/metrics/line breaking); this layer only
 * does mechanical "data → Konva attributes" mapping with no layout decisions. Rendering
 * correctness is therefore guaranteed by pptx-render unit tests, and this layer is thin enough
 * to need almost no testing.
 */
import type {
  RenderGlow,
  RenderNode,
  RenderFill,
  RenderStroke,
  RenderShadow,
  RenderTextLayout,
  ShapeRenderNode,
  PictureRenderNode,
  GlyphRun,
} from '@genoffice/pptx-render'
import { patternGrid } from '@genoffice/pptx-render'
import { classifyCjkScript } from '../shared/cjk-script'

/**
 * Konva container props for a placed box: rotation and flip pivot on the box CENTER
 * (OOXML semantics — a:xfrm off is the unrotated top-left, rot/flip apply about the center).
 * The node's reported position is the box center; model x/y = position − w/2, h/2.
 */
export function boxPivotProps(box: {
  x: number
  y: number
  w: number
  h: number
  rotationDeg: number
  flipH?: boolean
  flipV?: boolean
}) {
  return {
    x: box.x + box.w / 2,
    y: box.y + box.h / 2,
    offsetX: box.w / 2,
    offsetY: box.h / 2,
    rotation: box.rotationDeg,
    scaleX: box.flipH ? -1 : 1,
    scaleY: box.flipV ? -1 : 1,
  }
}

/** Konva fill attributes (for Rect/Path/Text). */
export interface KonvaFillProps {
  fill?: string
  fillLinearGradientStartPoint?: { x: number; y: number }
  fillLinearGradientEndPoint?: { x: number; y: number }
  fillLinearGradientColorStops?: Array<number | string>
  fillRadialGradientStartPoint?: { x: number; y: number }
  fillRadialGradientEndPoint?: { x: number; y: number }
  fillRadialGradientStartRadius?: number
  fillRadialGradientEndRadius?: number
  fillRadialGradientColorStops?: Array<number | string>
  fillPatternImage?: HTMLImageElement
  fillPatternRepeat?: string
  fillPatternScaleX?: number
  fillPatternScaleY?: number
  fillPatternX?: number
  fillPatternY?: number
  /** Translucent picture fill (alphaModFix) */
  opacity?: number
}

/**
 * Shift fill coordinates for Konva shapes whose local origin is the CENTER (Ellipse):
 * fillToKonva computes gradient points / pattern anchors in box-local top-left space,
 * so centered shapes would otherwise show the pattern wrapping around the middle and
 * only the first half of a gradient.
 */
export function centerFillProps(props: KonvaFillProps, w: number, h: number): KonvaFillProps {
  const shift = (p: { x: number; y: number }) => ({ x: p.x - w / 2, y: p.y - h / 2 })
  return {
    ...props,
    ...(props.fillLinearGradientStartPoint
      ? { fillLinearGradientStartPoint: shift(props.fillLinearGradientStartPoint) }
      : {}),
    ...(props.fillLinearGradientEndPoint
      ? { fillLinearGradientEndPoint: shift(props.fillLinearGradientEndPoint) }
      : {}),
    ...(props.fillRadialGradientStartPoint
      ? { fillRadialGradientStartPoint: shift(props.fillRadialGradientStartPoint) }
      : {}),
    ...(props.fillRadialGradientEndPoint
      ? { fillRadialGradientEndPoint: shift(props.fillRadialGradientEndPoint) }
      : {}),
    // Shift existing pattern anchors (e.g. stretch fillRect insets) rather than replacing them
    ...(props.fillPatternImage
      ? {
          fillPatternX: (props.fillPatternX ?? 0) - w / 2,
          fillPatternY: (props.fillPatternY ?? 0) - h / 2,
        }
      : {}),
  }
}

/** Collect the image dataUrl referenced by a fill (for preloading upstream). */
export function fillImageUrl(fill: RenderFill): string | undefined {
  return fill.kind === 'image' ? fill.dataUrl : undefined
}

// ── Gradient ramp interpolation ──────────────────────────────────────────────
// PowerPoint blends gradient ramps in linear sRGB (measured on tdf105739's
// FF0000→00B050 background: the midpoint renders (188,129,55), matching the
// linear-light mix; canvas gradients blend raw sRGB, which gives a muddy
// (128,88,40)). Subdivide each stop pair with linear-mixed intermediate stops
// so the canvas ramp tracks PowerPoint's.
const srgbToLin = (v: number) => (v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4)
const linToSrgb = (v: number) => (v <= 0.0031308 ? v * 12.92 : 1.055 * v ** (1 / 2.4) - 0.055)

function parseStopColor(c: string): [number, number, number, number] | null {
  const n = normalizeColor(c)
  const hex = /^#([0-9A-Fa-f]{6})$/.exec(n)
  if (hex) {
    const h = hex[1]!
    return [
      parseInt(h.slice(0, 2), 16),
      parseInt(h.slice(2, 4), 16),
      parseInt(h.slice(4, 6), 16),
      1,
    ]
  }
  const rgba = /^rgba\((\d+),(\d+),(\d+),([0-9.]+)\)$/.exec(n)
  if (rgba) return [Number(rgba[1]), Number(rgba[2]), Number(rgba[3]), Number(rgba[4])]
  return null
}

const SUBDIVISIONS = 8

/**
 * PowerPoint's color ramp ignores the color of fully-transparent stops: their RGB is
 * replaced by the visible-stop ramp sampled at their position (clamped at the ends), so
 * white@0% → navy → navy fades stay navy. Only when fewer than two visible stops remain
 * does the declared color survive and the fade washes toward it (navy → white@0% probes
 * measured both behaviors over black and white backdrops).
 */
function maskTransparentStopColors(
  sorted: Array<{ pos: number; color: string }>,
): Array<{ pos: number; color: string }> {
  const parsed = sorted.map((s) => parseStopColor(s.color))
  const visible: Array<{ pos: number; rgb: [number, number, number] }> = []
  for (let i = 0; i < sorted.length; i++) {
    const p = parsed[i]
    if (p && p[3] > 0) visible.push({ pos: sorted[i]!.pos, rgb: [p[0], p[1], p[2]] })
  }
  if (visible.length < 2 || visible.length === sorted.length) return sorted
  const sample = (pos: number): [number, number, number] => {
    if (pos <= visible[0]!.pos) return visible[0]!.rgb
    const last = visible[visible.length - 1]!
    if (pos >= last.pos) return last.rgb
    let j = 1
    while (visible[j]!.pos < pos) j++
    const lo = visible[j - 1]!
    const hi = visible[j]!
    const t = hi.pos === lo.pos ? 0 : (pos - lo.pos) / (hi.pos - lo.pos)
    const ch = (i: number) =>
      Math.round(
        linToSrgb(
          srgbToLin(lo.rgb[i]! / 255) +
            (srgbToLin(hi.rgb[i]! / 255) - srgbToLin(lo.rgb[i]! / 255)) * t,
        ) * 255,
      )
    return [ch(0), ch(1), ch(2)]
  }
  return sorted.map((s, i) => {
    const p = parsed[i]
    if (!p || p[3] > 0) return s
    const [r, g, b] = sample(s.pos)
    return { pos: s.pos, color: `rgba(${r},${g},${b},0)` }
  })
}

/** Konva colorStops array with linear-sRGB interpolated midpoints between each stop pair. */
function linearRampStops(stops: Array<{ pos: number; color: string }>): Array<number | string> {
  const sorted = maskTransparentStopColors([...stops].sort((a, b) => a.pos - b.pos))
  const out: Array<number | string> = []
  for (let i = 0; i < sorted.length; i++) {
    const cur = sorted[i]!
    out.push(cur.pos, normalizeColor(cur.color))
    const next = sorted[i + 1]
    if (!next || next.color === cur.color || next.pos - cur.pos < 0.02) continue
    const a = parseStopColor(cur.color)
    const b = parseStopColor(next.color)
    if (!a || !b) continue
    const [lr, lg, lb] = [srgbToLin(a[0] / 255), srgbToLin(a[1] / 255), srgbToLin(a[2] / 255)]
    const [mr, mg, mb] = [srgbToLin(b[0] / 255), srgbToLin(b[1] / 255), srgbToLin(b[2] / 255)]
    // Alpha-varying pairs interpolate straight (non-premultiplied): PowerPoint ramps the
    // color channels toward the transparent stop's hue and the alpha linearly — measured
    // on a controlled probe (navy→white@0 both stop orders, over black and white backdrops).
    for (let k = 1; k < SUBDIVISIONS; k++) {
      const t = k / SUBDIVISIONS
      const al = a[3] + (b[3] - a[3]) * t
      const r = Math.round(linToSrgb(lr + (mr - lr) * t) * 255)
      const g = Math.round(linToSrgb(lg + (mg - lg) * t) * 255)
      const bl = Math.round(linToSrgb(lb + (mb - lb) * t) * 255)
      out.push(
        cur.pos + (next.pos - cur.pos) * t,
        al >= 1 ? `rgb(${r},${g},${bl})` : `rgba(${r},${g},${bl},${al.toFixed(3)})`,
      )
    }
  }
  return out
}

/**
 * Unit direction of an a:lin gradient. scaled="1" declares the angle in a square
 * fill box that stretches with the shape, so the rendered direction is the angle
 * unsquashed by the aspect ratio: tanθ' = tanθ × (w/h). Measured on prod_048's
 * 45°-scaled full-width band: the ramp midline runs top-left → bottom-right through
 * the near-vertical direction (h·cosθ, w·sinθ), pixel-matching PowerPoint's export;
 * the untransformed 45° (and the diagonal direction (w·cosθ, h·sinθ)) both miss it.
 */
function linearGradientDirection(
  angleDeg: number,
  scaled: boolean | undefined,
  w: number,
  h: number,
): [number, number] {
  const rad = (angleDeg * Math.PI) / 180
  if (!scaled) return [Math.cos(rad), Math.sin(rad)]
  const vx = h * Math.cos(rad)
  const vy = w * Math.sin(rad)
  const n = Math.hypot(vx, vy) || 1
  return [vx / n, vy / n]
}

/** Cached shape-sized pattern canvases for rect/shape path gradients. */
const pathGradCache = new Map<string, HTMLCanvasElement>()

/**
 * rect/shape path gradient, rendered per pixel through a 256-entry ramp LUT (band/ring
 * painting shows AA seams; the metrics below are C1-discontinuous on the diagonals,
 * which per-pixel sampling reproduces exactly like PowerPoint).
 * - rect: t = max(|dx|/ex, |dy|/ey) with per-side extents to the fillToRect focus
 *   (seams run through the focus diagonals; iso-lines converge on the focus point).
 * - shape: t = 1 − inset/maxInset, uniform distance to the bounds (45° corner
 *   seams, medial plateau on non-square bounds; the focus is ignored, matching PPT's
 *   path type having no direction option).
 */
export function pathGradientCanvas(
  kind: 'rect' | 'shape',
  stops: Array<{ pos: number; color: string }>,
  w: number,
  h: number,
  cx01: number,
  cy01: number,
): HTMLCanvasElement | null {
  const cw = Math.max(1, Math.ceil(w))
  const ch = Math.max(1, Math.ceil(h))
  const key = `${kind}:${cw}x${ch}:${cx01}:${cy01}:${stops.map((s) => `${s.pos},${s.color}`).join(';')}`
  const hit = pathGradCache.get(key)
  if (hit) return hit

  // 1x256 ramp strip: reuse linearRampStops' linear-sRGB interpolation via a canvas gradient
  const strip = document.createElement('canvas')
  strip.width = 256
  strip.height = 1
  const sctx = strip.getContext('2d', { willReadFrequently: true })
  const cv = document.createElement('canvas')
  cv.width = cw
  cv.height = ch
  const ctx = cv.getContext('2d')
  if (!sctx || !ctx) return null
  const lg = sctx.createLinearGradient(0, 0, 256, 0)
  const ramp = linearRampStops(stops)
  for (let i = 0; i < ramp.length; i += 2)
    lg.addColorStop(Math.max(0, Math.min(1, ramp[i] as number)), ramp[i + 1] as string)
  sctx.fillStyle = lg
  sctx.fillRect(0, 0, 256, 1)
  const px = sctx.getImageData(0, 0, 256, 1).data

  const cx = cx01 * cw
  const cy = cy01 * ch
  const maxInset = Math.min(cw, ch) / 2
  const img = ctx.createImageData(cw, ch)
  const d = img.data
  for (let y = 0; y < ch; y++) {
    for (let x = 0; x < cw; x++) {
      let t: number
      if (kind === 'shape') {
        const inset = Math.min(x + 0.5, y + 0.5, cw - x - 0.5, ch - y - 0.5)
        t = 1 - Math.min(1, inset / maxInset)
      } else {
        const ex = x + 0.5 < cx ? cx : cw - cx
        const ey = y + 0.5 < cy ? cy : ch - cy
        const tx = ex > 0 ? Math.abs(x + 0.5 - cx) / ex : 1
        const ty = ey > 0 ? Math.abs(y + 0.5 - cy) / ey : 1
        t = Math.min(1, Math.max(tx, ty))
      }
      const j = Math.min(255, Math.round(t * 255)) * 4
      const o = (y * cw + x) * 4
      d[o] = px[j]!
      d[o + 1] = px[j + 1]!
      d[o + 2] = px[j + 2]!
      d[o + 3] = px[j + 3]!
    }
  }
  ctx.putImageData(img, 0, 0)
  if (pathGradCache.size > 64) pathGradCache.delete(pathGradCache.keys().next().value!)
  pathGradCache.set(key, cv)
  return cv
}

export function fillToKonva(
  fill: RenderFill,
  w: number,
  h: number,
  images?: Map<string, HTMLImageElement>,
  // Pattern phase: the shape's page position (GDI/PDF hatches anchor at the page
  // origin, so the cell grid must not restart at every shape's own origin)
  phase?: { x: number; y: number },
): KonvaFillProps {
  switch (fill.kind) {
    case 'solid':
      return { fill: normalizeColor(fill.color) }
    case 'pattern':
      return {
        fillPatternImage: patternCanvas(fill, w, h, phase) as HTMLImageElement,
        fillPatternRepeat: 'no-repeat',
      }
    case 'gradient': {
      if (fill.radial) {
        // rect / shape: concentric-rectangle band pattern toward the focus point
        if (fill.path === 'rect' || fill.path === 'shape') {
          const cv = pathGradientCanvas(
            fill.path,
            fill.stops,
            w,
            h,
            fill.center?.x ?? 0.5,
            fill.center?.y ?? 0.5,
          )
          if (cv)
            return {
              fillPatternImage: cv as unknown as HTMLImageElement,
              fillPatternRepeat: 'no-repeat',
            }
        }
        // circle: native radial. Center follows fillToRect; the 100% ring sits on
        // the farthest corner (pos-1 color lands exactly in that corner, like PowerPoint).
        const cx = (fill.center?.x ?? 0.5) * w
        const cy = (fill.center?.y ?? 0.5) * h
        const far = Math.max(
          Math.hypot(cx, cy),
          Math.hypot(w - cx, cy),
          Math.hypot(cx, h - cy),
          Math.hypot(w - cx, h - cy),
        )
        return {
          fillRadialGradientStartPoint: { x: cx, y: cy },
          fillRadialGradientEndPoint: { x: cx, y: cy },
          fillRadialGradientStartRadius: 0,
          fillRadialGradientEndRadius: far * 1.0,
          fillRadialGradientColorStops: linearRampStops(fill.stops),
        }
      }
      const [dx, dy] = linearGradientDirection(fill.angleDeg, fill.scaled, w, h)
      // Gradient vector length = the box's projection onto the gradient direction,
      // so the ramp spans exactly the shape (PowerPoint semantics for a:lin)
      const cx = w / 2
      const cy = h / 2
      const len = Math.abs(dx) * w + Math.abs(dy) * h
      return {
        fillLinearGradientStartPoint: { x: cx - (dx * len) / 2, y: cy - (dy * len) / 2 },
        fillLinearGradientEndPoint: { x: cx + (dx * len) / 2, y: cy + (dy * len) / 2 },
        fillLinearGradientColorStops: linearRampStops(fill.stops),
      }
    }
    case 'image': {
      const img = fill.dataUrl ? images?.get(fill.dataUrl) : undefined
      if (img) {
        // Konva accepts any CanvasImageSource at runtime; its typings only admit HTMLImageElement
        const src = processedImage(
          img,
          fill.dataUrl ?? '',
          fill.clrChange,
          fill.duotone,
          fill.lum,
        ) as HTMLImageElement
        // recolored variants must not share cache slots with the raw image
        const srcKey = processedImageKey(fill.dataUrl ?? '', fill.clrChange, fill.duotone, fill.lum)
        if (fill.mode === 'tile') {
          // PowerPoint tiles at the image's 144dpi natural size x sx/sy, anchored per algn
          // plus tx/ty offsets. Pre-composited into a shape-sized canvas: Konva pattern
          // transforms (scale/offset with repeat) hit the same Skia pixelRatio bug as
          // no-repeat tiles (#612), so only the pre-padded canvas + no-repeat combo is safe.
          const t = fill.tile
          if (t) {
            // 1:1 shape-sized canvas with no pattern transform at all — any pattern
            // scale/offset (even ~1.0) trips the Skia pixelRatio bug (#612)
            const tileCv = anchoredTileCanvas(src, srcKey, w, h, t)
            return {
              fillPatternImage: tileCv as HTMLImageElement,
              fillPatternRepeat: 'no-repeat',
              ...(fill.alpha != null ? { opacity: fill.alpha } : {}),
            }
          }
          return {
            fillPatternImage: src,
            ...(fill.alpha != null ? { opacity: fill.alpha } : {}),
          }
        }
        // Degenerate stretch textures (≤2×2 px): PowerPoint rasterizes these as one flat
        // color (verified against its PDF export), while bilinear stretching would smear
        // a gradient across the shape.
        if (isDegenerateImage(img)) {
          return {
            fill: averageColor(src, srcKey),
            ...(fill.alpha != null ? { opacity: fill.alpha } : {}),
          }
        }
        // stretch fillRect: the image maps into an inset subrect of the shape. Composited into
        // a transparent-padded tile covering the whole shape instead of fillPatternX/no-repeat:
        // Konva's pattern transform at pixelRatio > 1 makes Skia paint the outside of a
        // no-repeat tile opaque black instead of leaving it transparent.
        const fr = fill.fillRect
        const tile = fr ? insetFillTile(src, srcKey, fr) : src
        return {
          fillPatternImage: tile as HTMLImageElement,
          fillPatternRepeat: 'no-repeat',
          fillPatternScaleX: w / (tile.width || w),
          fillPatternScaleY: h / (tile.height || h),
          // alphaModFix: fades the whole node — picture fills with a visible stroke are rare
          ...(fill.alpha != null ? { opacity: fill.alpha } : {}),
        }
      }
      return {}
    }
    case 'none':
    default:
      return {}
  }
}

export function strokeToKonva(
  stroke: RenderStroke | undefined,
  // Shape-local box for gradient strokes (the gradient vector spans the shape, like fills)
  size?: { w: number; h: number },
): {
  stroke?: string
  strokeWidth?: number
  dash?: number[]
  lineCap?: 'butt' | 'round' | 'square'
  lineJoin?: 'round' | 'bevel' | 'miter'
  strokeLinearGradientStartPoint?: { x: number; y: number }
  strokeLinearGradientEndPoint?: { x: number; y: number }
  strokeLinearGradientColorStops?: Array<number | string>
} {
  if (!stroke) return {}
  const grad = stroke.gradient
  let gradProps = {}
  if (grad && size) {
    const [dx, dy] = linearGradientDirection(grad.angleDeg, grad.scaled, size.w, size.h)
    const cx = size.w / 2
    const cy = size.h / 2
    const len = Math.abs(dx) * size.w + Math.abs(dy) * size.h
    gradProps = {
      strokeLinearGradientStartPoint: { x: cx - (dx * len) / 2, y: cy - (dy * len) / 2 },
      strokeLinearGradientEndPoint: { x: cx + (dx * len) / 2, y: cy + (dy * len) / 2 },
      strokeLinearGradientColorStops: linearRampStops(grad.stops),
    }
  }
  return {
    stroke: normalizeColor(stroke.color),
    strokeWidth: stroke.widthPx,
    ...(stroke.dash ? { dash: stroke.dash } : {}),
    ...(stroke.cap ? { lineCap: stroke.cap } : {}),
    ...(stroke.join ? { lineJoin: stroke.join } : {}),
    ...gradProps,
  }
}

/** Inner/perspective shadows can't be expressed as canvas shadow props — they draw as an offscreen overlay instead. */
export function isOverlayShadow(s: RenderShadow | undefined): boolean {
  return !!s && (!!s.inner || s.scaleX != null || s.scaleY != null || !!s.skewXDeg || !!s.skewYDeg)
}

/** Outer shadow → Konva shadow attributes; without a shadow, glow is approximated with a zero-offset shadow.
 * Overlay shadows (inner/perspective) are excluded here — see shapeShadowOverlay. */
export function shadowToKonva(
  rawShadow: RenderShadow | undefined,
  glow?: RenderGlow,
): {
  shadowColor?: string
  shadowBlur?: number
  shadowOffsetX?: number
  shadowOffsetY?: number
  shadowEnabled?: boolean
} {
  const shadow = isOverlayShadow(rawShadow) ? undefined : rawShadow
  if (!shadow && glow) {
    return {
      shadowColor: normalizeColor(glow.color),
      shadowBlur: glow.blurPx,
      shadowOffsetX: 0,
      shadowOffsetY: 0,
      shadowEnabled: true,
    }
  }
  if (!shadow) return {}
  return {
    shadowColor: normalizeColor(shadow.color),
    shadowBlur: shadow.blurPx,
    shadowOffsetX: shadow.offsetX,
    shadowOffsetY: shadow.offsetY,
    shadowEnabled: true,
  }
}

// ── Overlay shadows (inner / perspective) ───────────────────────────────
// Canvas shadow props can only cast a blurred copy outward; inner shadows and the
// skewed/scaled perspective silhouettes are rendered into an offscreen canvas that
// NodeBody stacks as a Konva.Image around the geometry.

/** Shape geometry in local box px, expressible as a Path2D. */
export type ShadowGeom =
  | { kind: 'rect'; w: number; h: number; cornerRadius?: number }
  | { kind: 'ellipse'; w: number; h: number }
  | { kind: 'polygon'; points: number[] }
  | { kind: 'path'; data: string }

function geomPath2D(g: ShadowGeom): Path2D {
  const p = new Path2D()
  if (g.kind === 'rect') {
    if (g.cornerRadius) p.roundRect(0, 0, g.w, g.h, g.cornerRadius)
    else p.rect(0, 0, g.w, g.h)
  } else if (g.kind === 'ellipse') {
    p.ellipse(g.w / 2, g.h / 2, g.w / 2, g.h / 2, 0, 0, Math.PI * 2)
  } else if (g.kind === 'polygon') {
    for (let i = 0; i + 1 < g.points.length; i += 2) {
      if (i === 0) p.moveTo(g.points[0]!, g.points[1]!)
      else p.lineTo(g.points[i]!, g.points[i + 1]!)
    }
    p.closePath()
  } else {
    return new Path2D(g.data)
  }
  return p
}

export interface ShadowOverlay {
  canvas: HTMLCanvasElement
  /** Placement relative to the shape box's top-left (local px) */
  x: number
  y: number
  /** Local px covered by the canvas (draw size for the Konva.Image) */
  w: number
  h: number
  /** True = paint under the geometry (perspective silhouette); false = over it (inner shadow) */
  under: boolean
}

const shadowOverlayCache = new Map<string, ShadowOverlay | null>()

/** algn anchor → local point of the shape box the perspective transform pivots on ('b' is the OOXML default). */
function algnPoint(algn: string | undefined, w: number, h: number): [number, number] {
  const a = algn ?? 'b'
  const x = a.includes('l') ? 0 : a.includes('r') ? w : w / 2
  const y = a.includes('t') ? 0 : a.includes('b') ? h : h / 2
  return [x, y]
}

/**
 * Build the overlay canvas for an inner or perspective shadow.
 * Inner: clip to the geometry and fill its INVERSE (evenodd) offset by the shadow
 * distance — only the cast shadow lands inside the clip, giving a true inset shadow.
 * Perspective: the geometry silhouette filled with the shadow color, run through the
 * outerShdw sx/sy/kx/ky transform around the algn anchor and gaussian-blurred.
 * Returns null when the shadow needs no overlay or no 2D context exists (jsdom).
 */
export function shapeShadowOverlay(
  shadow: RenderShadow | undefined,
  geom: ShadowGeom,
  boxW: number,
  boxH: number,
): ShadowOverlay | null {
  if (!shadow || !isOverlayShadow(shadow) || boxW < 1 || boxH < 1) return null
  const key = JSON.stringify([shadow, geom, Math.round(boxW * 4), Math.round(boxH * 4)])
  const hit = shadowOverlayCache.get(key)
  if (hit !== undefined) return hit
  if (shadowOverlayCache.size > 200) shadowOverlayCache.clear()
  const dpr = Math.min(globalThis.devicePixelRatio || 1, 2)
  const build = (): ShadowOverlay | null => {
    const path = geomPath2D(geom)
    if (shadow.inner) {
      const canvas = document.createElement('canvas')
      canvas.width = Math.max(1, Math.ceil(boxW * dpr))
      canvas.height = Math.max(1, Math.ceil(boxH * dpr))
      const ctx = canvas.getContext('2d')
      if (!ctx) return null
      ctx.scale(dpr, dpr)
      ctx.clip(path)
      // canvas shadow params live in device space (transforms don't apply to them)
      ctx.shadowColor = normalizeColor(shadow.color)
      ctx.shadowBlur = shadow.blurPx * dpr
      ctx.shadowOffsetX = shadow.offsetX * dpr
      ctx.shadowOffsetY = shadow.offsetY * dpr
      const pad = shadow.blurPx + Math.abs(shadow.offsetX) + Math.abs(shadow.offsetY) + 4
      const inv = new Path2D()
      inv.rect(-pad, -pad, boxW + 2 * pad, boxH + 2 * pad)
      inv.addPath(path)
      ctx.fillStyle = '#000'
      ctx.fill(inv, 'evenodd')
      return { canvas, x: 0, y: 0, w: boxW, h: boxH, under: false }
    }
    // Perspective silhouette: generous padding for skew/flip/blur/offset overhang
    const pad =
      shadow.blurPx + Math.abs(shadow.offsetX) + Math.abs(shadow.offsetY) + Math.max(boxW, boxH)
    const canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.ceil((boxW + 2 * pad) * dpr))
    canvas.height = Math.max(1, Math.ceil((boxH + 2 * pad) * dpr))
    const ctx = canvas.getContext('2d')
    if (!ctx) return null
    ctx.scale(dpr, dpr)
    if (shadow.blurPx) ctx.filter = `blur(${shadow.blurPx / 2}px)`
    const [ax, ay] = algnPoint(shadow.algn, boxW, boxH)
    ctx.translate(pad + shadow.offsetX + ax, pad + shadow.offsetY + ay)
    ctx.transform(
      shadow.scaleX ?? 1,
      Math.tan(((shadow.skewYDeg ?? 0) * Math.PI) / 180),
      Math.tan(((shadow.skewXDeg ?? 0) * Math.PI) / 180),
      shadow.scaleY ?? 1,
      0,
      0,
    )
    ctx.translate(-ax, -ay)
    ctx.fillStyle = normalizeColor(shadow.color)
    ctx.fill(path)
    return { canvas, x: -pad, y: -pad, w: boxW + 2 * pad, h: boxH + 2 * pad, under: true }
  }
  const built = build()
  shadowOverlayCache.set(key, built)
  return built
}

// softEdge feathering: source image + radius → offscreen canvas with edges fading to transparent (cached by url+radius)
const featherCache = new Map<string, HTMLCanvasElement>()

/**
 * Soft-edge approximation: a blurred, inset rectangle as an alpha mask (destination-in).
 * srcRadPx is the feather radius in source-image pixel space (caller converts by display scale).
 */
const insetTileCache = new Map<string, HTMLCanvasElement>()

/** Image composited into a transparent-padded tile per <a:stretch><a:fillRect> insets (negative insets crop). */
function insetFillTile(
  src: HTMLImageElement | HTMLCanvasElement,
  cacheKey: string,
  fr: { l: number; t: number; r: number; b: number },
): HTMLCanvasElement | HTMLImageElement {
  const key = `${cacheKey}|${fr.l}|${fr.t}|${fr.r}|${fr.b}`
  let c = insetTileCache.get(key)
  if (!c) {
    const rw = Math.max(1 - fr.l - fr.r, 0.01)
    const rh = Math.max(1 - fr.t - fr.b, 0.01)
    const iw = src.width || 1
    const ih = src.height || 1
    const scaleCap = Math.min(1, 2048 / (iw / rw), 2048 / (ih / rh))
    c = document.createElement('canvas')
    c.width = Math.max(1, Math.round((iw / rw) * scaleCap))
    c.height = Math.max(1, Math.round((ih / rh) * scaleCap))
    const ctx = c.getContext('2d')
    if (!ctx) return src
    ctx.drawImage(src, fr.l * c.width, fr.t * c.height, rw * c.width, rh * c.height)
    if (insetTileCache.size > 100) insetTileCache.clear()
    insetTileCache.set(key, c)
  }
  return c
}

const patternCellCache = new Map<string, HTMLCanvasElement>()
const patternCanvasCache = new Map<string, HTMLCanvasElement>()

/** One pattern cell (8×8 mask pixels) rendered crisp at the viewport cell size. */
function patternCellCanvas(
  preset: string,
  fg: string,
  bg: string,
  size: number,
): HTMLCanvasElement {
  const key = `${preset}|${fg}|${bg}|${size}`
  let c = patternCellCache.get(key)
  if (!c) {
    const grid = patternGrid(preset)
    const base = document.createElement('canvas')
    base.width = 8
    base.height = 8
    const bctx = base.getContext('2d')!
    bctx.fillStyle = normalizeColor(bg)
    bctx.fillRect(0, 0, 8, 8)
    bctx.fillStyle = normalizeColor(fg)
    for (let v = 0; v < 8; v++)
      for (let u = 0; u < 8; u++) if (grid[v]![u]) bctx.fillRect(u, v, 1, 1)
    c = document.createElement('canvas')
    c.width = size
    c.height = size
    const ctx = c.getContext('2d')!
    ctx.imageSmoothingEnabled = false
    ctx.drawImage(base, 0, 0, size, size)
    if (patternCellCache.size > 200) patternCellCache.clear()
    patternCellCache.set(key, c)
  }
  return c
}

/**
 * Shape-sized pattern canvas, cell grid phase-locked to the page origin (pre-composited
 * with no-repeat for the same Skia pixelRatio reason as anchoredTileCanvas).
 */
function patternCanvas(
  fill: { preset: string; fg: string; bg: string; cellPx: number },
  w: number,
  h: number,
  phase?: { x: number; y: number },
): HTMLCanvasElement | HTMLImageElement {
  const size = Math.max(2, Math.round(fill.cellPx))
  const cell = patternCellCanvas(fill.preset, fill.fg, fill.bg, size)
  if (w * h > 4096 * 4096) return cell
  const px = ((Math.round(phase?.x ?? 0) % size) + size) % size
  const py = ((Math.round(phase?.y ?? 0) % size) + size) % size
  const key = `${fill.preset}|${fill.fg}|${fill.bg}|${size}|${Math.ceil(w)}x${Math.ceil(h)}|${px},${py}`
  let c = patternCanvasCache.get(key)
  if (!c) {
    c = document.createElement('canvas')
    c.width = Math.max(1, Math.ceil(w))
    c.height = Math.max(1, Math.ceil(h))
    const ctx = c.getContext('2d')!
    for (let y = -py; y < h; y += size)
      for (let x = -px; x < w; x += size) ctx.drawImage(cell, x, y)
    if (patternCanvasCache.size > 100) patternCanvasCache.clear()
    patternCanvasCache.set(key, c)
  }
  return c
}

const anchoredTileCache = new Map<string, HTMLCanvasElement>()

/**
 * Compose an a:tile grid into a canvas covering the shape: tiles at the image's 144dpi
 * natural size x sx/sy (the caller bakes the dpi into t.scaleX/Y), anchored per algn
 * (tl..br) with tx/ty offsets, repeating over the whole shape box.
 */
function anchoredTileCanvas(
  src: HTMLImageElement | HTMLCanvasElement,
  cacheKey: string,
  w: number,
  h: number,
  t: { scaleX: number; scaleY: number; txPx: number; tyPx: number; algn: string },
): HTMLCanvasElement | HTMLImageElement {
  // A not-yet-decoded image has 0x0 dimensions: skip (and never cache) so the
  // image-load redraw composes the real tile grid
  if (!src.width || !src.height) return src
  // The caller draws the canvas 1:1 with no pattern transform (Skia pixelRatio bug),
  // so it must cover the shape exactly; bail out on extreme sizes instead of capping
  if (w * h > 4096 * 4096) return src
  const key = `${cacheKey}|tile|${src.width}x${src.height}|${Math.ceil(w)}x${Math.ceil(h)}|${t.scaleX.toFixed(4)}|${t.scaleY.toFixed(4)}|${Math.round(t.txPx)}|${Math.round(t.tyPx)}|${t.algn}`
  let c = anchoredTileCache.get(key)
  if (!c) {
    const tw = Math.max(src.width * t.scaleX, 1)
    const th = Math.max(src.height * t.scaleY, 1)
    const xFrac: Record<string, number> = {
      tl: 0,
      l: 0,
      bl: 0,
      t: 0.5,
      ctr: 0.5,
      b: 0.5,
      tr: 1,
      r: 1,
      br: 1,
    }
    const yFrac: Record<string, number> = {
      tl: 0,
      t: 0,
      tr: 0,
      l: 0.5,
      ctr: 0.5,
      r: 0.5,
      bl: 1,
      b: 1,
      br: 1,
    }
    const ax = (xFrac[t.algn] ?? 0) * (w - tw) + t.txPx
    const ay = (yFrac[t.algn] ?? 0) * (h - th) + t.tyPx
    c = document.createElement('canvas')
    c.width = Math.max(1, Math.ceil(w))
    c.height = Math.max(1, Math.ceil(h))
    const ctx = c.getContext('2d')
    if (!ctx) return src
    const x0 = ax - Math.ceil(ax / tw) * tw
    const y0 = ay - Math.ceil(ay / th) * th
    for (let y = y0; y < h; y += th) {
      for (let x = x0; x < w; x += tw) {
        ctx.drawImage(src, x, y, tw, th)
      }
    }
    if (anchoredTileCache.size > 100) anchoredTileCache.clear()
    anchoredTileCache.set(key, c)
  }
  return c
}

/** Both dimensions ≤2px: PowerPoint rasterizes such stretched images as one flat color. */
export function isDegenerateImage(img: { width: number; height: number }): boolean {
  return img.width > 0 && img.height > 0 && img.width <= 2 && img.height <= 2
}

const flatImageCache = new Map<string, HTMLCanvasElement>()

/** 1×1 canvas of the image's mean color (degenerate pictures stretch to a flat surface). */
export function flatColorImage(
  img: HTMLImageElement,
  cacheKey: string,
): HTMLCanvasElement | HTMLImageElement {
  let c = flatImageCache.get(cacheKey)
  if (!c) {
    c = document.createElement('canvas')
    c.width = 1
    c.height = 1
    const ctx = c.getContext('2d')
    if (!ctx) return img
    ctx.fillStyle = averageColor(img, cacheKey)
    ctx.fillRect(0, 0, 1, 1)
    if (flatImageCache.size > 100) flatImageCache.clear()
    flatImageCache.set(cacheKey, c)
  }
  return c
}

const avgColorCache = new Map<string, string>()

/** Mean RGB of an image (degenerate-texture flat fill). */
function averageColor(img: HTMLImageElement | HTMLCanvasElement, cacheKey: string): string {
  let c = avgColorCache.get(cacheKey)
  if (!c) {
    c = '#ffffff'
    try {
      const cv = document.createElement('canvas')
      cv.width = img.width
      cv.height = img.height
      const ctx = cv.getContext('2d')
      if (!ctx) throw new Error('no 2d context')
      ctx.drawImage(img, 0, 0)
      const px = ctx.getImageData(0, 0, cv.width, cv.height).data
      let r = 0
      let g = 0
      let b = 0
      const n = px.length / 4
      for (let i = 0; i < px.length; i += 4) {
        r += px[i]!
        g += px[i + 1]!
        b += px[i + 2]!
      }
      const h = (v: number) =>
        Math.round(v / n)
          .toString(16)
          .padStart(2, '0')
      c = `#${h(r)}${h(g)}${h(b)}`
    } catch {
      // tainted canvas → keep white
    }
    avgColorCache.set(cacheKey, c)
  }
  return c
}

const duotoneCache = new Map<string, HTMLCanvasElement>()
const lumCache = new Map<string, HTMLCanvasElement>()

type ClrChange = { from: string; to: string }
type Lum = { bright: number; contrast: number }

export function processedImageKey(
  dataUrl: string,
  clrChange?: ClrChange,
  duotone?: [string, string],
  lum?: Lum,
): string {
  let key = dataUrl
  if (clrChange) key += `|cc:${clrChange.from}>${clrChange.to}`
  if (duotone) key += `|${duotone[0]}|${duotone[1]}`
  if (lum) key += `|lum:${lum.bright},${lum.contrast}`
  return key
}

/** Apply blip pixel effects in PowerPoint's order: clrChange, then duotone, then lum. */
export function processedImage(
  img: HTMLImageElement,
  dataUrl: string,
  clrChange?: ClrChange,
  duotone?: [string, string],
  lum?: Lum,
): CanvasImageSource {
  let src: HTMLImageElement | HTMLCanvasElement = img
  if (clrChange) src = clrChangeImage(src, `${dataUrl}|cc`, clrChange.from, clrChange.to)
  if (duotone)
    src = duotoneImage(src, processedImageKey(dataUrl, clrChange), duotone[0], duotone[1])
  if (lum) src = lumImage(src, processedImageKey(dataUrl, clrChange, duotone), lum)
  return src
}

/**
 * <a:lum> brightness/contrast: contrast scales around mid gray, then positive brightness
 * blends toward white / negative toward black. Calibrated against PowerPoint on a
 * bright=70% contrast=-70% washed-out photo box (navy 0.09 -> 0.813 measured on the ref).
 */
export function lumImage(
  img: HTMLImageElement | HTMLCanvasElement,
  cacheKey: string,
  lum: Lum,
): HTMLCanvasElement {
  const key = `${cacheKey}|lum:${lum.bright},${lum.contrast}`
  let c = lumCache.get(key)
  if (!c) {
    c = document.createElement('canvas')
    c.width = img.width || 1
    c.height = img.height || 1
    const ctx = c.getContext('2d')!
    ctx.drawImage(img, 0, 0)
    const slope = lum.contrast >= 0 ? 1 / (1 - Math.min(lum.contrast, 0.99)) : 1 + lum.contrast
    const lut = new Uint8ClampedArray(256)
    for (let i = 0; i < 256; i++) {
      let v = ((i / 255 - 0.5) * slope + 0.5) * 255
      v = lum.bright >= 0 ? v + lum.bright * (255 - v) : v * (1 + lum.bright)
      lut[i] = Math.max(0, Math.min(255, Math.round(v)))
    }
    try {
      const data = ctx.getImageData(0, 0, c.width, c.height)
      const px = data.data
      for (let i = 0; i < px.length; i += 4) {
        px[i] = lut[px[i]!]!
        px[i + 1] = lut[px[i + 1]!]!
        px[i + 2] = lut[px[i + 2]!]!
      }
      ctx.putImageData(data, 0, 0)
    } catch {
      /* tainted canvas: keep the original pixels */
    }
    if (lumCache.size > 100) lumCache.clear()
    lumCache.set(key, c)
  }
  return c
}

const clrChangeCache = new Map<string, HTMLCanvasElement>()

/** clrChange (<a:clrChange>): pixels matching `from` are replaced with `to` (#RRGGBBAA alpha honored). */
export function clrChangeImage(
  img: HTMLImageElement | HTMLCanvasElement,
  cacheKey: string,
  from: string,
  to: string,
): HTMLCanvasElement {
  const key = `${cacheKey}|${from}>${to}`
  let c = clrChangeCache.get(key)
  if (!c) {
    c = document.createElement('canvas')
    c.width = img.width || 1
    c.height = img.height || 1
    const ctx = c.getContext('2d')!
    ctx.drawImage(img, 0, 0)
    const hex = (s: string, i: number) => parseInt(s.slice(i, i + 2), 16) || 0
    const f = from.replace('#', '')
    const t = to.replace('#', '')
    const [fr, fg, fb] = [hex(f, 0), hex(f, 2), hex(f, 4)]
    const [tr, tg, tb] = [hex(t, 0), hex(t, 2), hex(t, 4)]
    const ta = t.length >= 8 ? hex(t, 6) : 255
    // small per-channel tolerance absorbs rasterizer rounding (metafile conversions)
    const TOL = 4
    try {
      const data = ctx.getImageData(0, 0, c.width, c.height)
      const px = data.data
      for (let i = 0; i < px.length; i += 4) {
        if (
          Math.abs(px[i]! - fr) <= TOL &&
          Math.abs(px[i + 1]! - fg) <= TOL &&
          Math.abs(px[i + 2]! - fb) <= TOL
        ) {
          px[i] = tr
          px[i + 1] = tg
          px[i + 2] = tb
          px[i + 3] = Math.round((px[i + 3]! * ta) / 255)
        }
      }
      ctx.putImageData(data, 0, 0)
    } catch {
      /* tainted canvas: keep the original pixels */
    }
    if (clrChangeCache.size > 100) clrChangeCache.clear()
    clrChangeCache.set(key, c)
  }
  return c
}

/** Duotone (<a:duotone>): image luminance interpolates [dark, light]; alpha preserved. */
export function duotoneImage(
  img: HTMLImageElement | HTMLCanvasElement,
  cacheKey: string,
  dark: string,
  light: string,
): HTMLCanvasElement {
  const key = `${cacheKey}|${dark}|${light}`
  let c = duotoneCache.get(key)
  if (!c) {
    c = document.createElement('canvas')
    c.width = img.width || 1
    c.height = img.height || 1
    const ctx = c.getContext('2d')!
    ctx.drawImage(img, 0, 0)
    const hex = (s: string) => {
      const h = s.replace('#', '')
      return [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16) || 0)
    }
    const d = hex(dark)
    const l = hex(light)
    try {
      const data = ctx.getImageData(0, 0, c.width, c.height)
      const px = data.data
      for (let i = 0; i < px.length; i += 4) {
        const lum = (0.299 * px[i]! + 0.587 * px[i + 1]! + 0.114 * px[i + 2]!) / 255
        px[i] = Math.round(d[0]! + (l[0]! - d[0]!) * lum)
        px[i + 1] = Math.round(d[1]! + (l[1]! - d[1]!) * lum)
        px[i + 2] = Math.round(d[2]! + (l[2]! - d[2]!) * lum)
      }
      ctx.putImageData(data, 0, 0)
    } catch {
      /* tainted canvas: keep the original pixels */
    }
    if (duotoneCache.size > 100) duotoneCache.clear()
    duotoneCache.set(key, c)
  }
  return c
}

/** Solid-fill shape with <a:softEdge>: the (rounded) rect pre-rendered with feathered edges. */
const featherShapeCache = new Map<string, HTMLCanvasElement>()
export function featheredShapeCanvas(
  color: string,
  w: number,
  h: number,
  radPx: number,
  cornerRadiusPx = 0,
  stroke?: { color: string; widthPx: number },
): { canvas: HTMLCanvasElement; pad: number } {
  const cw = Math.max(1, Math.round(w))
  const ch = Math.max(1, Math.round(h))
  const rad = Math.max(1, Math.round(radPx))
  // Canvas strokes are centered on the path: pad the canvas so the outer half survives
  const pad = stroke ? Math.ceil(stroke.widthPx / 2) : 0
  const key = `${color}|${cw}|${ch}|${rad}|${Math.round(cornerRadiusPx)}|${stroke ? `${stroke.color}/${stroke.widthPx}` : ''}`
  let c = featherShapeCache.get(key)
  if (!c) {
    c = document.createElement('canvas')
    c.width = cw + 2 * pad
    c.height = ch + 2 * pad
    const ctx = c.getContext('2d')!
    ctx.beginPath()
    if (cornerRadiusPx > 0) ctx.roundRect(pad, pad, cw, ch, cornerRadiusPx)
    else ctx.rect(pad, pad, cw, ch)
    ctx.fillStyle = normalizeColor(color)
    ctx.fill()
    // The outline feathers together with the fill (PowerPoint applies softEdge to the composite)
    if (stroke) {
      ctx.strokeStyle = normalizeColor(stroke.color)
      ctx.lineWidth = stroke.widthPx
      ctx.stroke()
    }
    // Same alpha ramp as featheredImage: mask edge at 0.5r inset, blur 0.67r. The ramp
    // starts from the composite's outer edge (stroke included) — the full padded canvas —
    // so the padded outline half is inside the mask interior, not under its falloff.
    ctx.globalCompositeOperation = 'destination-in'
    ctx.filter = `blur(${(rad * 2) / 3}px)`
    ctx.fillStyle = '#000'
    ctx.fillRect(rad * 0.5, rad * 0.5, c.width - rad, c.height - rad)
    if (featherShapeCache.size > 100) featherShapeCache.clear()
    featherShapeCache.set(key, c)
  }
  return { canvas: c, pad }
}

export function featheredImage(
  img: HTMLImageElement,
  cacheKey: string,
  srcRadPx: number,
): CanvasImageSource {
  const rad = Math.max(1, Math.round(srcRadPx))
  const key = `${cacheKey}|${rad}`
  let c = featherCache.get(key)
  if (!c) {
    c = document.createElement('canvas')
    c.width = img.width || 1
    c.height = img.height || 1
    const ctx = c.getContext('2d')!
    ctx.drawImage(img, 0, 0)
    ctx.globalCompositeOperation = 'destination-in'
    // PowerPoint's feather ramps from opaque at rad inside the edge to transparent AT the
    // edge (50% alpha ~0.5r inside — NASA moon-limb measurement); a mask edge at 0.5r with
    // a 0.67r blur (sigma r/3) reproduces that ramp. The previous 1.5r inset ate a full
    // extra radius of image.
    ctx.filter = `blur(${(rad * 2) / 3}px)`
    ctx.fillStyle = '#000'
    ctx.fillRect(rad * 0.5, rad * 0.5, c.width - rad, c.height - rad)
    if (featherCache.size > 100) featherCache.clear()
    featherCache.set(key, c)
  }
  return c
}

/**
 * Picture srcRect crop ratios → Konva Image crop (source-image pixel coordinates).
 * Negative srcRect values are insets: the source rect extends past the image, so the
 * image occupies only a sub-rect of the frame and the rest stays empty (PowerPoint
 * leaves those bands blank; common in Google Slides exports).
 */
export function cropToKonva(
  pic: PictureRenderNode,
  img: HTMLImageElement | undefined,
): {
  crop?: { x: number; y: number; width: number; height: number }
  x?: number
  y?: number
  width?: number
  height?: number
} {
  const sr = pic.srcRect
  if (!sr || !img || !img.width || !img.height) return {}
  const x0 = Math.max(sr.l, 0)
  const y0 = Math.max(sr.t, 0)
  const x1 = 1 - Math.max(sr.r, 0)
  const y1 = 1 - Math.max(sr.b, 0)
  const crop = {
    x: x0 * img.width,
    y: y0 * img.height,
    width: Math.max(img.width * (x1 - x0), 1),
    height: Math.max(img.height * (y1 - y0), 1),
  }
  if (sr.l >= 0 && sr.t >= 0 && sr.r >= 0 && sr.b >= 0) return { crop }
  // The frame maps source span [l, 1-r]×[t, 1-b] onto the full box; place the
  // visible sub-rect of the image at its share of that span.
  const spanX = 1 - sr.l - sr.r
  const spanY = 1 - sr.t - sr.b
  if (spanX <= 0 || spanY <= 0) return { crop }
  const boxW = pic.box.w
  const boxH = pic.box.h
  return {
    crop,
    x: (boxW * (x0 - sr.l)) / spanX,
    y: (boxH * (y0 - sr.t)) / spanY,
    width: Math.max((boxW * (x1 - x0)) / spanX, 1),
    height: Math.max((boxH * (y1 - y0)) / spanY, 1),
  }
}

/** Preset geometry name → Konva shape type (Phase 3 supports the common ones, the rest approximated as rectangles). */
export type ShapeKind = 'rect' | 'roundRect' | 'ellipse'

export function presetToShapeKind(preset: string | undefined): ShapeKind {
  switch (preset) {
    case 'ellipse':
    case 'circle':
      return 'ellipse'
    case 'roundRect':
      return 'roundRect'
    default:
      return 'rect'
  }
}

/** Absolute positioning of one glyph run as Konva Text (relative to the text box top-left). */
export interface GlyphDraw {
  text: string
  x: number
  /** Konva Text y = top, derived from the baseline: baselineY - ascent ≈ fontSize*0.8 */
  y: number
  fontSize: number
  fontFamily: string
  fill: string
  fontStyle: string // 'bold' | 'italic' | 'bold italic' | 'normal'
  textDecoration: string // 'underline' / 'line-through' space-joined combination, '' = none
  /** Letter spacing (px/char, may be negative) — layout width already includes it; drawing must feed it to Konva too */
  letterSpacing?: number
  /** Text outline (WordArt): stroke first then fill, so the stroke does not eat into glyph interiors */
  stroke?: string
  strokeWidth?: number
  fillAfterStrokeEnabled?: boolean
  /** RTL run: feeds Konva Text's direction so neutral punctuation lands on the correct side */
  direction?: 'rtl'
  /** Rotated glyph (eaVert latin words / vert blocks: 90; vert270 blocks: -90; x/y is the rotation anchor) */
  rotation?: number
  /** Text highlight (<a:rPr><a:highlight>): background rect drawn behind the run, covering the line box */
  highlight?: { x: number; y: number; w: number; h: number; color: string }
  shadowColor?: string
  shadowBlur?: number
  shadowOffsetX?: number
  shadowOffsetY?: number
  shadowEnabled?: boolean
  /** WordArt gradient text fill (mapped to Konva's linear-gradient fill over the run box) */
  fillPriority?: 'linear-gradient'
  fillLinearGradientStartPoint?: { x: number; y: number }
  fillLinearGradientEndPoint?: { x: number; y: number }
  fillLinearGradientColorStops?: Array<number | string>
  /** Run reflection: the renderer draws a faded mirrored copy below the text */
  reflection?: boolean
  /** WordArt warp transforms (per-character; position is the char center when offsets are set) */
  scaleX?: number
  scaleY?: number
  offsetX?: number
  offsetY?: number
}

// Same-script fallback chains for Japanese/Korean/Traditional Chinese (win/mac family names back each other up); shared by FONT_STACK and the unknown-font fallback
const JA_SANS = "'Yu Gothic', 'Hiragino Sans', Meiryo, 'Noto Sans JP', sans-serif"
const JA_SERIF = "'Yu Mincho', 'Hiragino Mincho ProN', 'MS Mincho', 'Noto Serif JP', serif"
const KO_SANS = "'Malgun Gothic', 'Apple SD Gothic Neo', 'Noto Sans KR', sans-serif"
const KO_SERIF = "Batang, AppleMyungjo, 'Noto Serif KR', serif"
const TC_SANS = "'Microsoft JhengHei', 'PingFang TC', 'Heiti TC', 'Noto Sans TC', sans-serif"
const TC_SERIF = "PMingLiU, 'Songti TC', 'Noto Serif TC', serif"
const SERIF_HINT_RE =
  /mincho|明朝|batang|바탕|myeongjo|명조|gungsuh|궁서|mingliu|細明|標楷|宋|song/i

/**
 * Display font stack: font names in the file may not be installed locally (Microsoft YaHei
 * on mac / PingFang on win), so append cross-platform equivalents as CSS-level fallbacks.
 * Metrics are handled by the main process FontMetricsProvider's alias table.
 */
const FONT_STACK: Record<string, string> = {
  'microsoft yahei': "'Microsoft YaHei', 'PingFang SC', 'Noto Sans SC', sans-serif",
  微软雅黑: "'Microsoft YaHei', 'PingFang SC', 'Noto Sans SC', sans-serif",
  'pingfang sc': "'PingFang SC', 'Microsoft YaHei', 'Noto Sans SC', sans-serif",
  苹方: "'PingFang SC', 'Microsoft YaHei', 'Noto Sans SC', sans-serif",
  宋体: "SimSun, 'Songti SC', serif",
  simsun: "SimSun, 'Songti SC', serif",
  黑体: "SimHei, 'Heiti SC', sans-serif",
  simhei: "SimHei, 'Heiti SC', sans-serif",
  楷体: "KaiTi, 'Kaiti SC', serif",
  仿宋: "FangSong, 'Songti SC', serif",
  等线: "DengXian, 'Microsoft YaHei', 'PingFang SC', sans-serif",
  dengxian: "DengXian, 'Microsoft YaHei', 'PingFang SC', sans-serif",
  // Western: consistent with the main-process metrics alias chain (calibri→Carlito→Arial etc.),
  // otherwise metrics use Arial while drawing falls back to the system default font, misaligning word spacing/line breaks.
  calibri: 'Calibri, Carlito, Arial, sans-serif',
  'calibri light': "'Calibri Light', Carlito, Arial, sans-serif",
  helvetica: 'Helvetica, Arial, sans-serif',
  'helvetica neue': "'Helvetica Neue', Helvetica, Arial, sans-serif",
  cambria: 'Cambria, Georgia, serif',
  // Japanese (win family names <-> mac Hiragino back each other up; Japanese fonts first, then Chinese fallback, so kanji don't render with Chinese glyph shapes)
  'yu gothic': JA_SANS,
  游ゴシック: "'游ゴシック', " + JA_SANS,
  meiryo: 'Meiryo, ' + JA_SANS,
  メイリオ: "'メイリオ', Meiryo, " + JA_SANS,
  'ms gothic': "'MS Gothic', 'MS PGothic', " + JA_SANS,
  'ms pgothic': "'MS PGothic', 'MS Gothic', " + JA_SANS,
  'ms ui gothic': "'MS UI Gothic', 'MS PGothic', " + JA_SANS,
  'ms ゴシック': "'ＭＳ ゴシック', 'MS Gothic', " + JA_SANS,
  'ms pゴシック': "'ＭＳ Ｐゴシック', 'MS PGothic', " + JA_SANS,
  'hiragino sans': "'Hiragino Sans', " + JA_SANS,
  'hiragino kaku gothic pron': "'Hiragino Kaku Gothic ProN', " + JA_SANS,
  ヒラギノ角ゴシック: "'Hiragino Sans', " + JA_SANS,
  'noto sans jp': "'Noto Sans JP', " + JA_SANS,
  'yu mincho': JA_SERIF,
  游明朝: "'游明朝', " + JA_SERIF,
  'ms mincho': "'MS Mincho', 'MS PMincho', " + JA_SERIF,
  'ms pmincho': "'MS PMincho', 'MS Mincho', " + JA_SERIF,
  'ms 明朝': "'ＭＳ 明朝', 'MS Mincho', " + JA_SERIF,
  'ms p明朝': "'ＭＳ Ｐ明朝', 'MS PMincho', " + JA_SERIF,
  'hiragino mincho pron': "'Hiragino Mincho ProN', " + JA_SERIF,
  ヒラギノ明朝: "'Hiragino Mincho ProN', " + JA_SERIF,
  'noto serif jp': "'Noto Serif JP', " + JA_SERIF,
  // Korean
  'malgun gothic': KO_SANS,
  '맑은 고딕': "'맑은 고딕', " + KO_SANS,
  'apple sd gothic neo': "'Apple SD Gothic Neo', 'Malgun Gothic', 'Noto Sans KR', sans-serif",
  gulim: 'Gulim, Dotum, ' + KO_SANS,
  굴림: "'굴림', Gulim, Dotum, " + KO_SANS,
  dotum: 'Dotum, Gulim, ' + KO_SANS,
  돋움: "'돋움', Dotum, Gulim, " + KO_SANS,
  'noto sans kr': "'Noto Sans KR', " + KO_SANS,
  batang: KO_SERIF,
  바탕: "'바탕', " + KO_SERIF,
  gungsuh: 'Gungsuh, ' + KO_SERIF,
  궁서: "'궁서', Gungsuh, " + KO_SERIF,
  // Traditional Chinese
  'microsoft jhenghei': TC_SANS,
  微軟正黑體: "'微軟正黑體', " + TC_SANS,
  'pingfang tc': "'PingFang TC', 'Microsoft JhengHei', 'Heiti TC', 'Noto Sans TC', sans-serif",
  'pingfang hk': "'PingFang HK', 'PingFang TC', 'Microsoft JhengHei', 'Noto Sans TC', sans-serif",
  pmingliu: TC_SERIF,
  新細明體: "'新細明體', " + TC_SERIF,
  mingliu: "MingLiU, 'PMingLiU', 'Songti TC', serif",
  細明體: "'細明體', MingLiU, 'Songti TC', serif",
  'dfkai-sb': "'DFKai-SB', BiauKai, 'Kaiti TC', serif",
  標楷體: "'標楷體', 'DFKai-SB', BiauKai, 'Kaiti TC', serif",
}

// These families resolve to Lucida Grande on macOS, which has no bold face; PowerPoint
// for Mac renders their b="1" runs at regular weight, while Chromium would fake-bold them
const NO_SYNTHETIC_BOLD = new Set(['lucida sans unicode', 'lucida sans', 'lucida grande'])

export function displayFontFamily(name: string): string {
  const stack = FONT_STACK[name.normalize('NFKC').toLowerCase()]
  if (stack) return stack
  // Unknown fonts first get script detection by family name and same-script fallback, so Japanese/Korean/Traditional glyphs don't render as Simplified Chinese shapes
  const script = classifyCjkScript(name)
  if (script === 'ja') return `'${name}', ${SERIF_HINT_RE.test(name) ? JA_SERIF : JA_SANS}`
  if (script === 'ko') return `'${name}', ${SERIF_HINT_RE.test(name) ? KO_SERIF : KO_SANS}`
  if (script === 'tc') return `'${name}', ${SERIF_HINT_RE.test(name) ? TC_SERIF : TC_SANS}`
  return `'${name}', 'PingFang SC', 'Microsoft YaHei', sans-serif`
}

/**
 * Baseline anchoring for painted text, measured per family/style at a fixed size
 * (font metrics scale linearly). Konva paints with canvas2d textBaseline='middle'
 * anchored 0.5em below the node top; the middle→alphabetic distance is a per-font
 * quantity Chromium derives from the drawing font's metrics — measuring the same ink
 * from both baselines yields it exactly.
 *
 * Whether to paint the alphabetic baseline exactly on the engine's baselineY splits
 * by font resolution (both branches pixel-matched against PowerPoint on prod run3):
 * - The requested family resolves to a real face (system font or a registered private
 *   FontFace): exact anchoring matches PowerPoint — the fixed 0.8em legacy rule sat
 *   several px low on JP faces (prod_020/024/072 Meiryo UI improved 3-5pp with exact).
 * - The family falls back to a script default (localized names, missing fonts):
 *   PowerPoint's own substitute rendering sits where the legacy 0.8em rule painted,
 *   and exact anchoring via the fallback font's metrics lands ~2px high
 *   (prod_050 bisected: the engine line model was calibrated through the 0.8em rule
 *   for fallback-drawn text).
 */
const MEASURE_PX = 100
let baselineCtx: CanvasRenderingContext2D | null | undefined
const baselineAnchor = new Map<string, { ratio: number; resolved: boolean }>()
// Private/embedded FontFaces register after the first React render, and a Konva batchDraw
// repaints existing nodes without recomputing their glyph y props — anchors measured
// against the fallback font would stick. Bump an epoch on 'loadingdone' so glyph-building
// components re-render (recomputing anchors against the now-loaded faces).
let fontsEpoch = 0
const fontsEpochSubs = new Set<() => void>()
if (typeof document !== 'undefined')
  document.fonts?.addEventListener?.('loadingdone', () => {
    baselineAnchor.clear()
    fontsEpoch++
    for (const cb of fontsEpochSubs) cb()
  })

export function subscribeFontsEpoch(cb: () => void): () => void {
  fontsEpochSubs.add(cb)
  return () => fontsEpochSubs.delete(cb)
}
export function getFontsEpoch(): number {
  return fontsEpoch
}

const RESOLVE_PROBE = 'Ilmg19日本語'

function measureBaselineAnchor(
  familyList: string,
  bold?: boolean,
  italic?: boolean,
): { ratio: number; resolved: boolean } {
  const key = `${familyList}|${bold ? 'b' : ''}${italic ? 'i' : ''}`
  const hit = baselineAnchor.get(key)
  if (hit) return hit
  if (baselineCtx === undefined)
    baselineCtx =
      typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null
  if (!baselineCtx) return { ratio: 0.8, resolved: false } // no canvas (unit tests): legacy rule
  const ctx = baselineCtx
  const style = `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}`
  // Does the primary family resolve? Its advance must differ from a bare generic under
  // at least one generic fallback (a real face can't metric-match both serif and monospace)
  const primary = /^'([^']+)'/.exec(familyList)?.[1] ?? familyList.split(',')[0]!.trim()
  const w = (font: string) => {
    ctx.font = `${style}${MEASURE_PX}px ${font}`
    return ctx.measureText(RESOLVE_PROBE).width
  }
  const resolved =
    w(`'${primary}', monospace`) !== w('monospace') || w(`'${primary}', serif`) !== w('serif')
  // Same ink measured from both baselines: the ascent difference IS the exact distance
  // from the 'middle' anchor down to the alphabetic baseline for the drawing font
  ctx.font = `${style}${MEASURE_PX}px ${familyList}`
  ctx.textBaseline = 'alphabetic'
  const a1 = ctx.measureText('Hg').actualBoundingBoxAscent
  ctx.textBaseline = 'middle'
  const a2 = ctx.measureText('Hg').actualBoundingBoxAscent
  const ratio = a1 > 0 || a2 > 0 ? 0.5 + (a1 - a2) / MEASURE_PX : 0.8
  const anchor = { ratio, resolved }
  baselineAnchor.set(key, anchor)
  return anchor
}

/** baselineY − node top for painted text (see measureBaselineAnchor for the split). */
function konvaBaselineAscent(
  familyList: string,
  sizePx: number,
  bold?: boolean,
  italic?: boolean,
): number {
  const a = measureBaselineAnchor(familyList, bold, italic)
  return (a.resolved ? a.ratio : 0.8) * sizePx
}

/** Painted-baseline − engine-baselineY delta (the edit overlay cancels it to match pixels). */
export function konvaBaselineDrop(
  familyList: string,
  sizePx: number,
  bold?: boolean,
  italic?: boolean,
): number {
  const a = measureBaselineAnchor(familyList, bold, italic)
  return a.resolved ? 0 : (a.ratio - 0.8) * sizePx
}

export function glyphToDraw(run: GlyphRun): GlyphDraw {
  const styleParts: string[] = []
  const effectiveBold =
    !!run.bold && !NO_SYNTHETIC_BOLD.has(run.fontFamily.normalize('NFKC').toLowerCase())
  if (effectiveBold) styleParts.push('bold')
  if (run.italic) styleParts.push('italic')
  const family = displayFontFamily(run.fontFamily)
  return {
    text: run.text,
    x: run.x,
    y: run.baselineY - konvaBaselineAscent(family, run.fontSizePx, effectiveBold, run.italic),
    fontSize: run.fontSizePx,
    fontFamily: family,
    fill: normalizeColor(run.color),
    fontStyle: styleParts.join(' ') || 'normal',
    textDecoration: [run.underline ? 'underline' : '', run.strike ? 'line-through' : '']
      .filter(Boolean)
      .join(' '),
    ...((run.letterSpacingPx ?? 0) + (run.justifyExtraPx ?? 0)
      ? { letterSpacing: (run.letterSpacingPx ?? 0) + (run.justifyExtraPx ?? 0) }
      : run.kerningOff && !run.rtl
        ? // Kerning must stay off (fontSize below the rPr kern threshold): a negligible
          // letterSpacing flips Konva into per-letter drawing, which never kerns —
          // matching the unkerned measurement (canvas fillText would kern the string)
          { letterSpacing: 1e-4 }
        : {}),
    ...(run.outline
      ? {
          stroke: normalizeColor(run.outline.color),
          strokeWidth: Math.max(run.outline.widthPx, 0.5),
          fillAfterStrokeEnabled: true,
        }
      : {}),
    ...(run.rtl ? { direction: 'rtl' as const } : {}),
    ...(run.rotate90 ? { rotation: 90 } : run.rotate270 ? { rotation: -90 } : {}),
    ...(run.shadow
      ? {
          shadowColor: normalizeColor(run.shadow.color),
          shadowBlur: run.shadow.blurPx,
          shadowOffsetX: run.shadow.offsetX,
          shadowOffsetY: run.shadow.offsetY,
          shadowEnabled: true,
        }
      : run.glow
        ? {
            shadowColor: normalizeColor(run.glow.color),
            shadowBlur: run.glow.blurPx,
            shadowOffsetX: 0,
            shadowOffsetY: 0,
            shadowEnabled: true,
          }
        : {}),
    ...(run.gradient
      ? (() => {
          // Gradient in Text-node-local coords (origin = glyph top-left); 90° = top→bottom
          // across the em box, 0° = left→right across the run width
          const [gx, gy] = linearGradientDirection(
            run.gradient.angleDeg,
            run.gradient.scaled,
            run.widthPx,
            run.fontSizePx,
          )
          const cx = run.widthPx / 2
          const cy = run.fontSizePx * 0.5
          // Same projection as shape fills: ramp length = the box's projection onto the direction
          const len = Math.abs(gx) * run.widthPx + Math.abs(gy) * run.fontSizePx
          return {
            fillPriority: 'linear-gradient' as const,
            fillLinearGradientStartPoint: { x: cx - (gx * len) / 2, y: cy - (gy * len) / 2 },
            fillLinearGradientEndPoint: { x: cx + (gx * len) / 2, y: cy + (gy * len) / 2 },
            fillLinearGradientColorStops: linearRampStops(run.gradient.stops),
          }
        })()
      : {}),
    ...(run.reflection ? { reflection: true } : {}),
  }
}

/** Collect all glyph draws of a laid-out text block (shared by shapes / table cells). */
export function layoutGlyphs(text: RenderTextLayout | undefined): GlyphDraw[] {
  if (!text) return []
  const out: GlyphDraw[] = []
  for (const line of text.lines) {
    for (const run of line.runs) {
      const g = glyphToDraw(run)
      // PowerPoint draws the highlight over the full line box (not just the glyph extent).
      // Vertical layouts are skipped entirely: their "lines" are full columns, so the
      // line box would paint a column-tall background.
      if (run.highlight && !text.vert) {
        g.highlight = {
          x: run.x,
          y: line.top,
          w: run.widthPx,
          h: line.height,
          color: normalizeColor(run.highlight),
        }
      }
      out.push(g)
    }
  }
  return out
}

/** Collect all glyph draws in a shape node (including the offset relative to the text box). */
export function shapeGlyphs(node: ShapeRenderNode): GlyphDraw[] {
  return layoutGlyphs(node.text)
}

/** Normalize a color for CSS (#RRGGBB / #RRGGBBAA → rgba). */
export function normalizeColor(c: string): string {
  if (c === 'none') return 'transparent'
  if (/^#?[0-9A-Fa-f]{8}$/.test(c)) {
    const h = c.replace(/^#/, '')
    const r = parseInt(h.slice(0, 2), 16)
    const g = parseInt(h.slice(2, 4), 16)
    const b = parseInt(h.slice(4, 6), 16)
    const a = parseInt(h.slice(6, 8), 16) / 255
    return `rgba(${r},${g},${b},${a.toFixed(3)})`
  }
  return c.startsWith('#') || c.startsWith('rgb') ? c : `#${c}`
}

export function isEditableText(node: RenderNode): node is ShapeRenderNode {
  if (node.type !== 'text' && node.type !== 'shape') return false
  // A shape with no text body yet is still editable — setText creates the body and the
  // engine injects <p:txBody>. Connectors are the exception: CT_Connector has no
  // txBody child, so a line can never hold text.
  return !(node as ShapeRenderNode).line
}

/** Polyline smooth → Konva Line tension (0 = polyline, 0.4 ≈ PPT smooth curve). */
export function smoothTension(smooth: boolean | undefined): number {
  return smooth ? 0.4 : 0
}
