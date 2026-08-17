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

export function fillToKonva(
  fill: RenderFill,
  w: number,
  h: number,
  images?: Map<string, HTMLImageElement>,
): KonvaFillProps {
  switch (fill.kind) {
    case 'solid':
      return { fill: normalizeColor(fill.color) }
    case 'gradient': {
      if (fill.radial) {
        // Radial (circle/rect/shape all approximated as circular). Center follows fillToRect;
        // the 100% ring sits well beyond the farthest corner (×1.8): PowerPoint's path-gradient
        // falloff is much slower than a corner-normalized radial (calibrated against PPT renders).
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
          fillRadialGradientColorStops: fill.stops.flatMap((s) => [s.pos, normalizeColor(s.color)]),
        }
      }
      const sorted = [...fill.stops].sort((a, b) => a.pos - b.pos)
      const rad = (fill.angleDeg * Math.PI) / 180
      const dx = Math.cos(rad)
      const dy = Math.sin(rad)
      // Gradient vector length = the box's projection onto the gradient direction,
      // so the ramp spans exactly the shape (PowerPoint semantics for a:lin)
      const cx = w / 2
      const cy = h / 2
      const len = Math.abs(dx) * w + Math.abs(dy) * h
      return {
        fillLinearGradientStartPoint: { x: cx - (dx * len) / 2, y: cy - (dy * len) / 2 },
        fillLinearGradientEndPoint: { x: cx + (dx * len) / 2, y: cy + (dy * len) / 2 },
        fillLinearGradientColorStops: sorted.flatMap((s) => [s.pos, normalizeColor(s.color)]),
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
        ) as HTMLImageElement
        // recolored variants must not share cache slots with the raw image
        const srcKey = processedImageKey(fill.dataUrl ?? '', fill.clrChange, fill.duotone)
        if (fill.mode === 'tile') {
          // PowerPoint tiles at the image's 96dpi natural size x sx/sy, anchored per algn
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

export function strokeToKonva(stroke: RenderStroke | undefined): {
  stroke?: string
  strokeWidth?: number
  dash?: number[]
} {
  if (!stroke) return {}
  return {
    stroke: normalizeColor(stroke.color),
    strokeWidth: stroke.widthPx,
    ...(stroke.dash ? { dash: stroke.dash } : {}),
  }
}

/** Outer shadow → Konva shadow attributes; without a shadow, glow is approximated with a zero-offset shadow. */
export function shadowToKonva(
  shadow: RenderShadow | undefined,
  glow?: RenderGlow,
): {
  shadowColor?: string
  shadowBlur?: number
  shadowOffsetX?: number
  shadowOffsetY?: number
  shadowEnabled?: boolean
} {
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

const anchoredTileCache = new Map<string, HTMLCanvasElement>()

/**
 * Compose an a:tile grid into a canvas covering the shape: tiles at the image's 96dpi
 * natural size x sx/sy, anchored per algn (tl..br) with tx/ty offsets, repeating over
 * the whole shape box.
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

type ClrChange = { from: string; to: string }

export function processedImageKey(
  dataUrl: string,
  clrChange?: ClrChange,
  duotone?: [string, string],
): string {
  let key = dataUrl
  if (clrChange) key += `|cc:${clrChange.from}>${clrChange.to}`
  if (duotone) key += `|${duotone[0]}|${duotone[1]}`
  return key
}

/** Apply blip pixel effects in PowerPoint's order: clrChange first, then duotone. */
export function processedImage(
  img: HTMLImageElement,
  dataUrl: string,
  clrChange?: ClrChange,
  duotone?: [string, string],
): CanvasImageSource {
  let src: HTMLImageElement | HTMLCanvasElement = img
  if (clrChange) src = clrChangeImage(src, `${dataUrl}|cc`, clrChange.from, clrChange.to)
  if (duotone)
    src = duotoneImage(src, processedImageKey(dataUrl, clrChange), duotone[0], duotone[1])
  return src
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

/** Picture srcRect crop ratios → Konva Image crop (source-image pixel coordinates). */
export function cropToKonva(
  pic: PictureRenderNode,
  img: HTMLImageElement | undefined,
): { crop?: { x: number; y: number; width: number; height: number } } {
  const sr = pic.srcRect
  if (!sr || !img || !img.width || !img.height) return {}
  const x = Math.max(sr.l, 0) * img.width
  const y = Math.max(sr.t, 0) * img.height
  const w = Math.max(img.width * (1 - Math.max(sr.l, 0) - Math.max(sr.r, 0)), 1)
  const h = Math.max(img.height * (1 - Math.max(sr.t, 0) - Math.max(sr.b, 0)), 1)
  return { crop: { x, y, width: w, height: h } }
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
  /** Vertical latin word: rotated 90° clockwise (x/y is the rotation anchor) */
  rotation?: number
  /** Text highlight (<a:rPr><a:highlight>): background rect drawn behind the run, covering the line box */
  highlight?: { x: number; y: number; w: number; h: number; color: string }
  shadowColor?: string
  shadowBlur?: number
  shadowOffsetX?: number
  shadowOffsetY?: number
  shadowEnabled?: boolean
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

export function glyphToDraw(run: GlyphRun): GlyphDraw {
  const styleParts: string[] = []
  if (run.bold && !NO_SYNTHETIC_BOLD.has(run.fontFamily.normalize('NFKC').toLowerCase()))
    styleParts.push('bold')
  if (run.italic) styleParts.push('italic')
  return {
    text: run.text,
    x: run.x,
    // Konva Text positions by top; its browser-backed text baseline is closest to 0.8em.
    y: run.baselineY - run.fontSizePx * 0.8,
    fontSize: run.fontSizePx,
    fontFamily: displayFontFamily(run.fontFamily),
    fill: normalizeColor(run.color),
    fontStyle: styleParts.join(' ') || 'normal',
    textDecoration: [run.underline ? 'underline' : '', run.strike ? 'line-through' : '']
      .filter(Boolean)
      .join(' '),
    ...((run.letterSpacingPx ?? 0) + (run.justifyExtraPx ?? 0)
      ? { letterSpacing: (run.letterSpacingPx ?? 0) + (run.justifyExtraPx ?? 0) }
      : {}),
    ...(run.outline
      ? {
          stroke: normalizeColor(run.outline.color),
          strokeWidth: Math.max(run.outline.widthPx, 0.5),
          fillAfterStrokeEnabled: true,
        }
      : {}),
    ...(run.rtl ? { direction: 'rtl' as const } : {}),
    ...(run.rotate90 ? { rotation: 90 } : {}),
    ...(run.shadow
      ? {
          shadowColor: normalizeColor(run.shadow.color),
          shadowBlur: run.shadow.blurPx,
          shadowOffsetX: run.shadow.offsetX,
          shadowOffsetY: run.shadow.offsetY,
          shadowEnabled: true,
        }
      : {}),
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
  return c.startsWith('#') ? c : `#${c}`
}

export function isEditableText(node: RenderNode): node is ShapeRenderNode {
  return (node.type === 'text' || node.type === 'shape') && !!(node as ShapeRenderNode).text
}

/** Polyline smooth → Konva Line tension (0 = polyline, 0.4 ≈ PPT smooth curve). */
export function smoothTension(smooth: boolean | undefined): number {
  return smooth ? 0.4 : 0
}
