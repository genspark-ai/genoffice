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
    ...(props.fillPatternImage ? { fillPatternX: -w / 2, fillPatternY: -h / 2 } : {}),
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
          fillRadialGradientEndRadius: far * 1.8,
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
        return fill.mode === 'tile'
          ? { fillPatternImage: img }
          : {
              fillPatternImage: img,
              fillPatternRepeat: 'no-repeat',
              fillPatternScaleX: w / (img.width || w),
              fillPatternScaleY: h / (img.height || h),
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
    ctx.filter = `blur(${rad}px)`
    ctx.fillStyle = '#000'
    // Inset the blur mask by 1.5r: after the blur spreads outward it fades to transparent right at the original boundary
    ctx.fillRect(rad * 1.5, rad * 1.5, c.width - rad * 3, c.height - rad * 3)
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
  if (run.bold) styleParts.push('bold')
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
