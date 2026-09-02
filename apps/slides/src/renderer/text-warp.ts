/**
 * WordArt preset text warp (<a:bodyPr><a:prstTxWarp>): bend the laid-out glyphs along
 * the preset's envelope, per character.
 *
 * Model (an approximation of PowerPoint's glyph-outline warping):
 *   1. The text block scales to fill the shape's text area (WordArt bodies stretch
 *      their text to the frame; the plain layout is much smaller than the reference).
 *   2. Each character maps by its normalized center u ∈ [0,1] across the block:
 *      - baseline presets (waves, arches, cans) move the character along a curve,
 *        arches also rotate it to the tangent;
 *      - envelope presets (inflate/triangle) squeeze it vertically between a top(u)
 *        and bottom(u) curve via scaleY;
 *      - ring presets place characters around an ellipse.
 *   Whole-glyph transforms only — glyph outlines are not themselves bent, which reads
 *   correctly at slide sizes and degrades gracefully.
 *
 * Every glyph is emitted center-anchored: position = the character center, offset =
 * half the local box, so rotation and scaling spin around the center.
 *
 * Unsupported presets return null and the caller keeps the straight layout.
 */
import type { GlyphDraw } from './konva-adapter'

/** Canvas measurer on the same font stack Konva draws with (module-level cached ctx). */
let mctx: CanvasRenderingContext2D | null = null
export function measureGlyph(text: string, g: GlyphDraw): number {
  mctx ??= document.createElement('canvas').getContext('2d')
  if (!mctx) return text.length * g.fontSize * 0.6
  mctx.font = `${g.fontStyle && g.fontStyle !== 'normal' ? g.fontStyle + ' ' : ''}${g.fontSize}px ${g.fontFamily}`
  return mctx.measureText(text).width
}

interface WarpSpec {
  /** Character-center y offset at u (px, + = down); amplitude a is already in px */
  baseline?: (u: number, a: number) => number
  /** Rotate baseline characters to the curve tangent (arches; waves stay upright) */
  tangent?: boolean
  /** Vertical envelope at u in box units: [top, bottom], 0 = box top, 1 = box bottom */
  envelope?: (u: number) => [number, number]
  /** Characters march around an ellipse, tops toward (in) or away from (out) the center */
  ring?: 'in' | 'out'
  /** Default primary adjust (1/100000 of the box height) */
  defAdj: number
}

const bump = (u: number) => 4 * u * (1 - u) // 0 at the edges, 1 mid-box
const vee = (u: number) => Math.abs(2 * u - 1) // 1 at the edges, 0 mid-box

const WARPS: Record<string, WarpSpec> = {
  // Waves: the line undulates, glyphs stay upright (PowerPoint reference)
  textWave1: { baseline: (u, a) => 1.5 * a * Math.sin(2 * Math.PI * u), defAdj: 12500 },
  textWave2: { baseline: (u, a) => -1.5 * a * Math.sin(2 * Math.PI * u), defAdj: 12500 },
  textDoubleWave1: { baseline: (u, a) => 1.5 * a * Math.sin(4 * Math.PI * u), defAdj: 6250 },
  textWave4: { baseline: (u, a) => -1.5 * a * Math.sin(4 * Math.PI * u), defAdj: 6250 },
  // Arches/curves: glyphs follow the curve tangent
  textArchUp: { baseline: (u, a) => -a * bump(u), tangent: true, defAdj: 25000 },
  textArchDown: { baseline: (u, a) => a * bump(u), tangent: true, defAdj: 25000 },
  textCurveUp: { baseline: (u, a) => -a * bump(u), tangent: true, defAdj: 45977 },
  textCurveDown: { baseline: (u, a) => a * bump(u), tangent: true, defAdj: 45977 },
  // Cans: both edges bow the same way; characters shift, height stays
  textCanUp: { baseline: (u, a) => -a * bump(u), defAdj: 18750 },
  textCanDown: { baseline: (u, a) => a * bump(u), defAdj: 18750 },
  // Envelope presets: characters stretch between the curves
  textInflate: { envelope: (u) => [0.2 - 0.2 * bump(u), 0.8 + 0.2 * bump(u)], defAdj: 12500 },
  textDeflate: { envelope: (u) => [0.28 * bump(u), 1 - 0.28 * bump(u)], defAdj: 12500 },
  textInflateTop: { envelope: (u) => [0.4 - 0.4 * bump(u), 1], defAdj: 12500 },
  textInflateBottom: { envelope: (u) => [0, 0.6 + 0.4 * bump(u)], defAdj: 12500 },
  textDeflateTop: { envelope: (u) => [0.45 * bump(u), 1], defAdj: 12500 },
  textDeflateBottom: { envelope: (u) => [0, 1 - 0.45 * bump(u)], defAdj: 12500 },
  textDeflateInflate: {
    envelope: (u) => {
      const h = 0.3 + 0.7 * vee(u) // pinched mid-box, full height at the edges
      return [0.5 - h / 2, 0.5 + h / 2]
    },
    defAdj: 12500,
  },
  // Triangles: one flat edge, the other a V/Λ
  textTriangle: { envelope: (u) => [0.6 - 0.6 * vee(u), 1], defAdj: 50000 },
  textTriangleInverted: { envelope: (u) => [0, 0.4 + 0.6 * vee(u)], defAdj: 50000 },
  // Rings/circle: not yet — the per-character ellipse walk needs its own calibration
  // pass; a straight fallback reads far better than a wrong ring.
}

/**
 * Split the runs into per-character glyphs, repacked contiguously by canvas-measured
 * widths, grouped into lines by their layout y. Layout x is only used for ordering:
 * mixing layout positions with canvas widths would tear word gaps open.
 */
function explode(
  glyphs: GlyphDraw[],
  measure: (text: string, g: GlyphDraw) => number,
): Array<Array<{ g: GlyphDraw; w: number; x: number }>> {
  const byLine = new Map<number, GlyphDraw[]>()
  for (const g of glyphs) {
    const key = Math.round(g.y)
    const arr = byLine.get(key)
    if (arr) arr.push(g)
    else byLine.set(key, [g])
  }
  const lines: Array<Array<{ g: GlyphDraw; w: number; x: number }>> = []
  for (const key of [...byLine.keys()].sort((a, b) => a - b)) {
    const runs = byLine.get(key)!.sort((a, b) => a.x - b.x)
    const line: Array<{ g: GlyphDraw; w: number; x: number }> = []
    let x = 0
    for (const g of runs) {
      // letterSpacing carries justify spread from the straight layout — the repacked
      // warp text must not inherit it. Reflections are dropped (a mirrored warp copy
      // would need its own transform pass; omitting reads better than misplacing).
      let runOff = 0
      for (const ch of [...g.text]) {
        const w = measure(ch, g)
        if (ch.trim()) {
          const cg: GlyphDraw = {
            ...g,
            text: ch,
            letterSpacing: 0,
            highlight: undefined,
            reflection: undefined,
          }

          // Run-local gradients: shift the ramp by the character's offset inside the
          // run so the letters together still paint one continuous gradient
          if (cg.fillLinearGradientStartPoint && cg.fillLinearGradientEndPoint) {
            cg.fillLinearGradientStartPoint = {
              x: cg.fillLinearGradientStartPoint.x - runOff,
              y: cg.fillLinearGradientStartPoint.y,
            }
            cg.fillLinearGradientEndPoint = {
              x: cg.fillLinearGradientEndPoint.x - runOff,
              y: cg.fillLinearGradientEndPoint.y,
            }
          }
          line.push({ g: cg, w, x })
        }
        x += w
        runOff += w
      }
    }
    if (line.length) lines.push(line)
  }
  return lines
}

/**
 * Warp a shape's glyphs. Returns null when the preset is unsupported (the caller keeps
 * the straight layout). `measure` must use the same font stack Konva draws with.
 */
export function warpGlyphs(
  glyphs: GlyphDraw[],
  boxW: number,
  boxH: number,
  warp: { prst: string; adj?: Record<string, number> },
  measure: (text: string, g: GlyphDraw) => number,
): GlyphDraw[] | null {
  const spec = WARPS[warp.prst]
  if (!spec || boxW <= 1 || boxH <= 1 || !glyphs.length) return null
  const lines = explode(glyphs, measure)
  if (!lines.length) return null
  const adjRaw = warp.adj?.adj ?? warp.adj?.adj1
  const a = ((adjRaw ?? spec.defAdj) / 100000) * boxH

  if (spec.ring) {
    // All characters march once around an ellipse centered in the box, starting at the
    // left (9 o'clock), clockwise over the top; rotation follows the tangent, flipped
    // 180° for the outside variant (reference: inside = upright at the top).
    const chars = lines.flat()
    const em = chars[0]!.g.fontSize
    const rx = Math.max(boxW / 2 - em / 2, 1)
    const ry = Math.max(boxH / 2 - em / 2, 1)
    const total = chars.reduce((s, c) => s + c.w, 0)
    // Scale glyphs so the text covers the circumference (rough ellipse perimeter)
    const perim = Math.PI * (3 * (rx + ry) - Math.sqrt((3 * rx + ry) * (rx + 3 * ry)))
    const scale = Math.min(Math.max(perim / total, 0.4), 2.2)
    let acc = 0
    const out: GlyphDraw[] = []
    for (const { g, w } of chars) {
      const u = (acc + w / 2) / total
      acc += w
      const th = Math.PI - 2 * Math.PI * u // left → top → right → bottom, clockwise
      const px = boxW / 2 + rx * Math.cos(th)
      const py = boxH / 2 - ry * Math.sin(th)
      const rot = 90 - (th * 180) / Math.PI + (spec.ring === 'out' ? 180 : 0)
      out.push({
        ...g,
        x: px,
        y: py,
        scaleX: scale,
        scaleY: scale,
        offsetX: w / 2,
        offsetY: (g.fontSize * 1.2) / 2,
        rotation: rot,
      })
    }
    return out
  }

  const nLines = lines.length
  const out: GlyphDraw[] = []
  for (const [li, line] of lines.entries()) {
    const lineW = Math.max(
      line.reduce((s, c) => Math.max(s, c.x + c.w), 0),
      1,
    )
    // Each line gets an equal horizontal band of the box
    const bandTop = (boxH / nLines) * li
    const bandH = boxH / nLines
    const localH = line[0]!.g.fontSize * 1.2
    // PowerPoint fits WordArt text uniformly (aspect kept) into the frame and centers
    // it; the warp curve then rides on top and may overflow the frame
    // 1.18: PowerPoint's fitted WordArt runs ~18% larger than a naive fit of the
    // canvas-measured width (calibrated against tdf114848 renders; text may overflow)
    const fit = Math.min((boxW / lineW) * 1.18, bandH / localH)
    const x0 = (boxW - lineW * fit) / 2
    for (const { g, w, x } of line) {
      const u = Math.min(Math.max((x + w / 2) / lineW, 0), 1)
      const cx = x0 + (x + w / 2) * fit
      if (spec.envelope) {
        const [t, b] = spec.envelope(u)
        const syChar = Math.max(((b - t) * bandH) / localH, 0.05)
        out.push({
          ...g,
          x: cx,
          y: bandTop + ((t + b) / 2) * bandH,
          scaleX: fit,
          scaleY: syChar,
          rotation: 0,
          offsetX: w / 2,
          offsetY: localH / 2,
        })
      } else {
        const yCenter = bandTop + bandH / 2 + spec.baseline!(u, a)
        const rot = spec.tangent
          ? (Math.atan2(
              spec.baseline!(Math.min(u + 0.02, 1), a) - spec.baseline!(Math.max(u - 0.02, 0), a),
              0.04 * boxW,
            ) *
              180) /
            Math.PI
          : 0
        out.push({
          ...g,
          x: cx,
          y: yCenter,
          scaleX: fit,
          scaleY: fit,
          rotation: rot,
          offsetX: w / 2,
          offsetY: localH / 2,
        })
      }
    }
  }
  return out
}
