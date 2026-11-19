/**
 * Table cell <a:cell3D> bevel → flat shading bands (PowerPoint-measured, 96dpi probe deck).
 *
 * PowerPoint draws a bevelled cell as a darkened face (fill × 0.85) with four mitred bands
 * of the bevel width along the edges. With the default flood rig lit from the top the left
 * band is a specular highlight (≈ ×2.6, clamps to white), the top a mild highlight
 * (≈ ×1.35 at 40 % depth), the right a deep shadow (×0.40) and the bottom a mild one
 * (×0.89); the whole pattern rotates with the light-rig direction (dir="r" moves the
 * highlight to the top, the deep shadow to the bottom). Circle-type bevels hold the band
 * value for the outer 45 % and fade into the face; the "angle" preset stays flat.
 */

export type BevelEdge = 't' | 'r' | 'b' | 'l'

export interface CellBevelRender {
  /** Band width (px) along each edge */
  widthPx: number
  /** Gradient stops per edge; pos 0 = outer edge, 1 = inner edge */
  edges: Record<BevelEdge, Array<{ pos: number; color: string }>>
}

const FACE_FACTOR = 0.85
/** Light-rig dir → azimuth (degrees clockwise from top) */
const DIR_DEG: Record<string, number> = {
  t: 0,
  tr: 45,
  r: 90,
  br: 135,
  b: 180,
  bl: 225,
  l: 270,
  tl: 315,
}
/** Flood rig lights the face from 60° left of its nominal direction (probe: dir="t" puts
 *  the specular on the left band and the deep shadow on the right) */
const LIGHT_SKEW_DEG = -60
const EDGE_NORMAL_DEG: Record<BevelEdge, number> = { t: 0, r: 90, b: 180, l: 270 }
/** Measured peak band factor (relative to the face) by lit value cos(normal, light) */
const AMPLITUDE: Array<[number, number]> = [
  [-1, -0.62],
  [-0.87, -0.6],
  [-0.5, -0.11],
  [0, 0],
  [0.5, 0.35],
  [0.87, 1.6],
  [1, 1.7],
]
/** Flat presets keep the band value to the inner edge; the rest fade over the inner half */
const FLAT_PRESETS = new Set(['angle', 'hardEdge', 'slope'])

function parseHex(hex: string): [number, number, number] {
  const h = hex.replace('#', '')
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]
}

/** #RRGGBBAA alpha suffix (translucent fills keep it on the face and the bands) */
function alphaOf(hex: string): string {
  const h = hex.replace('#', '')
  return h.length === 8 ? h.slice(6, 8).toUpperCase() : ''
}

function toHex(rgb: [number, number, number], alpha: string): string {
  return (
    '#' +
    rgb
      .map((v) =>
        Math.max(0, Math.min(255, Math.round(v)))
          .toString(16)
          .padStart(2, '0'),
      )
      .join('')
      .toUpperCase() +
    alpha
  )
}

function scale(rgb: [number, number, number], k: number, alpha: string): string {
  return toHex([rgb[0] * k, rgb[1] * k, rgb[2] * k], alpha)
}

function amplitude(lit: number): number {
  for (let i = 1; i < AMPLITUDE.length; i++) {
    const [x0, y0] = AMPLITUDE[i - 1]!
    const [x1, y1] = AMPLITUDE[i]!
    if (lit <= x1) return y0 + ((lit - x0) / (x1 - x0)) * (y1 - y0)
  }
  return AMPLITUDE[AMPLITUDE.length - 1]![1]
}

/** Darkened face color of a bevelled cell (#RRGGBB). */
export function bevelFaceColor(fill: string): string {
  return scale(parseHex(fill), FACE_FACTOR, alphaOf(fill))
}

export function buildCellBevel(
  fill: string,
  widthPx: number,
  preset: string | undefined,
  lightDir: string | undefined,
): CellBevelRender {
  const face = parseHex(bevelFaceColor(fill))
  const alpha = alphaOf(fill)
  const light = (DIR_DEG[lightDir ?? 't'] ?? 0) + LIGHT_SKEW_DEG
  const flat = FLAT_PRESETS.has(preset ?? 'circle')
  const edges = {} as CellBevelRender['edges']
  for (const e of ['t', 'r', 'b', 'l'] as const) {
    const lit = Math.cos(((EDGE_NORMAL_DEG[e] - light) * Math.PI) / 180)
    const peak = 1 + amplitude(lit)
    edges[e] = flat
      ? [
          { pos: 0, color: scale(face, peak, alpha) },
          { pos: 0.95, color: scale(face, peak, alpha) },
          { pos: 1, color: toHex(face, alpha) },
        ]
      : [
          { pos: 0, color: scale(face, peak, alpha) },
          { pos: 0.45, color: scale(face, peak, alpha) },
          { pos: 0.95, color: toHex(face, alpha) },
          { pos: 1, color: toHex(face, alpha) },
        ]
  }
  return { widthPx, edges }
}
