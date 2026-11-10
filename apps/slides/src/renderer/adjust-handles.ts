/**
 * PowerPoint-style yellow adjust handles: per-preset specs mapping a shape's
 * avLst adjust values to draggable points in box-local px and back.
 *
 * Units follow OOXML: linear adjusts are 1/1000 % (100000 = 100%) of the
 * shortest side unless a spec says otherwise; angular adjusts are 1/60000°.
 * A handle may drive several adjust names at once (callout tips drive x+y).
 *
 * Positions/inverses mirror the formulas in pptx-render/preset-geometry.ts —
 * when a geometry there is an approximation, the handle approximates the same
 * way, so what you drag is what you see. Not covered (no handles): gear6/9,
 * mathNotEqual, circularArrow (approximated multi-adj geometries where a
 * faithful inverse doesn't exist).
 */
import type { ShapeRenderNode } from '@genoffice/pptx-render'

/** Raw adjust getter: current value or the preset default. */
type AdjVal = (name: string, def: number) => number

export interface AdjustHandleSpec {
  /** avLst names this handle drives (defaults seed the full rewrite) */
  keys: ReadonlyArray<{ name: string; def: number }>
  /** box-local handle position */
  pos(w: number, h: number, val: AdjVal): { x: number; y: number }
  /** box-local pointer → the driven entries (clamped, rounded) */
  values(w: number, h: number, x: number, y: number, val: AdjVal): Record<string, number>
}

const clampRound = (v: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, Math.round(v)))
const ss = (w: number, h: number) => Math.min(w, h)

/** Single-key spec builder. */
const one = (
  name: string,
  def: number,
  pos: AdjustHandleSpec['pos'],
  value: (w: number, h: number, x: number, y: number, val: AdjVal) => number,
): AdjustHandleSpec => ({
  keys: [{ name, def }],
  pos,
  values: (w, h, x, y, val) => ({ [name]: value(w, h, x, y, val) }),
})

/** Fraction-of-shortest-side handle riding a horizontal edge. */
const ssEdge = (
  name: string,
  def: number,
  max: number,
  opts: { fromRight?: boolean; bottom?: boolean } = {},
): AdjustHandleSpec =>
  one(
    name,
    def,
    (w, h, val) => {
      const d = (ss(w, h) * val(name, def)) / 100000
      return { x: opts.fromRight ? w - d : d, y: opts.bottom ? h : 0 }
    },
    (w, h, x) => clampRound(((opts.fromRight ? w - x : x) / ss(w, h)) * 100000, 0, max),
  )

/** Angular adjust riding the ellipse rim (pie/arc/chord; y-down clockwise degrees). */
const angleRim = (name: string, def: number): AdjustHandleSpec =>
  one(
    name,
    def,
    (w, h, val) => {
      const rad = (val(name, def) / 60000) * (Math.PI / 180)
      return { x: w / 2 + (w / 2) * Math.cos(rad), y: h / 2 + (h / 2) * Math.sin(rad) }
    },
    (w, h, x, y) => {
      const deg = Math.atan2((y - h / 2) / (h / 2), (x - w / 2) / (w / 2)) * (180 / Math.PI)
      return Math.round(((deg + 360) % 360) * 60000)
    },
  )

/** Ring-thickness handle on the horizontal center line (donut/blockArc/noSmoking). */
const ringThickness = (name: string, def: number, max = 50000): AdjustHandleSpec =>
  one(
    name,
    def,
    (w, h, val) => ({ x: (ss(w, h) * val(name, def)) / 100000, y: h / 2 }),
    (w, h, x) => clampRound((x / ss(w, h)) * 100000, 0, max),
  )

/** Star inner-radius handle: rides the first inner vertex; radial-ratio inverse. */
const starInner = (def: number): AdjustHandleSpec =>
  one(
    'adj',
    def,
    (w, h, val) => {
      // starPoints: inner factor = min(adj/100000 * 2, 1) of the half-extent
      const f = Math.min((val('adj', def) / 100000) * 2, 1)
      const ang = -Math.PI / 2 + Math.PI / 5 // representative inner vertex direction
      return { x: w / 2 + Math.cos(ang) * (w / 2) * f, y: h / 2 + Math.sin(ang) * (h / 2) * f }
    },
    (w, h, x, y) => {
      const r = Math.hypot((x - w / 2) / (w / 2), (y - h / 2) / (h / 2))
      return clampRound((r / 2) * 100000, 0, 50000)
    },
  )

/** Callout tail tip: one handle drives adj1 (x, width fraction) + adj2 (y, height fraction). */
const calloutTip = (): AdjustHandleSpec => ({
  keys: [
    { name: 'adj1', def: -20833 },
    { name: 'adj2', def: 62500 },
  ],
  pos: (w, h, val) => ({
    x: w / 2 + (w * val('adj1', -20833)) / 100000,
    y: h / 2 + (h * val('adj2', 62500)) / 100000,
  }),
  values: (w, h, x, y) => ({
    adj1: clampRound(((x - w / 2) / w) * 100000, -200000, 200000),
    adj2: clampRound(((y - h / 2) / h) * 100000, -200000, 200000),
  }),
})

/** Straight-arrow pair: adj1 = shaft thickness (axis fraction), adj2 = head length (ss fraction). */
function arrowSpecs(dir: 'right' | 'left' | 'up' | 'down' | 'leftRight' | 'upDown') {
  const horizontal = dir === 'right' || dir === 'left' || dir === 'leftRight'
  const thickness = one(
    'adj1',
    50000,
    (w, h, val) => {
      const a = val('adj1', 50000) / 100000
      if (horizontal) {
        const x = dir === 'left' ? w : dir === 'leftRight' ? w / 2 : 0
        return { x, y: (h - h * a) / 2 }
      }
      const y = dir === 'up' ? h : dir === 'upDown' ? h / 2 : 0
      return { x: (w - w * a) / 2, y }
    },
    (w, h, x, y) =>
      clampRound(horizontal ? ((h - 2 * y) / h) * 100000 : ((w - 2 * x) / w) * 100000, 0, 100000),
  )
  const head = one(
    'adj2',
    50000,
    (w, h, val) => {
      const d = (ss(w, h) * val('adj2', 50000)) / 100000
      switch (dir) {
        case 'right':
        case 'leftRight':
          return { x: w - Math.min(dir === 'leftRight' ? w / 2 : w, d), y: 0 }
        case 'left':
          return { x: Math.min(w, d), y: 0 }
        case 'up':
          return { x: 0, y: Math.min(h, d) }
        case 'down':
          return { x: 0, y: h - Math.min(h, d) }
        case 'upDown':
          return { x: 0, y: Math.min(h / 2, d) }
      }
    },
    (w, h, x, y) => {
      const d =
        dir === 'right' || dir === 'leftRight'
          ? w - x
          : dir === 'left'
            ? x
            : dir === 'down'
              ? h - y
              : y
      return clampRound((d / ss(w, h)) * 100000, 0, 100000)
    },
  )
  return [thickness, head]
}

export const ADJUST_HANDLE_SPECS: Record<string, AdjustHandleSpec[]> = {
  // ── rectangle corner family ──────────────────────────────────────────────
  roundRect: [ssEdge('adj', 16667, 50000)],
  round1Rect: [ssEdge('adj', 16667, 50000, { fromRight: true })],
  round2SameRect: [ssEdge('adj1', 16667, 50000), ssEdge('adj2', 0, 50000, { bottom: true })],
  round2DiagRect: [ssEdge('adj1', 16667, 50000), ssEdge('adj2', 0, 50000, { fromRight: true })],
  snipRoundRect: [ssEdge('adj1', 16667, 50000), ssEdge('adj2', 16667, 50000, { fromRight: true })],
  snip1Rect: [ssEdge('adj', 16667, 50000, { fromRight: true })],
  snip2SameRect: [ssEdge('adj1', 16667, 50000), ssEdge('adj2', 0, 50000, { bottom: true })],
  snip2DiagRect: [ssEdge('adj1', 0, 50000), ssEdge('adj2', 16667, 50000, { fromRight: true })],
  plaque: [ssEdge('adj', 16667, 50000)],

  // ── polygons with an edge inset ──────────────────────────────────────────
  parallelogram: [ssEdge('adj', 25000, 100000)],
  // ECMA caps trapezoid at maxAdj 50000 — beyond it the top edge crosses itself
  trapezoid: [ssEdge('adj', 25000, 50000)],
  hexagon: [ssEdge('adj', 25000, 50000)],
  octagon: [ssEdge('adj', 29289, 50000)],
  plus: [ssEdge('adj', 25000, 50000)],
  chevron: [ssEdge('adj', 50000, 100000, { fromRight: true })],
  homePlate: [ssEdge('adj', 50000, 100000, { fromRight: true })],
  triangle: [
    one(
      'adj',
      50000,
      (w, _h, val) => ({ x: (w * val('adj', 50000)) / 100000, y: 0 }),
      (w, _h, x) => clampRound((x / w) * 100000, 0, 100000),
    ),
  ],
  diagStripe: [
    one(
      'adj',
      50000,
      (w, _h, val) => ({ x: (w * val('adj', 50000)) / 100000, y: 0 }),
      (w, _h, x) => clampRound((x / w) * 100000, 0, 100000),
    ),
  ],

  // ── frames / corners ─────────────────────────────────────────────────────
  halfFrame: [
    one(
      'adj1',
      33333,
      (w, h, val) => ({ x: 0, y: (ss(w, h) * val('adj1', 33333)) / 100000 }),
      (w, h, _x, y) => clampRound((y / ss(w, h)) * 100000, 0, 100000),
    ),
    ssEdge('adj2', 33333, 100000),
  ],
  corner: [
    one(
      'adj1',
      50000,
      (w, h, val) => ({ x: w, y: h - (ss(w, h) * val('adj1', 50000)) / 100000 }),
      (w, h, _x, y) => clampRound(((h - y) / ss(w, h)) * 100000, 0, 100000),
    ),
    ssEdge('adj2', 50000, 100000),
  ],
  frame: [
    one(
      'adj1',
      12500,
      (w, h, val) => {
        const t = (ss(w, h) * val('adj1', 12500)) / 100000
        return { x: t, y: t }
      },
      (w, h, x) => clampRound((x / ss(w, h)) * 100000, 0, 50000),
    ),
  ],
  bevel: [
    one(
      'adj',
      12500,
      (w, h, val) => {
        const t = (ss(w, h) * val('adj', 12500)) / 100000
        return { x: t, y: t }
      },
      (w, h, x) => clampRound((x / ss(w, h)) * 100000, 0, 50000),
    ),
  ],
  cube: [
    one(
      'adj',
      25000,
      (w, h, val) => {
        const d = (ss(w, h) * val('adj', 25000)) / 100000
        return { x: w - d, y: d }
      },
      (w, h, _x, y) => clampRound((y / ss(w, h)) * 100000, 0, 100000),
    ),
  ],
  foldedCorner: [
    one(
      'adj',
      16667,
      (w, h, val) => ({ x: w - (ss(w, h) * val('adj', 16667)) / 100000, y: h }),
      (w, h, x) => clampRound(((w - x) / ss(w, h)) * 100000, 0, 50000),
    ),
  ],
  can: [
    one(
      'adj',
      25000,
      (w, h, val) => ({ x: w / 2, y: (h * val('adj', 25000)) / 100000 }),
      (_w, h, _x, y) => clampRound((y / h) * 100000, 0, 50000),
    ),
  ],

  // ── angular (rim handles) ────────────────────────────────────────────────
  pie: [angleRim('adj1', 0), angleRim('adj2', 16200000)],
  arc: [angleRim('adj1', 16200000), angleRim('adj2', 0)],
  chord: [angleRim('adj1', 2700000), angleRim('adj2', 16200000)],
  blockArc: [angleRim('adj1', 10800000), angleRim('adj2', 0), ringThickness('adj3', 25000)],

  // ── rings / celestial ────────────────────────────────────────────────────
  donut: [ringThickness('adj', 25000)],
  noSmoking: [ringThickness('adj', 18750)],
  moon: [
    one(
      'adj',
      50000,
      (w, h, val) => ({ x: (w * val('adj', 50000)) / 100000, y: h / 2 }),
      (w, _h, x) => clampRound((x / w) * 100000, 0, 87500),
    ),
  ],
  sun: [
    one(
      'adj',
      25000,
      (w, h, val) => ({ x: w / 2 + (w * val('adj', 25000)) / 100000, y: h / 2 }),
      (w, _h, x) => clampRound(((x - w / 2) / w) * 100000, 12500, 46875),
    ),
  ],
  teardrop: [
    one(
      'adj',
      100000,
      (w, h, val) => {
        const a = val('adj', 100000) / 100000
        return { x: w / 2 + (w / 2) * a, y: h / 2 - (h / 2) * a }
      },
      (w, h, x, y) => {
        const a = ((x - w / 2) / (w / 2) - (y - h / 2) / (h / 2)) / 2
        return clampRound(a * 100000, 0, 200000)
      },
    ),
  ],
  smileyFace: [
    one(
      'adj',
      4653,
      (w, h, val) => ({
        x: w / 2,
        y: h * Math.min(Math.max(0.67 + (4 * val('adj', 4653)) / 100000, 0.4), 0.95),
      }),
      (_w, h, _x, y) => clampRound(((y / h - 0.67) / 4) * 100000, -4653, 4653),
    ),
  ],

  // ── waves ────────────────────────────────────────────────────────────────
  wave: [
    one(
      'adj1',
      12500,
      (w, h, val) => ({ x: 0, y: (h * Math.min(val('adj1', 12500), 25000)) / 100000 }),
      (_w, h, _x, y) => clampRound((y / h) * 100000, 0, 25000),
    ),
  ],
  doubleWave: [
    one(
      'adj1',
      6250,
      (w, h, val) => ({ x: 0, y: (h * Math.min(val('adj1', 6250), 20000)) / 100000 }),
      (_w, h, _x, y) => clampRound((y / h) * 100000, 0, 20000),
    ),
  ],

  // ── straight arrows ──────────────────────────────────────────────────────
  rightArrow: arrowSpecs('right'),
  notchedRightArrow: arrowSpecs('right'),
  stripedRightArrow: arrowSpecs('right'),
  leftArrow: arrowSpecs('left'),
  upArrow: arrowSpecs('up'),
  downArrow: arrowSpecs('down'),
  leftRightArrow: arrowSpecs('leftRight'),
  upDownArrow: arrowSpecs('upDown'),
  uturnArrow: [
    one(
      'adj1',
      25000,
      (w, h, val) => ({ x: (ss(w, h) * val('adj1', 25000)) / 100000, y: h }),
      (w, h, x) => clampRound((x / ss(w, h)) * 100000, 0, 50000),
    ),
  ],
  curvedRightArrow: [
    one(
      'adj1',
      25000,
      (w, h, val) => ({ x: w - (ss(w, h) * val('adj1', 25000)) / 100000, y: h / 2 }),
      (w, h, x) => clampRound(((w - x) / ss(w, h)) * 100000, 0, 50000),
    ),
  ],
  bentArrow: [
    ssEdge('adj1', 25000, 50000, { bottom: true }),
    one(
      'adj2',
      25000,
      (w, h, val) => {
        const t = (ss(w, h) * val('adj1', 25000)) / 100000
        const hw = (ss(w, h) * val('adj2', 25000)) / 100000
        const yc = Math.max(hw, t / 2)
        return { x: w - (ss(w, h) * val('adj3', 25000)) / 100000, y: yc - hw }
      },
      (w, h, _x, y, val) => {
        const t = (ss(w, h) * val('adj1', 25000)) / 100000
        const hw0 = (ss(w, h) * val('adj2', 25000)) / 100000
        const yc = Math.max(hw0, t / 2)
        return clampRound(((yc - y) / ss(w, h)) * 100000, 0, 50000)
      },
    ),
    one(
      'adj3',
      25000,
      (w, h, val) => {
        const t = (ss(w, h) * val('adj1', 25000)) / 100000
        const hw = (ss(w, h) * val('adj2', 25000)) / 100000
        return { x: w - (ss(w, h) * val('adj3', 25000)) / 100000, y: Math.max(hw, t / 2) + hw }
      },
      (w, h, x) => clampRound(((w - x) / ss(w, h)) * 100000, 0, 50000),
    ),
  ],
  quadArrow: [
    one(
      'adj1',
      22500,
      (w, h, val) => {
        const sw2 = (ss(w, h) * val('adj1', 22500)) / 200000
        return { x: w / 2 + sw2, y: h / 2 - sw2 }
      },
      (w, h, x) => clampRound((((x - w / 2) * 2) / ss(w, h)) * 100000, 0, 50000),
    ),
    {
      keys: [
        { name: 'adj2', def: 22500 },
        { name: 'adj3', def: 22500 },
      ],
      pos: (w, h, val) => ({
        x: w / 2 + (ss(w, h) * val('adj2', 22500)) / 100000,
        y: (ss(w, h) * val('adj3', 22500)) / 100000,
      }),
      values: (w, h, x, y) => ({
        adj2: clampRound(((x - w / 2) / ss(w, h)) * 100000, 0, 50000),
        adj3: clampRound((y / ss(w, h)) * 100000, 0, 50000),
      }),
    },
  ],

  // ── math ─────────────────────────────────────────────────────────────────
  mathPlus: [
    one(
      'adj1',
      23520,
      (w, h, val) => ({ x: w / 2 + (ss(w, h) * val('adj1', 23520)) / 100000, y: h / 2 }),
      (w, h, x) => clampRound(((x - w / 2) / ss(w, h)) * 100000, 0, 36745),
    ),
  ],
  mathEqual: [
    one(
      'adj1',
      23520,
      (w, h, val) => {
        const a2 = Math.min(val('adj2', 11760) / 100000, 1)
        const a1 = Math.min(val('adj1', 23520) / 100000, 0.36745)
        return { x: w / 2, y: h / 2 - (h * a2) / 2 - h * a1 }
      },
      (w, h, _x, y, val) => {
        const a2 = Math.min(val('adj2', 11760) / 100000, 1)
        return clampRound(((h / 2 - (h * a2) / 2 - y) / h) * 100000, 0, 36745)
      },
    ),
    one(
      'adj2',
      11760,
      (w, h, val) => ({ x: w / 2, y: h / 2 - (h * Math.min(val('adj2', 11760) / 100000, 1)) / 2 }),
      (_w, h, _x, y) => clampRound((((h / 2 - y) * 2) / h) * 100000, 0, 50000),
    ),
  ],

  // ── stars ────────────────────────────────────────────────────────────────
  star4: [starInner(12500)],
  star5: [starInner(19098)],
  star6: [starInner(28868)],
  star7: [starInner(34601)],
  star8: [starInner(37500)],
  star10: [starInner(42533)],
  star12: [starInner(37500)],
  star16: [starInner(37500)],
  star24: [starInner(37500)],
  star32: [starInner(37500)],

  // ── callouts (tail tip drives both adjusts) ──────────────────────────────
  wedgeRectCallout: [calloutTip()],
  wedgeEllipseCallout: [calloutTip()],
  cloudCallout: [calloutTip()],
  wedgeRoundRectCallout: [calloutTip(), ssEdge('adj3', 16667, 50000)],

  // ── brackets / braces ────────────────────────────────────────────────────
  leftBracket: [
    one(
      'adj',
      8333,
      (w, h, val) => ({ x: 0, y: Math.min(h / 2, (ss(w, h) * val('adj', 8333)) / 100000) }),
      (w, h, _x, y) => clampRound((y / ss(w, h)) * 100000, 0, 50000),
    ),
  ],
  rightBracket: [
    one(
      'adj',
      8333,
      (w, h, val) => ({ x: w, y: Math.min(h / 2, (ss(w, h) * val('adj', 8333)) / 100000) }),
      (w, h, _x, y) => clampRound((y / ss(w, h)) * 100000, 0, 50000),
    ),
  ],
  leftBrace: [
    one(
      'adj1',
      8333,
      (w, h, val) => ({ x: w / 2, y: Math.min(h / 4, (ss(w, h) * val('adj1', 8333)) / 100000) }),
      (w, h, _x, y) => clampRound((y / ss(w, h)) * 100000, 0, 50000),
    ),
    one(
      'adj2',
      50000,
      (w, h, val) => ({ x: 0, y: (h * val('adj2', 50000)) / 100000 }),
      (_w, h, _x, y) => clampRound((y / h) * 100000, 0, 100000),
    ),
  ],
  rightBrace: [
    one(
      'adj1',
      8333,
      (w, h, val) => ({ x: w / 2, y: Math.min(h / 4, (ss(w, h) * val('adj1', 8333)) / 100000) }),
      (w, h, _x, y) => clampRound((y / ss(w, h)) * 100000, 0, 50000),
    ),
    one(
      'adj2',
      50000,
      (w, h, val) => ({ x: w, y: (h * val('adj2', 50000)) / 100000 }),
      (_w, h, _x, y) => clampRound((y / h) * 100000, 0, 100000),
    ),
  ],
}

/** The node's draggable adjust handles (empty when the preset has none / isn't supported). */
export function adjustHandleSpecs(node: ShapeRenderNode): AdjustHandleSpec[] {
  if (node.line || !node.presetGeometry) return []
  return ADJUST_HANDLE_SPECS[node.presetGeometry] ?? []
}
