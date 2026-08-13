/**
 * Pure logic for animation playback — step grouping + per-shape visual state at any moment.
 *
 * A step = the group of animations advanced together by one trigger (a click, or automatic
 * on page entry): onClick opens a new step; withPrev joins the current step (start = previous
 * animation's start + delay); afterPrev joins the current step (start = previous animation's
 * end + delay). If the first animation is not onClick, step 0 is marked auto and plays
 * automatically when the page is entered.
 *
 * The visual state NodeAnimState is applied by the show/preview layer on a Group wrapped
 * around the Konva node (opacity/scale/rotation/offset/wipe clip) — a CSS approximation
 * matching one-to-one the OOXML effects the engine writes into the pptx.
 */
import type { AnimEffectKind, AnimationItem } from '../shared/ipc'

export interface TimedAnim {
  item: AnimationItem
  /** Start/end within the step (ms) */
  startMs: number
  endMs: number
}

export interface AnimStep {
  /** true = auto-plays on page entry (when the first animation is withPrev/afterPrev) */
  auto: boolean
  items: TimedAnim[]
  totalMs: number
}

/** Effect category (entrance/emphasis/exit/motion path). */
export function animClassOf(effect: AnimEffectKind): 'entrance' | 'emphasis' | 'exit' | 'path' {
  switch (effect) {
    case 'appear':
    case 'fade':
    case 'flyIn':
    case 'wipe':
    case 'wipeDown':
    case 'splitIn':
    case 'bounce':
    case 'flipIn':
    case 'zoom':
      return 'entrance'
    case 'pulse':
    case 'spin':
    case 'grow':
    case 'teeter':
      return 'emphasis'
    case 'motionPath':
      return 'path'
    default:
      return 'exit'
  }
}

/** Group the animation list into playback steps (same grouping rules the engine uses when writing <p:timing>). */
export function buildSteps(items: AnimationItem[]): AnimStep[] {
  const steps: AnimStep[] = []
  let prevStart = 0
  let prevEnd = 0
  for (const it of items) {
    if (it.trigger === 'onClick' || steps.length === 0) {
      steps.push({ auto: it.trigger !== 'onClick', items: [], totalMs: 0 })
      prevStart = 0
      prevEnd = 0
    }
    const base = it.trigger === 'afterPrev' ? prevEnd : it.trigger === 'withPrev' ? prevStart : 0
    const start = base + Math.max(0, it.delayMs)
    const end = start + Math.max(1, it.durationMs)
    const step = steps[steps.length - 1]!
    step.items.push({ item: it, startMs: start, endMs: end })
    step.totalMs = Math.max(step.totalMs, end)
    prevStart = start
    prevEnd = end
  }
  return steps
}

// NUL never appears in sourceId, so use it as the paragraph-key separator
const SEP = '\u0000'

/**
 * Key of the state Map: whole-shape animation = sourceId; paragraph animation =
 * sourceId + SEP + paragraph index, so each paragraph of the same shape enters/exits
 * independently without overwriting the others.
 */
export function animStateKey(sourceId: string, paragraph?: number): string {
  return paragraph == null ? sourceId : `${sourceId}${SEP}${paragraph}`
}

/** Paragraph key -> { sourceId, para }; returns null for whole-shape keys. */
export function parseParaStateKey(key: string): { sourceId: string; para: number } | null {
  const at = key.indexOf(SEP)
  return at < 0 ? null : { sourceId: key.slice(0, at), para: Number(key.slice(at + 1)) }
}

/** Visual state of a node at a given moment (defaults = displayed normally). */
export interface NodeAnimState {
  hidden: boolean
  opacity: number
  scale: number
  /** Extra horizontal scale (for flip-style entrances; render layer uses scaleX = scale * scaleX) */
  scaleX: number
  rotationDeg: number
  /** Extra offset (px, canvas coordinates) */
  dx: number
  dy: number
  /** Clip reveal: t=visible ratio 0..1; mode=visible area anchored to bottom (btm)/top (top)/center (mid); null = no clipping */
  clip: { t: number; mode: 'btm' | 'top' | 'mid' } | null
}

const NORMAL: NodeAnimState = {
  hidden: false,
  opacity: 1,
  scale: 1,
  scaleX: 1,
  rotationDeg: 0,
  dx: 0,
  dy: 0,
  clip: null,
}

function easeOut(t: number): number {
  return 1 - (1 - t) * (1 - t)
}

/** Bounce easing (standard easeOutBounce): decaying rebound after landing. */
function easeOutBounce(t: number): number {
  const n = 7.5625
  const d = 2.75
  if (t < 1 / d) return n * t * t
  if (t < 2 / d) return n * (t -= 1.5 / d) * t + 0.75
  if (t < 2.5 / d) return n * (t -= 2.25 / d) * t + 0.9375
  return n * (t -= 2.625 / d) * t + 0.984375
}

// ── Motion-path sampling (SVG subset M/L/C/Z, coordinates 0..1 relative to slide size) ──────

/** Parse a path into polyline sample points (cubic beziers subdivided into 16 segments); on parse failure return a single origin point. */
export function samplePathPoints(path: string): Array<{ x: number; y: number }> {
  const tokens = path.match(/[MLCZ]|-?(?:\d*\.\d+|\d+)(?:e-?\d+)?/gi) ?? []
  const pts: Array<{ x: number; y: number }> = []
  let i = 0
  let cx = 0
  let cy = 0
  let sx = 0
  let sy = 0
  let cmd = ''
  const num = () => Number(tokens[i++] ?? 0)
  while (i < tokens.length) {
    const t = tokens[i]!
    if (/^[MLCZ]$/i.test(t)) {
      cmd = t.toUpperCase()
      i++
    } else if (cmd === 'M') {
      cmd = 'L' // Consecutive coordinates after M are treated as L, per SVG convention
    }
    if (cmd === 'M') {
      cx = num()
      cy = num()
      sx = cx
      sy = cy
      pts.push({ x: cx, y: cy })
    } else if (cmd === 'L') {
      cx = num()
      cy = num()
      pts.push({ x: cx, y: cy })
    } else if (cmd === 'C') {
      const x1 = num()
      const y1 = num()
      const x2 = num()
      const y2 = num()
      const x = num()
      const y = num()
      for (let k = 1; k <= 16; k++) {
        const u = k / 16
        const v = 1 - u
        pts.push({
          x: v * v * v * cx + 3 * v * v * u * x1 + 3 * v * u * u * x2 + u * u * u * x,
          y: v * v * v * cy + 3 * v * v * u * y1 + 3 * v * u * u * y2 + u * u * u * y,
        })
      }
      cx = x
      cy = y
    } else if (cmd === 'Z') {
      pts.push({ x: sx, y: sy })
      cx = sx
      cy = sy
    } else {
      i++ // Skip unrecognized tokens (such as OOXML's E end marker)
    }
  }
  return pts.length > 0 ? pts : [{ x: 0, y: 0 }]
}

/** Point on the path at progress q (0..1), by arc length. */
export function pointAtPath(
  pts: Array<{ x: number; y: number }>,
  q: number,
): { x: number; y: number } {
  if (pts.length === 1) return pts[0]!
  const lens: number[] = [0]
  for (let i = 1; i < pts.length; i++) {
    lens.push(lens[i - 1]! + Math.hypot(pts[i]!.x - pts[i - 1]!.x, pts[i]!.y - pts[i - 1]!.y))
  }
  const total = lens[lens.length - 1]!
  if (total === 0) return pts[0]!
  const target = Math.min(1, Math.max(0, q)) * total
  let i = 1
  while (i < lens.length - 1 && lens[i]! < target) i++
  const seg = lens[i]! - lens[i - 1]!
  const u = seg > 0 ? (target - lens[i - 1]!) / seg : 0
  return {
    x: pts[i - 1]!.x + (pts[i]!.x - pts[i - 1]!.x) * u,
    y: pts[i - 1]!.y + (pts[i]!.y - pts[i - 1]!.y) * u,
  }
}

/** Small cache of parsed paths (sampled every frame during a show) */
const pathCache = new Map<string, Array<{ x: number; y: number }>>()
function cachedPathPoints(path: string): Array<{ x: number; y: number }> {
  let pts = pathCache.get(path)
  if (!pts) {
    pts = samplePathPoints(path)
    if (pathCache.size > 64) pathCache.clear()
    pathCache.set(path, pts)
  }
  return pts
}

/** Write the state of a single animation at progress p (0..1) into st; p=1 means finished. */
function applyEffect(
  st: NodeAnimState,
  item: AnimationItem,
  p: number,
  canvasW: number,
  canvasH: number,
): void {
  const q = easeOut(Math.min(1, Math.max(0, p)))
  switch (item.effect) {
    case 'appear':
      st.hidden = false
      break
    case 'fade':
      st.hidden = false
      st.opacity = q
      break
    case 'flyIn':
      st.hidden = false
      st.dy = (1 - q) * canvasH
      break
    case 'wipe':
      st.hidden = false
      st.clip = p >= 1 ? null : { t: q, mode: 'btm' }
      break
    case 'wipeDown':
      st.hidden = false
      st.clip = p >= 1 ? null : { t: q, mode: 'top' }
      break
    case 'splitIn':
      st.hidden = false
      st.clip = p >= 1 ? null : { t: q, mode: 'mid' }
      break
    case 'bounce':
      // Drop from above + decaying bounce (written as (b-1)*positive to avoid -0 at the end)
      st.hidden = false
      st.dy = (easeOutBounce(Math.min(1, Math.max(0, p))) - 1) * 0.25 * canvasH
      break
    case 'flipIn':
      // Flip-style (2D approximation of rotateY: unfolds horizontally from 0)
      st.hidden = false
      st.scaleX = Math.max(0.01, q)
      st.opacity = Math.min(1, q * 2)
      break
    case 'zoom':
      st.hidden = false
      st.scale = q
      st.opacity = Math.min(1, q * 2)
      break
    case 'pulse': {
      const s = p >= 1 ? 0 : Math.sin(Math.PI * p)
      st.scale = 1 + 0.08 * s
      break
    }
    case 'spin':
      st.rotationDeg = p >= 1 ? 0 : 360 * q
      break
    case 'grow':
      st.scale = 1 + 0.5 * q
      break
    case 'teeter': {
      // Seesaw: swings left/right with decay, settling back in place
      const s = p >= 1 ? 0 : Math.sin(3 * Math.PI * p) * (1 - p)
      st.rotationDeg = 9 * s
      break
    }
    case 'disappear':
      if (p >= 1) st.hidden = true
      break
    case 'fadeOut':
      st.opacity = 1 - q
      if (p >= 1) st.hidden = true
      break
    case 'flyOut':
      st.dy = q * canvasH
      if (p >= 1) st.hidden = true
      break
    case 'wipeOut':
      st.clip = p >= 1 ? null : { t: 1 - q, mode: 'btm' }
      if (p >= 1) st.hidden = true
      break
    case 'shrink':
      // Shrink and rotate on exit
      st.scale = Math.max(0.01, 1 - q)
      st.rotationDeg = -90 * q
      st.opacity = 1 - q
      if (p >= 1) st.hidden = true
      break
    case 'zoomOut':
      st.scale = Math.max(0.01, 1 - q)
      st.opacity = 1 - q
      if (p >= 1) st.hidden = true
      break
    case 'motionPath': {
      // Move along path: offset = sample point - path start (relative to slide size); stops at the end when done
      st.hidden = false
      const pts = cachedPathPoints(item.motionPath || 'M 0 0 L 0.25 0')
      const pt = pointAtPath(pts, q)
      st.dx = (pt.x - pts[0]!.x) * canvasW
      st.dy = (pt.y - pts[0]!.y) * canvasH
      break
    }
  }
}

/**
 * Compute the state of every animated node when "played steps are done and the current step
 * has been running for activeMs". activeMs=null means no step is playing (resting state
 * after the played steps). Returns Map<animStateKey, NodeAnimState>; targets without
 * animations are not in the Map (displayed normally).
 */
export function computeNodeStates(
  steps: AnimStep[],
  played: number,
  activeMs: number | null,
  canvasH: number,
  // Legacy call sites pass only height: estimate width as 16:9 (currently only motion-path horizontal offsets use it)
  canvasW: number = (canvasH * 16) / 9,
): Map<string, NodeAnimState> {
  const states = new Map<string, NodeAnimState>()
  // Initial visibility: if the target's (shape or one paragraph) first animation is an entrance, start hidden until it plays
  for (const step of steps) {
    for (const t of step.items) {
      const id = animStateKey(t.item.sourceId, t.item.paragraph)
      if (!states.has(id)) {
        states.set(id, { ...NORMAL, hidden: animClassOf(t.item.effect) === 'entrance' })
      }
    }
  }

  steps.forEach((step, si) => {
    for (const t of step.items) {
      const st = states.get(animStateKey(t.item.sourceId, t.item.paragraph))!
      const cls = animClassOf(t.item.effect)
      // Progress of this animation: played steps = 1; current step by time; not started = null (state unchanged)
      let p: number | null
      if (si < played) p = 1
      else if (si === played && activeMs != null) {
        p =
          activeMs <= t.startMs ? null : Math.min(1, (activeMs - t.startMs) / (t.endMs - t.startMs))
      } else p = null
      if (p == null) continue
      // Playing/finished: reset to normal first, then apply this effect (later animations on the same node override earlier ones)
      if (cls === 'entrance' || p >= 1) {
        Object.assign(st, NORMAL, { hidden: st.hidden })
      }
      applyEffect(st, t.item, p, canvasW, canvasH)
    }
  })
  return states
}
