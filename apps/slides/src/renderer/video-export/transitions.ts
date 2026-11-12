/**
 * Canvas compositing of the page transitions for one exported frame: the show plays them as
 * CSS animations on the entering page, this reproduces each with both pages in hand so the
 * leaving page can move too (push). Morph is not here — it is a Konva tween of the elements.
 */
import type { PlayedTransition } from './timeline'

type Easing = (t: number) => number

/** cubic-bezier(x1, y1, x2, y2) as CSS defines it, solved for y at a given x. */
export function cubicBezier(x1: number, y1: number, x2: number, y2: number): Easing {
  const sample = (a: number, b: number, t: number) =>
    ((1 - 3 * b + 3 * a) * t + (3 * b - 6 * a)) * t * t + 3 * a * t
  return (x) => {
    if (x <= 0) return 0
    if (x >= 1) return 1
    let lo = 0
    let hi = 1
    let t = x
    for (let i = 0; i < 24; i++) {
      const cx = sample(x1, x2, t)
      if (Math.abs(cx - x) < 1e-5) break
      if (cx < x) lo = t
      else hi = t
      t = (lo + hi) / 2
    }
    return sample(y1, y2, t)
  }
}

const EASE = cubicBezier(0.25, 0.1, 0.25, 1)
const EASE_OUT = cubicBezier(0, 0, 0.58, 1)

const EASING: Record<Exclude<PlayedTransition, 'morph'>, Easing> = {
  fade: EASE,
  push: EASE_OUT,
  wipe: EASE_OUT,
  split: EASE_OUT,
  circle: EASE_OUT,
  cover: EASE,
  pull: EASE,
  dissolve: (t) => Math.floor(Math.min(1, Math.max(0, t)) * 12) / 12,
  zoom: EASE,
}

const DISSOLVE_COLS = 32
const DISSOLVE_ROWS = 18

/** Fixed pseudo-random order in which dissolve blocks switch to the new page. */
const DISSOLVE_ORDER: number[] = (() => {
  const n = DISSOLVE_COLS * DISSOLVE_ROWS
  const order = Array.from({ length: n }, (_, i) => i)
  let seed = 0x9e3779b9
  for (let i = n - 1; i > 0; i--) {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0
    const j = seed % (i + 1)
    ;[order[i], order[j]] = [order[j]!, order[i]!]
  }
  return order
})()

/**
 * Draw the frame at raw progress u (0..1, linear in time) of `kind` between `prev` and `next`,
 * both already the output size.
 */
export function drawTransition(
  ctx: CanvasRenderingContext2D,
  prev: CanvasImageSource,
  next: CanvasImageSource,
  kind: Exclude<PlayedTransition, 'morph'>,
  u: number,
  w: number,
  h: number,
): void {
  const p = EASING[kind](Math.min(1, Math.max(0, u)))
  ctx.save()
  ctx.globalAlpha = 1
  switch (kind) {
    case 'fade':
      ctx.drawImage(prev, 0, 0)
      ctx.globalAlpha = p
      ctx.drawImage(next, 0, 0)
      break
    case 'push':
      ctx.drawImage(prev, 0, -p * h)
      ctx.drawImage(next, 0, (1 - p) * h)
      break
    case 'cover':
      ctx.drawImage(prev, 0, 0)
      ctx.drawImage(next, (1 - p) * w, 0)
      break
    case 'wipe':
      ctx.drawImage(prev, 0, 0)
      clipRect(ctx, 0, 0, p * w, h)
      ctx.drawImage(next, 0, 0)
      break
    case 'pull':
      ctx.drawImage(prev, 0, 0)
      clipRect(ctx, (1 - p) * w, 0, p * w, h)
      ctx.drawImage(next, 0, 0)
      break
    case 'split':
      ctx.drawImage(prev, 0, 0)
      clipRect(ctx, 0, (h / 2) * (1 - p), w, h * p)
      ctx.drawImage(next, 0, 0)
      break
    case 'circle': {
      ctx.drawImage(prev, 0, 0)
      // CSS circle(75%): the percentage is of sqrt(w² + h²) / sqrt(2)
      const r = p * 0.75 * (Math.hypot(w, h) / Math.SQRT2)
      ctx.beginPath()
      ctx.arc(w / 2, h / 2, Math.max(0, r), 0, Math.PI * 2)
      ctx.clip()
      ctx.drawImage(next, 0, 0)
      break
    }
    case 'dissolve': {
      ctx.drawImage(prev, 0, 0)
      const shown = Math.round(p * DISSOLVE_ORDER.length)
      const bw = w / DISSOLVE_COLS
      const bh = h / DISSOLVE_ROWS
      ctx.beginPath()
      for (let i = 0; i < shown; i++) {
        const k = DISSOLVE_ORDER[i]!
        const col = k % DISSOLVE_COLS
        const row = Math.floor(k / DISSOLVE_COLS)
        // Overlap a hair so block seams do not show as hairlines
        ctx.rect(col * bw, row * bh, bw + 0.5, bh + 0.5)
      }
      ctx.clip()
      ctx.drawImage(next, 0, 0)
      break
    }
    case 'zoom': {
      ctx.drawImage(prev, 0, 0)
      const s = 0.2 + 0.8 * p
      ctx.globalAlpha = p
      ctx.translate(w / 2, h / 2)
      ctx.scale(s, s)
      ctx.drawImage(next, -w / 2, -h / 2)
      break
    }
  }
  ctx.restore()
}

function clipRect(ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number) {
  ctx.beginPath()
  ctx.rect(x, y, Math.max(0, w), Math.max(0, h))
  ctx.clip()
}
