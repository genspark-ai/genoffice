/**
 * Alignment snapping — while dragging, snaps an element's left/center/right and top/middle/bottom
 * edges to target edges (page center lines/page edges/other elements' edges), returning the
 * snapped x/y and the guides to draw. Pure functions, unit-testable.
 */

export interface SnapTarget {
  /** Set of x coordinates of vertical guides (an element's left/center/right can snap to these) */
  v: number[]
  /** Set of y coordinates of horizontal guides */
  h: number[]
}

export interface Guide {
  axis: 'v' | 'h'
  pos: number
}

export interface SnapResult {
  x: number | null
  y: number | null
  guides: Guide[]
}

interface DragBox {
  x: number
  y: number
  w: number
  h: number
}

/**
 * Compute snapping. Compare the box's three x anchors (left/centerX/right) against all target
 * v lines and snap to the nearest one within threshold; same for y.
 */
export function computeSnap(box: DragBox, targets: SnapTarget[], threshold = 6): SnapResult {
  const allV = targets.flatMap((t) => t.v)
  const allH = targets.flatMap((t) => t.h)

  const xAnchors = [
    { edge: 'left' as const, val: box.x, offset: 0 },
    { edge: 'center' as const, val: box.x + box.w / 2, offset: box.w / 2 },
    { edge: 'right' as const, val: box.x + box.w, offset: box.w },
  ]
  const yAnchors = [
    { edge: 'top' as const, val: box.y, offset: 0 },
    { edge: 'middle' as const, val: box.y + box.h / 2, offset: box.h / 2 },
    { edge: 'bottom' as const, val: box.y + box.h, offset: box.h },
  ]

  const bestX = pickSnap(xAnchors, allV, threshold)
  const bestY = pickSnap(yAnchors, allH, threshold)

  const guides: Guide[] = []
  if (bestX) guides.push({ axis: 'v', pos: bestX.line })
  if (bestY) guides.push({ axis: 'h', pos: bestY.line })

  return {
    x: bestX ? bestX.line - bestX.offset : null,
    y: bestY ? bestY.line - bestY.offset : null,
    guides,
  }
}

function pickSnap(
  anchors: Array<{ val: number; offset: number }>,
  lines: number[],
  threshold: number,
): { line: number; offset: number } | null {
  let best: { line: number; offset: number; dist: number } | null = null
  for (const a of anchors) {
    for (const line of lines) {
      const dist = Math.abs(a.val - line)
      if (dist <= threshold && (!best || dist < best.dist)) {
        best = { line, offset: a.offset, dist }
      }
    }
  }
  return best ? { line: best.line, offset: best.offset } : null
}

// ── Equal-spacing snapping (smart guides) ───────────────────────────────────────────

/** Equal-spacing indicator: double-headed arrow at `at` from `from` to `to` (axis=x means horizontal spacing). */
export interface SpacingIndicator {
  axis: 'x' | 'y'
  from: number
  to: number
  at: number
}

export interface SpacingSnapResult {
  x: number | null
  y: number | null
  indicators: SpacingIndicator[]
}

/** Interval overlap length (used to decide same row/column band). */
function overlap(a0: number, a1: number, b0: number, b1: number): number {
  return Math.min(a1, b1) - Math.max(a0, b0)
}

interface SpacingCandidate {
  pos: number
  dist: number
  indicators: Array<Omit<SpacingIndicator, 'axis'>>
}

/**
 * Equal-spacing snapping (smart guides):
 * - Centering: the dragged element sits between two neighbors (left/right or top/bottom); snap when both gaps are equal.
 * - Continuation: one side already has two elements with gap g; snap when the dragged element's gap to the nearest neighbor equals g.
 * Neighbor = element in the same band as the dragged box (projection overlap on the other axis).
 * Returns the snapped x/y and a double-headed arrow indicator for each gap.
 */
export function computeSpacingSnap(
  box: DragBox,
  others: DragBox[],
  threshold = 6,
): SpacingSnapResult {
  const x = spacingAxis(
    { lo: box.x, len: box.w, band0: box.y, band1: box.y + box.h },
    others.map((o) => ({ lo: o.x, len: o.w, band0: o.y, band1: o.y + o.h })),
    threshold,
  )
  const y = spacingAxis(
    { lo: box.y, len: box.h, band0: box.x, band1: box.x + box.w },
    others.map((o) => ({ lo: o.y, len: o.h, band0: o.x, band1: o.x + o.w })),
    threshold,
  )
  const indicators = [
    ...(x?.indicators.map((i) => ({ ...i, axis: 'x' as const })) ?? []),
    ...(y?.indicators.map((i) => ({ ...i, axis: 'y' as const })) ?? []),
  ]
  return { x: x ? x.pos : null, y: y ? y.pos : null, indicators }
}

interface AxisBox {
  lo: number
  len: number
  band0: number
  band1: number
}

/** Single-axis equal spacing: returns the snapped lo on that axis (same coordinate axis as input). */
function spacingAxis(
  box: AxisBox,
  others: AxisBox[],
  threshold: number,
): { pos: number; indicators: Array<Omit<SpacingIndicator, 'axis'>> } | null {
  const hi = box.lo + box.len
  // Same-band neighbors (projection overlap on the other axis)
  const band = others.filter((o) => overlap(box.band0, box.band1, o.band0, o.band1) > 0)
  // Nearest neighbor on the left (top) / right (bottom)
  const lefts = band.filter((o) => o.lo + o.len <= box.lo + threshold)
  const rights = band.filter((o) => o.lo >= hi - threshold)
  const ln = lefts.length ? lefts.reduce((a, b) => (a.lo + a.len > b.lo + b.len ? a : b)) : null
  const rn = rights.length ? rights.reduce((a, b) => (a.lo < b.lo ? a : b)) : null
  // Draw arrows on the midline of the two bands' overlap region
  const midOf = (a: AxisBox, b: AxisBox) =>
    (Math.max(a.band0, b.band0) + Math.min(a.band1, b.band1)) / 2

  const cands: SpacingCandidate[] = []

  // Centering: equal gaps between the two neighbors
  if (ln && rn) {
    const lEdge = ln.lo + ln.len
    const target = (lEdge + rn.lo - box.len) / 2
    const gap = (rn.lo - lEdge - box.len) / 2
    if (gap >= 0 && Math.abs(target - box.lo) <= threshold) {
      const at = midOf(box, ln)
      cands.push({
        pos: target,
        dist: Math.abs(target - box.lo),
        indicators: [
          { from: lEdge, to: target, at },
          { from: target + box.len, to: rn.lo, at: midOf(box, rn) },
        ],
      })
    }
  }
  // Continuation (left/top side): gap g between B and ln -> snap the dragged box at distance g to the right of ln
  if (ln) {
    const beyond = band.filter((o) => o !== ln && o.lo + o.len <= ln.lo)
    const b = beyond.length ? beyond.reduce((a, c) => (a.lo + a.len > c.lo + c.len ? a : c)) : null
    if (b) {
      const g = ln.lo - (b.lo + b.len)
      const target = ln.lo + ln.len + g
      if (g >= 0 && Math.abs(target - box.lo) <= threshold) {
        cands.push({
          pos: target,
          dist: Math.abs(target - box.lo),
          indicators: [
            { from: b.lo + b.len, to: ln.lo, at: midOf(ln, b) },
            { from: ln.lo + ln.len, to: target, at: midOf(box, ln) },
          ],
        })
      }
    }
  }
  // Continuation (right/bottom side)
  if (rn) {
    const beyond = band.filter((o) => o !== rn && o.lo >= rn.lo + rn.len)
    const b = beyond.length ? beyond.reduce((a, c) => (a.lo < c.lo ? a : c)) : null
    if (b) {
      const g = b.lo - (rn.lo + rn.len)
      const target = rn.lo - g - box.len
      if (g >= 0 && Math.abs(target - box.lo) <= threshold) {
        cands.push({
          pos: target,
          dist: Math.abs(target - box.lo),
          indicators: [
            { from: target + box.len, to: rn.lo, at: midOf(box, rn) },
            { from: rn.lo + rn.len, to: b.lo, at: midOf(rn, b) },
          ],
        })
      }
    }
  }

  if (!cands.length) return null
  const best = cands.reduce((a, b) => (a.dist <= b.dist ? a : b))
  return { pos: best.pos, indicators: best.indicators }
}
