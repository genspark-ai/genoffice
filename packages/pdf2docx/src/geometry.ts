/**
 * Geometry primitives shared by every layer. All comparisons go through the
 * tolerance helpers below — exact float equality on PDF coordinates is the
 * single biggest crash source in the reference implementation (pdf2docx).
 */

/** Axis-aligned rectangle in PDF page space: origin bottom-left, y grows up. */
export interface Rect {
  /** left */
  x0: number
  /** bottom */
  y0: number
  /** right */
  x1: number
  /** top */
  y1: number
}

/** |a - b| <= tol */
export const approxEq = (a: number, b: number, tol: number): boolean => Math.abs(a - b) <= tol

/** a >= b within tolerance */
export const gteTol = (a: number, b: number, tol: number): boolean => a >= b - tol

/** a <= b within tolerance */
export const lteTol = (a: number, b: number, tol: number): boolean => a <= b + tol

export const rectWidth = (r: Rect): number => r.x1 - r.x0
export const rectHeight = (r: Rect): number => r.y1 - r.y0
export const rectArea = (r: Rect): number => Math.max(0, rectWidth(r)) * Math.max(0, rectHeight(r))
export const rectCenterX = (r: Rect): number => (r.x0 + r.x1) / 2
export const rectCenterY = (r: Rect): number => (r.y0 + r.y1) / 2

export function rectUnion(a: Rect, b: Rect): Rect {
  return {
    x0: Math.min(a.x0, b.x0),
    y0: Math.min(a.y0, b.y0),
    x1: Math.max(a.x1, b.x1),
    y1: Math.max(a.y1, b.y1),
  }
}

export function rectUnionAll(rects: readonly Rect[]): Rect {
  if (rects.length === 0) return { x0: 0, y0: 0, x1: 0, y1: 0 }
  return rects.reduce(rectUnion)
}

export function intersectArea(a: Rect, b: Rect): number {
  const w = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0)
  const h = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0)
  return w > 0 && h > 0 ? w * h : 0
}

/** intersection ÷ area of `a` (how much of `a` sits inside `b`); 0 for empty `a` */
export function overlapRatio(a: Rect, b: Rect): number {
  const area = rectArea(a)
  return area > 0 ? intersectArea(a, b) / area : 0
}

/** vertical overlap length ÷ the shorter rect's height (0 when either is flat) */
export function verticalOverlapRatio(a: Rect, b: Rect): number {
  const overlap = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0)
  const minH = Math.min(rectHeight(a), rectHeight(b))
  return minH > 0 ? Math.max(0, overlap) / minH : 0
}

// ── 1-D intervals (x/y projections for tables & column detection, P3) ──

export interface Interval {
  lo: number
  hi: number
}

/** merge overlapping intervals, also closing gaps narrower than `minGap` */
export function mergeIntervals(intervals: readonly Interval[], minGap = 0): Interval[] {
  const sorted = [...intervals].sort((a, b) => a.lo - b.lo)
  const out: Interval[] = []
  for (const iv of sorted) {
    const last = out[out.length - 1]
    if (last && iv.lo - last.hi < minGap) last.hi = Math.max(last.hi, iv.hi)
    else out.push({ ...iv })
  }
  return out
}

/** pairwise intersection of two merged interval lists */
export function intersectIntervals(a: readonly Interval[], b: readonly Interval[]): Interval[] {
  const out: Interval[] = []
  let i = 0
  let j = 0
  while (i < a.length && j < b.length) {
    const lo = Math.max(a[i]!.lo, b[j]!.lo)
    const hi = Math.min(a[i]!.hi, b[j]!.hi)
    if (hi > lo) out.push({ lo, hi })
    if (a[i]!.hi < b[j]!.hi) i++
    else j++
  }
  return out
}

/** uncovered sub-intervals of [lo, hi] given merged `covered` */
export function complementIntervals(
  covered: readonly Interval[],
  lo: number,
  hi: number,
): Interval[] {
  const out: Interval[] = []
  let cursor = lo
  for (const iv of covered) {
    if (iv.lo > cursor) out.push({ lo: cursor, hi: Math.min(iv.lo, hi) })
    cursor = Math.max(cursor, iv.hi)
    if (cursor >= hi) break
  }
  if (cursor < hi) out.push({ lo: cursor, hi })
  return out.filter((iv) => iv.hi > iv.lo)
}

/** median of a non-empty list; 0 for an empty one */
export function median(values: readonly number[]): number {
  if (values.length === 0) return 0
  const sorted = [...values].sort((a, b) => a - b)
  const mid = Math.floor(sorted.length / 2)
  return sorted.length % 2 === 1 ? sorted[mid]! : (sorted[mid - 1]! + sorted[mid]!) / 2
}

/** grid resolution for coverageRatio (cells per axis) */
const COVERAGE_GRID = 48

/**
 * Share of the page area covered by the union of `boxes`, approximated on a
 * coarse occupancy grid (a cell counts once it overlaps any box's interior).
 */
export function coverageRatio(boxes: readonly Rect[], widthPt: number, heightPt: number): number {
  if (widthPt <= 0 || heightPt <= 0) return 0
  const n = COVERAGE_GRID
  const grid = new Uint8Array(n * n)
  const cw = widthPt / n
  const ch = heightPt / n
  for (const b of boxes) {
    const cx0 = Math.max(0, Math.floor(b.x0 / cw))
    const cx1 = Math.min(n - 1, Math.ceil(b.x1 / cw) - 1)
    const cy0 = Math.max(0, Math.floor(b.y0 / ch))
    const cy1 = Math.min(n - 1, Math.ceil(b.y1 / ch) - 1)
    for (let y = cy0; y <= cy1; y++) {
      for (let x = cx0; x <= cx1; x++) grid[y * n + x] = 1
    }
  }
  let covered = 0
  for (const cell of grid) covered += cell
  return covered / (n * n)
}
