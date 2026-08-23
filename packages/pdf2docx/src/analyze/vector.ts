/**
 * Vector-illustration region detection (P4): the paths shape normalization
 * ignores (curves, diagonals, non-rect fills) are exactly what charts, logos
 * and diagrams are made of. Dense clusters of such subpaths over text-sparse
 * areas become regions the EXTRACTION layer rasterizes via PDFium (replacing
 * the OpenCV step of the reference implementation). Pure geometry — the
 * caller does the rendering.
 */
import type { Rect } from '../geometry'
import { approxEq, median, rectArea, rectUnion } from '../geometry'
import type { PdfChar, RawPath, RawSubpath } from '../ir'
import { rectOfSubpath } from './shapes'

/** subpath bboxes within this distance (pt) cluster into one region */
const ART_CLUSTER_TOL_PT = 10
/** clusters with fewer art subpaths are stray decorations, not illustrations */
const ART_MIN_SUBPATHS = 6
/** minimum region size (pt) */
const ART_MIN_DIM_PT = 24
/** char-ink share of the region area above this = text region, leave it alone */
const ART_TEXT_DENSITY_MAX = 0.15
/**
 * regions holding at most this many chars skip the density gate (P10 C): a
 * lone display glyph on a stamp/seal fills the region's area without being
 * running text — the density gate exists to protect paragraphs, not monograms
 */
const ART_MONOGRAM_MAX_CHARS = 2
/** padding around the final region (pt) so hairlines at the edge survive */
const ART_PAD_PT = 2
/** same tolerance as shape normalization uses for axis-aligned checks */
const AXIS_TOL = 0.5

// ── running-text veto (P5): magazine-style pages decorate whole text areas
// with curves (card corners, blob backgrounds); the ink-density gate alone
// misses them because generous padding dilutes the density. A candidate
// containing paragraphs of body-size text must never be rasterized. ──

/** this many body-size text lines inside a candidate region veto it */
export const ART_VETO_MIN_LINES = 3
/** a veto line needs at least this many visible characters */
const VETO_LINE_MIN_CHARS = 6
/** ... spanning at least this width (pt) — rules out chart tick labels */
const VETO_LINE_MIN_WIDTH_PT = 60
/** chars at or above this share of the page's dominant font size are "body" */
const VETO_BODY_SIZE_RATIO = 0.75

/** count body-size text lines (grouped by baseline) inside a region */
function bodyTextLineCount(
  visibleChars: readonly PdfChar[],
  region: Rect,
  pageBodySize: number,
): number {
  const rows = new Map<number, { count: number; x0: number; x1: number }>()
  for (const c of visibleChars) {
    if (c.fontSize < VETO_BODY_SIZE_RATIO * pageBodySize) continue
    if (!centerInside(c.box, region)) continue
    // quantize the baseline to half the font size so one visual line = one bucket
    const key = Math.round(c.originY / Math.max(1, pageBodySize / 2))
    const row = rows.get(key)
    if (row) {
      row.count++
      row.x0 = Math.min(row.x0, c.box.x0)
      row.x1 = Math.max(row.x1, c.box.x1)
    } else {
      rows.set(key, { count: 1, x0: c.box.x0, x1: c.box.x1 })
    }
  }
  let lines = 0
  for (const row of rows.values()) {
    if (row.count >= VETO_LINE_MIN_CHARS && row.x1 - row.x0 >= VETO_LINE_MIN_WIDTH_PT) lines++
  }
  return lines
}

const bboxOfSubpath = (sub: RawSubpath): Rect | null => {
  if (sub.points.length === 0) return null
  const xs = sub.points.map((p) => p.x)
  const ys = sub.points.map((p) => p.y)
  return { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) }
}

/** a subpath the P2 normalization cannot represent (the raw material of vector art) */
function isArtSubpath(path: RawPath, sub: RawSubpath): boolean {
  if (sub.hasCurves) return true
  if (rectOfSubpath(sub)) return false
  if (path.filled) return true // non-rect filled polygon
  // stroked polyline: art when any segment is diagonal
  const pts = sub.closed && sub.points.length > 2 ? [...sub.points, sub.points[0]!] : sub.points
  for (let i = 1; i < pts.length; i++) {
    const a = pts[i - 1]!
    const b = pts[i]!
    if (!approxEq(a.x, b.x, AXIS_TOL) && !approxEq(a.y, b.y, AXIS_TOL)) return true
  }
  return false
}

const touches = (a: Rect, b: Rect, tol: number): boolean =>
  a.x0 <= b.x1 + tol && b.x0 <= a.x1 + tol && a.y0 <= b.y1 + tol && b.y0 <= a.y1 + tol

/** union-find clustering of rects by proximity */
function clusterRects(boxes: readonly Rect[], tol: number): Rect[][] {
  const parent = boxes.map((_, i) => i)
  const find = (i: number): number => {
    while (parent[i] !== i) {
      parent[i] = parent[parent[i]!]!
      i = parent[i]!
    }
    return i
  }
  for (let i = 0; i < boxes.length; i++) {
    for (let j = i + 1; j < boxes.length; j++) {
      if (touches(boxes[i]!, boxes[j]!, tol)) {
        const ri = find(i)
        const rj = find(j)
        if (ri !== rj) parent[rj] = ri
      }
    }
  }
  const byRoot = new Map<number, Rect[]>()
  for (let i = 0; i < boxes.length; i++) {
    const root = find(i)
    let list = byRoot.get(root)
    if (!list) byRoot.set(root, (list = []))
    list.push(boxes[i]!)
  }
  return [...byRoot.values()]
}

const centerInside = (r: Rect, region: Rect): boolean => {
  const cx = (r.x0 + r.x1) / 2
  const cy = (r.y0 + r.y1) / 2
  return cx >= region.x0 && cx <= region.x1 && cy >= region.y0 && cy <= region.y1
}

/**
 * Detect vector-illustration regions on a page. Returns page-space rects the
 * caller should rasterize; text and paths inside them are covered by the
 * bitmap and must be dropped from normal extraction.
 */
export function detectVectorRegions(
  paths: readonly RawPath[],
  chars: readonly PdfChar[],
  pageRect: Rect,
): Rect[] {
  const artBoxes: Rect[] = []
  const allBoxes: Rect[] = []
  for (const path of paths) {
    if (!path.filled && !path.stroked) continue
    for (const sub of path.subpaths) {
      const box = bboxOfSubpath(sub)
      if (!box) continue
      allBoxes.push(box)
      if (isArtSubpath(path, sub)) artBoxes.push(box)
    }
  }
  if (artBoxes.length < ART_MIN_SUBPATHS) return []

  const visibleChars = chars.filter((c) => !c.isGenerated && c.code > 0x20)
  const pageBodySize = median(visibleChars.map((c) => c.fontSize)) || 12

  /** density + running-text gates; null = the candidate must not be rasterized */
  const admit = (candidate: Rect): Rect | null => {
    if (
      candidate.x1 - candidate.x0 < ART_MIN_DIM_PT ||
      candidate.y1 - candidate.y0 < ART_MIN_DIM_PT
    )
      return null
    // text-density gate: paragraphs decorated by a curve are NOT illustrations
    const area = rectArea(candidate)
    if (area <= 0) return null
    let charInk = 0
    let charCount = 0
    for (const c of visibleChars) {
      if (centerInside(c.box, candidate)) {
        charInk += rectArea(c.box)
        charCount++
      }
    }
    if (charCount > ART_MONOGRAM_MAX_CHARS && charInk / area > ART_TEXT_DENSITY_MAX) return null
    // running-text veto: body paragraphs inside the box → not an illustration
    if (bodyTextLineCount(visibleChars, candidate, pageBodySize) >= ART_VETO_MIN_LINES) return null
    return candidate
  }

  const regions: Rect[] = []
  for (const cluster of clusterRects(artBoxes, ART_CLUSTER_TOL_PT)) {
    if (cluster.length < ART_MIN_SUBPATHS) continue
    const base = cluster.reduce(rectUnion)
    let region = base

    // pull in the illustration's axis-aligned furniture (chart axes, frames):
    // any subpath box centered inside the region extends it; run twice so the
    // frame the first pass absorbed can pull its own overhang in
    for (let pass = 0; pass < 2; pass++) {
      for (const box of allBoxes) {
        if (centerInside(box, region)) region = rectUnion(region, box)
      }
    }

    // when the grown region fails (usually the running-text veto after the
    // furniture absorption swallowed a text area), fall back to the bare art
    // cluster; if that still contains paragraphs, the text pipeline wins and
    // nothing is rasterized here
    const admitted = admit(region) ?? (region !== base ? admit(base) : null)
    if (!admitted) continue

    regions.push({
      x0: Math.max(pageRect.x0, admitted.x0 - ART_PAD_PT),
      y0: Math.max(pageRect.y0, admitted.y0 - ART_PAD_PT),
      x1: Math.min(pageRect.x1, admitted.x1 + ART_PAD_PT),
      y1: Math.min(pageRect.y1, admitted.y1 + ART_PAD_PT),
    })
  }

  // merge regions that grew into each other
  const merged: Rect[] = []
  for (const cluster of clusterRects(regions, 0)) {
    merged.push(cluster.reduce(rectUnion))
  }
  return merged
}
