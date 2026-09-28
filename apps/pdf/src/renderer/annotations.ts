import type { MarkupInput, MarkupType } from '../shared/ipc'

export interface LocalMarkup extends MarkupInput {
  id: string
}

export const MARKUP_COLORS: Record<MarkupType, [number, number, number]> = {
  highlight: [1, 0.87, 0.35],
  underline: [0.17, 0.4, 1],
  strikeout: [0.86, 0.22, 0.18],
}

/** Page geometry: unrotated PDF page size + total display rotation (base /Rotate + unsaved delta) */
export interface PageGeom {
  pw: number
  ph: number
  rot: number
  /** CropBox lower-left in PDF user space. */
  x0?: number
  y0?: number
  /** PDF /UserUnit; layout coordinates use this multiplier while PDF APIs use raw units. */
  userUnit?: number
}

const normRot = (r: number) => ((r % 360) + 360) % 360

/** Display size (scale=1; width/height swap with rotation) */
export function geomDispSize(g: PageGeom): { width: number; height: number } {
  return normRot(g.rot) % 180 === 0 ? { width: g.pw, height: g.ph } : { width: g.ph, height: g.pw }
}

/** Display coords (origin at page top-left, y down, scale=1) → PDF user space (y up) */
export function viewToPdf(g: PageGeom, vx: number, vy: number): [number, number] {
  const x0 = g.x0 ?? 0
  const y0 = g.y0 ?? 0
  const unit = g.userUnit ?? 1
  const pw = g.pw / unit
  const ph = g.ph / unit
  vx /= unit
  vy /= unit
  switch (normRot(g.rot)) {
    case 90:
      return [x0 + vy, y0 + vx]
    case 180:
      return [x0 + pw - vx, y0 + vy]
    case 270:
      return [x0 + pw - vy, y0 + ph - vx]
    default:
      return [x0 + vx, y0 + ph - vy]
  }
}

/** PDF user space → display coords (scale=1) */
export function pdfToView(g: PageGeom, x: number, y: number): [number, number] {
  x -= g.x0 ?? 0
  y -= g.y0 ?? 0
  const unit = g.userUnit ?? 1
  const pw = g.pw / unit
  const ph = g.ph / unit
  let result: [number, number]
  switch (normRot(g.rot)) {
    case 90:
      result = [y, x]
      break
    case 180:
      result = [pw - x, y]
      break
    case 270:
      result = [ph - y, pw - x]
      break
    default:
      result = [x, ph - y]
  }
  return [result[0] * unit, result[1] * unit]
}

/** PDF-space rect [x1,y1,x2,y2] → displayed pixel box (scaled) */
export function pdfRectToCss(
  g: PageGeom,
  rect: readonly [number, number, number, number],
  scale: number,
): { left: number; top: number; width: number; height: number } {
  const [ax, ay] = pdfToView(g, rect[0], rect[1])
  const [bx, by] = pdfToView(g, rect[2], rect[3])
  return {
    left: Math.min(ax, bx) * scale,
    top: Math.min(ay, by) * scale,
    width: Math.abs(bx - ax) * scale,
    height: Math.abs(by - ay) * scale,
  }
}

/** quad (corners in any order) → PDF-space bounding rect */
export function quadToRect(q: number[]): [number, number, number, number] {
  const xs = [q[0]!, q[2]!, q[4]!, q[6]!]
  const ys = [q[1]!, q[3]!, q[5]!, q[7]!]
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]
}

/**
 * Quad-set equality within a tolerance, order-insensitive. Used for the Word-style
 * markup toggle: the same text selected at a different zoom level produces slightly
 * different coordinates, and QuadPoints saved to the file round-trip through float32.
 */
export function quadSetsMatch(a: number[][], b: number[][], tol = 2): boolean {
  if (a.length !== b.length) return false
  const used = new Array<boolean>(b.length).fill(false)
  for (const qa of a) {
    const i = b.findIndex(
      (qb, j) =>
        !used[j] && qb.length === qa.length && qa.every((v, k) => Math.abs(v - qb[k]!) <= tol),
    )
    if (i < 0) return false
    used[i] = true
  }
  return true
}

interface ViewRect {
  left: number
  right: number
  top: number
  bottom: number
}

interface SelectionLine {
  crossStart: number
  crossEnd: number
  minCrossSize: number
  rects: ViewRect[]
}

const crossBounds = (r: ViewRect): [number, number] => [r.top, r.bottom]

const mainBounds = (r: ViewRect): [number, number] => [r.left, r.right]

/** Pick the one common line band that overlaps most; ties stay independent. */
function matchingLine(lines: readonly SelectionLine[], rect: ViewRect): SelectionLine | null {
  const [start, end] = crossBounds(rect)
  const crossSize = end - start
  let match: SelectionLine | null = null
  let bestOverlap = -1
  let tied = false
  for (const line of lines) {
    const overlap = Math.min(end, line.crossEnd) - Math.max(start, line.crossStart)
    if (overlap * 2 < Math.min(crossSize, line.minCrossSize)) continue
    if (overlap > bestOverlap) {
      match = line
      bestOverlap = overlap
      tied = false
    } else if (overlap === bestOverlap) tied = true
  }
  return tied ? null : match
}

/** Unify each nearby fragment cluster on a visual line, without crossing wide gaps. */
function normalizeSelectionRects(rects: readonly ViewRect[]): ViewRect[] {
  const lineBands: SelectionLine[] = []
  for (const rect of rects) {
    const [crossStart, crossEnd] = crossBounds(rect)
    const crossSize = crossEnd - crossStart
    const line = matchingLine(lineBands, rect)
    if (line) {
      line.crossStart = Math.max(line.crossStart, crossStart)
      line.crossEnd = Math.min(line.crossEnd, crossEnd)
      line.minCrossSize = Math.min(line.minCrossSize, crossSize)
      line.rects.push(rect)
    } else lineBands.push({ crossStart, crossEnd, minCrossSize: crossSize, rects: [rect] })
  }

  return lineBands.flatMap((line) => {
    // the running union of each cluster, widened in place: rebuilding it from the
    // cluster's rects on every fragment made the sweep quadratic in line length
    const clusters: ViewRect[] = []
    // Text-layer spans arrive in content-stream order, not visual order
    for (const rect of [...line.rects].sort((a, b) => a.left - b.left)) {
      const previous = clusters.at(-1)
      if (!previous) {
        clusters.push(rect)
        continue
      }
      const [previousStart, previousEnd] = mainBounds(previous)
      const [start, end] = mainBounds(rect)
      const gap = Math.max(start - previousEnd, previousStart - end, 0)
      if (gap <= line.minCrossSize / 2) {
        clusters[clusters.length - 1] = {
          left: Math.min(previous.left, rect.left),
          right: Math.max(previous.right, rect.right),
          top: Math.min(previous.top, rect.top),
          bottom: Math.max(previous.bottom, rect.bottom),
        }
      } else clusters.push(rect)
    }
    return clusters
  })
}

/**
 * Drop the element-level boxes that merely enclose span-level ones, so the union a
 * line cluster collapses to stays tight. getClientRects hands back both, and a big
 * box left in place would widen every quad on its line to the whole paragraph.
 *
 * getClientRects is in tree order, so a box precedes what it contains: one sweep
 * keeping the containers still open to the right of and above the cursor replaces
 * the all-pairs containment test, which was quadratic over the thousands of boxes
 * a selection across a large text page returns.
 */
function dropEnclosingRects(rects: readonly DOMRect[]): DOMRect[] {
  const enclosing = new Set<DOMRect>()
  const open: DOMRect[] = []
  for (const r of rects) {
    // a container can hold nothing more once the sweep has moved past its box:
    // tree order only ever moves right along a line and down to the next line.
    // A box starting at another's right edge is a sibling, not a child, so
    // abutting spans on one line drop off the stack instead of piling up
    while (open.length > 0) {
      const top = open[open.length - 1]!
      if (r.left < top.right - 1 && r.top < top.bottom - 1) break
      open.pop()
    }
    for (const c of open) {
      if (
        r.left < c.left - 1 ||
        r.right > c.right + 1 ||
        r.top < c.top - 1 ||
        r.bottom > c.bottom + 1
      )
        continue
      if (r.width < c.width - 1 || r.height < c.height - 1) enclosing.add(c)
    }
    open.push(r)
  }
  return rects.filter((r) => !enclosing.has(r))
}

/**
 * Current selection → PDF-coordinate quads grouped by visible page (y up; each quad
 * [x1,yMax,x2,yMax,x1,yMin,x2,yMin]). Returns null when the selection is empty.
 * geoms map 1:1 to .pdf-page elements and include rotation conversion.
 */
export function selectionQuadsByPage(
  scrollEl: HTMLElement,
  geoms: PageGeom[],
  scale: number,
): Map<number, number[][]> | null {
  const sel = window.getSelection()
  if (!sel || sel.isCollapsed || sel.rangeCount === 0) return null
  const range = sel.getRangeAt(0)
  if (!scrollEl.contains(range.commonAncestorContainer)) return null

  const pageEls = [...scrollEl.querySelectorAll<HTMLElement>('.pdf-page')]
  const pageRects = pageEls.map((el) => el.getBoundingClientRect())

  // Selection rects mix element-level big boxes with overlapping span-level ones: drop boxes containing others, then dedupe by pixel
  const raw = [...range.getClientRects()].filter((r) => r.width > 1 && r.height > 1)
  const rects = dropEnclosingRects(raw)

  const rectsByPage = new Map<number, ViewRect[]>()
  const seen = new Set<string>()
  for (const r of rects) {
    const cx = (r.left + r.right) / 2
    const cy = (r.top + r.bottom) / 2
    const idx = pageRects.findIndex(
      (p) => cx >= p.left && cx <= p.right && cy >= p.top && cy <= p.bottom,
    )
    if (idx < 0 || !geoms[idx]) continue
    const key = `${idx}:${Math.round(r.left)}:${Math.round(r.right)}:${Math.round(r.top)}:${Math.round(r.bottom)}`
    if (seen.has(key)) continue
    seen.add(key)
    const list = rectsByPage.get(idx)
    if (list) list.push(r)
    else rectsByPage.set(idx, [r])
  }

  const byPage = new Map<number, number[][]>()
  for (const [idx, pageSelectionRects] of rectsByPage) {
    const p = pageRects[idx]!
    const g = geoms[idx]!
    const rects =
      normRot(g.rot) % 180 === 0 ? normalizeSelectionRects(pageSelectionRects) : pageSelectionRects
    const quads = rects.map((r) => {
      const [ax, ay] = viewToPdf(g, (r.left - p.left) / scale, (r.top - p.top) / scale)
      const [bx, by] = viewToPdf(g, (r.right - p.left) / scale, (r.bottom - p.top) / scale)
      const [x1, x2] = [Math.min(ax, bx), Math.max(ax, bx)]
      const [yMin, yMax] = [Math.min(ay, by), Math.max(ay, by)]
      return [x1, yMax, x2, yMax, x1, yMin, x2, yMin]
    })
    if (quads.length > 0) byPage.set(idx, quads)
  }
  return byPage.size > 0 ? byPage : null
}
