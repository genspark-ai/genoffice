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
}

const normRot = (r: number) => ((r % 360) + 360) % 360

/** Display size (scale=1; width/height swap with rotation) */
export function geomDispSize(g: PageGeom): { width: number; height: number } {
  return normRot(g.rot) % 180 === 0 ? { width: g.pw, height: g.ph } : { width: g.ph, height: g.pw }
}

/** Display coords (origin at page top-left, y down, scale=1) → PDF user space (y up) */
export function viewToPdf(g: PageGeom, vx: number, vy: number): [number, number] {
  switch (normRot(g.rot)) {
    case 90:
      return [vy, vx]
    case 180:
      return [g.pw - vx, vy]
    case 270:
      return [g.pw - vy, g.ph - vx]
    default:
      return [vx, g.ph - vy]
  }
}

/** PDF user space → display coords (scale=1) */
export function pdfToView(g: PageGeom, x: number, y: number): [number, number] {
  switch (normRot(g.rot)) {
    case 90:
      return [y, x]
    case 180:
      return [g.pw - x, y]
    case 270:
      return [g.ph - y, g.pw - x]
    default:
      return [x, g.ph - y]
  }
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
  const rects = raw.filter(
    (r, i) =>
      !raw.some(
        (o, j) =>
          j !== i &&
          r.left <= o.left + 1 &&
          r.right >= o.right - 1 &&
          r.top <= o.top + 1 &&
          r.bottom >= o.bottom - 1 &&
          (o.width < r.width - 1 || o.height < r.height - 1),
      ),
  )

  const byPage = new Map<number, number[][]>()
  const seen = new Set<string>()
  for (const r of rects) {
    const cx = (r.left + r.right) / 2
    const cy = (r.top + r.bottom) / 2
    const idx = pageRects.findIndex(
      (p) => cx >= p.left && cx <= p.right && cy >= p.top && cy <= p.bottom,
    )
    if (idx < 0 || !geoms[idx]) continue
    const p = pageRects[idx]!
    const g = geoms[idx]!
    const [ax, ay] = viewToPdf(g, (r.left - p.left) / scale, (r.top - p.top) / scale)
    const [bx, by] = viewToPdf(g, (r.right - p.left) / scale, (r.bottom - p.top) / scale)
    const [x1, x2] = [Math.min(ax, bx), Math.max(ax, bx)]
    const [yMin, yMax] = [Math.min(ay, by), Math.max(ay, by)]
    const key = `${idx}:${Math.round(x1)}:${Math.round(x2)}:${Math.round(yMin)}:${Math.round(yMax)}`
    if (seen.has(key)) continue
    seen.add(key)
    const quad = [x1, yMax, x2, yMax, x1, yMin, x2, yMin]
    const list = byPage.get(idx)
    if (list) list.push(quad)
    else byPage.set(idx, [quad])
  }
  return byPage.size > 0 ? byPage : null
}
