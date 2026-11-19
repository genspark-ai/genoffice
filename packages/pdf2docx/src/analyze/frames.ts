/**
 * Empty stroke frames (P16 K): a stroked rectangle with NOTHING inside — a
 * quote/answer box on a form. The lattice solver rightly rejects 1×1 grids,
 * so the four strokes used to degrade into two paragraph rules (the vertical
 * edges have no rendering path at all). Extracted here as a whole frame and
 * pinned behind the text as a transparent bordered bitmap.
 */
import type { Rect } from '../geometry'
import { rectUnionAll } from '../geometry'
import type { Stroke } from '../ir'
import { groupStrokes } from './table'

/** frames narrower than this are checkbox-scale (P6 territory), not boxes */
const FRAME_MIN_W_PT = 40
const FRAME_MIN_H_PT = 12
/** an edge stroke must reach this close to the frame's corners */
const FRAME_EDGE_TOL_PT = 2

export interface EmptyFrame {
  box: Rect
  color: string
  widthPt: number
}

const centerInside = (b: Rect, r: Rect): boolean => {
  const cx = (b.x0 + b.x1) / 2
  const cy = (b.y0 + b.y1) / 2
  return cx >= r.x0 && cx <= r.x1 && cy >= r.y0 && cy <= r.y1
}

/**
 * Pull empty 4-stroke rectangles out of the stroke pool (mutates `strokes`).
 * Groups already claimed by a solved lattice grid and rectangles holding any
 * text stay put.
 */
export function extractEmptyFrames(
  strokes: Stroke[],
  chars: ReadonlyArray<{ box: Rect; code: number }>,
  gridBoxes: readonly Rect[],
): EmptyFrame[] {
  const frames: EmptyFrame[] = []
  const consumed = new Set<Stroke>()
  for (const group of groupStrokes(strokes)) {
    if (group.length !== 4) continue
    const hs = group.filter((s) => s.orientation === 'h')
    const vs = group.filter((s) => s.orientation === 'v')
    if (hs.length !== 2 || vs.length !== 2) continue
    const box = rectUnionAll(group.map((s) => s.box))
    if (box.x1 - box.x0 < FRAME_MIN_W_PT || box.y1 - box.y0 < FRAME_MIN_H_PT) continue
    const spansSide =
      hs.every(
        (s) => s.box.x0 <= box.x0 + FRAME_EDGE_TOL_PT && s.box.x1 >= box.x1 - FRAME_EDGE_TOL_PT,
      ) &&
      vs.every(
        (s) => s.box.y0 <= box.y0 + FRAME_EDGE_TOL_PT && s.box.y1 >= box.y1 - FRAME_EDGE_TOL_PT,
      )
    if (!spansSide) continue
    if (gridBoxes.some((g) => centerInside(box, g))) continue
    if (chars.some((c) => c.code > 0x20 && centerInside(c.box, box))) continue
    frames.push({
      box,
      color: group[0]!.color,
      widthPt: Math.max(...group.map((s) => s.widthPt)),
    })
    for (const s of group) consumed.add(s)
  }
  if (consumed.size > 0) {
    for (let i = strokes.length - 1; i >= 0; i--) {
      if (consumed.has(strokes[i]!)) strokes.splice(i, 1)
    }
  }
  return frames
}
