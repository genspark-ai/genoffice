/**
 * Edit Points: pure path helpers behind the vertex/control handles. Paths are
 * absolute M/L/C/Q/Z command lists in box-local px (the render node's space);
 * the commit converts them to a CustGeomPath in EMU.
 */
import type { ShapeRenderNode } from '@genoffice/pptx-render'

export type PathOp = 'M' | 'L' | 'C' | 'Q' | 'Z'
export interface PathCmd {
  op: PathOp
  pts: number[]
}

export interface Vertex {
  x: number
  y: number
  /** Command whose end point this vertex is */
  cmd: number
  /** Closing command whose end point coincides with this subpath start (moves with it) */
  dup?: number
  cIn?: { x: number; y: number; cmd: number; i: number }
  cOut?: { x: number; y: number; cmd: number; i: number }
}

const PT_COUNT: Record<PathOp, number> = { M: 2, L: 2, C: 6, Q: 4, Z: 0 }
const BEZIER_K = 0.5522847498
const EPS = 0.01

const endOf = (c: PathCmd) => ({ x: c.pts[c.pts.length - 2]!, y: c.pts[c.pts.length - 1]! })
const inIndex = (c: PathCmd) => (c.op === 'C' ? 2 : c.op === 'Q' ? 0 : -1)
const near = (a: { x: number; y: number }, b: { x: number; y: number }) =>
  Math.abs(a.x - b.x) < EPS && Math.abs(a.y - b.y) < EPS

export function parseSvgPath(d: string): PathCmd[] | null {
  const toks = d
    .trim()
    .split(/[\s,]+/)
    .filter(Boolean)
  const out: PathCmd[] = []
  let i = 0
  while (i < toks.length) {
    const op = toks[i++]!
    if (!(op in PT_COUNT)) return null
    const n = PT_COUNT[op as PathOp]
    const pts: number[] = []
    for (let k = 0; k < n; k++) {
      const v = Number(toks[i++])
      if (!Number.isFinite(v)) return null
      pts.push(v)
    }
    out.push({ op: op as PathOp, pts })
  }
  return out.length && out[0]!.op === 'M' ? out : null
}

export function toSvgPath(cmds: PathCmd[]): string {
  return cmds
    .map((c) =>
      c.op === 'Z' ? 'Z' : `${c.op} ${c.pts.map((v) => Math.round(v * 100) / 100).join(' ')}`,
    )
    .join(' ')
}

function rectPath(w: number, h: number): PathCmd[] {
  return [
    { op: 'M', pts: [0, 0] },
    { op: 'L', pts: [w, 0] },
    { op: 'L', pts: [w, h] },
    { op: 'L', pts: [0, h] },
    { op: 'Z', pts: [] },
  ]
}

/** Sides shrink to nothing on pills / fully rounded squares; a zero-length L would stack two vertices. */
function roundRectPath(w: number, h: number, r: number): PathCmd[] {
  const k = r * (1 - BEZIER_K)
  const cmds: PathCmd[] = [
    { op: 'M', pts: [r, 0] },
    { op: 'L', pts: [w - r, 0] },
    { op: 'C', pts: [w - k, 0, w, k, w, r] },
    { op: 'L', pts: [w, h - r] },
    { op: 'C', pts: [w, h - k, w - k, h, w - r, h] },
    { op: 'L', pts: [r, h] },
    { op: 'C', pts: [k, h, 0, h - k, 0, h - r] },
    { op: 'L', pts: [0, r] },
    { op: 'C', pts: [0, k, k, 0, r, 0] },
    { op: 'Z', pts: [] },
  ]
  return cmds.filter((c, i) => c.op !== 'L' || !near(endOf(c), endOf(cmds[i - 1]!)))
}

/**
 * The editable outline of a shape node, or null when Edit Points is not offered:
 * connectors, multi-layer geometry (separate fill/stroke paths) and presets the
 * renderer only approximates as a rectangle.
 */
export function shapeToEditablePath(node: ShapeRenderNode): PathCmd[] | null {
  if (node.line || node.fillPathData || node.strokePathData || !node.box) return null
  const { w, h } = node.box
  if (!(w > 0 && h > 0)) return null
  if (node.pathData) return parseSvgPath(node.pathData)
  const poly = node.polygonPoints
  if (poly && poly.length >= 6) {
    const cmds: PathCmd[] = [{ op: 'M', pts: [poly[0]!, poly[1]!] }]
    for (let i = 2; i < poly.length; i += 2) cmds.push({ op: 'L', pts: [poly[i]!, poly[i + 1]!] })
    cmds.push({ op: 'Z', pts: [] })
    return cmds
  }
  if (node.cornerRadiusPx && node.cornerRadiusPx > 0)
    return roundRectPath(w, h, Math.min(node.cornerRadiusPx, w / 2, h / 2))
  if (!node.presetGeometry || node.presetGeometry === 'rect') return rectPath(w, h)
  return null
}

export function canEditPoints(node: { type: string } | undefined): boolean {
  if (!node || (node.type !== 'shape' && node.type !== 'text')) return false
  return shapeToEditablePath(node as ShapeRenderNode) != null
}

/** Vertices in path order; a closing point that lands on its subpath start is folded into it. */
export function vertices(cmds: PathCmd[]): Vertex[] {
  const out: Vertex[] = []
  let start: Vertex | null = null
  for (let i = 0; i < cmds.length; i++) {
    const c = cmds[i]!
    if (c.op === 'Z') {
      start = null
      continue
    }
    const p = endOf(c)
    const v: Vertex = { x: p.x, y: p.y, cmd: i }
    const ii = inIndex(c)
    if (ii >= 0) v.cIn = { x: c.pts[ii]!, y: c.pts[ii + 1]!, cmd: i, i: ii }
    const next = cmds[i + 1]
    if (next && (next.op === 'C' || next.op === 'Q'))
      v.cOut = { x: next.pts[0]!, y: next.pts[1]!, cmd: i + 1, i: 0 }
    if (c.op === 'M') {
      start = v
    } else if (start && next?.op === 'Z' && near(p, start)) {
      start.dup = i
      if (v.cIn) start.cIn = v.cIn
      continue
    }
    out.push(v)
  }
  return out
}

const shift = (cmds: PathCmd[], cmd: number, i: number, dx: number, dy: number) => {
  const c = cmds[cmd]!
  c.pts[i] = c.pts[i]! + dx
  c.pts[i + 1] = c.pts[i + 1]! + dy
}

const clone = (cmds: PathCmd[]): PathCmd[] => cmds.map((c) => ({ op: c.op, pts: [...c.pts] }))

/** Move a vertex; its incoming and outgoing controls travel with it (corner-point semantics). */
export function moveVertex(cmds: PathCmd[], v: Vertex, dx: number, dy: number): PathCmd[] {
  const out = clone(cmds)
  const moveEnd = (cmd: number) => {
    const c = out[cmd]!
    shift(out, cmd, c.pts.length - 2, dx, dy)
    const ii = inIndex(c)
    if (ii >= 0) shift(out, cmd, ii, dx, dy)
  }
  moveEnd(v.cmd)
  if (v.dup != null) moveEnd(v.dup)
  if (v.cOut) shift(out, v.cOut.cmd, v.cOut.i, dx, dy)
  return out
}

export function moveControl(
  cmds: PathCmd[],
  v: Vertex,
  which: 'in' | 'out',
  dx: number,
  dy: number,
): PathCmd[] {
  const h = which === 'in' ? v.cIn : v.cOut
  if (!h) return cmds
  const out = clone(cmds)
  shift(out, h.cmd, h.i, dx, dy)
  return out
}

/** Remove a vertex; null when fewer than three would remain or the subpath would vanish. */
export function deleteVertex(cmds: PathCmd[], v: Vertex): PathCmd[] | null {
  if (vertices(cmds).length <= 3) return null
  const out = clone(cmds)
  const c = out[v.cmd]!
  if (c.op === 'M') {
    const next = out[v.cmd + 1]
    if (!next || next.op === 'Z') return null
    const p = endOf(next)
    out.splice(v.cmd, 2, { op: 'M', pts: [p.x, p.y] })
    if (v.dup != null) out.splice(v.dup - 1, 1)
  } else {
    out.splice(v.cmd, 1)
  }
  return vertices(out).length >= 3 ? out : null
}

/** Pen position before command i (a close returns the pen to the subpath start). */
function startOf(cmds: PathCmd[], i: number): { x: number; y: number } {
  const prev = cmds[i - 1]!
  return prev.op === 'Z' ? subpathStart(cmds, i - 1) : endOf(prev)
}

function subpathStart(cmds: PathCmd[], i: number): { x: number; y: number } {
  for (let k = i; k >= 0; k--) if (cmds[k]!.op === 'M') return endOf(cmds[k]!)
  return endOf(cmds[0]!)
}

function pointAt(cmds: PathCmd[], i: number, t: number): { x: number; y: number } {
  const c = cmds[i]!
  const p0 = startOf(cmds, i)
  const u = 1 - t
  if (c.op === 'L' || c.op === 'M')
    return { x: p0.x + (c.pts[0]! - p0.x) * t, y: p0.y + (c.pts[1]! - p0.y) * t }
  if (c.op === 'Z') {
    const s = subpathStart(cmds, i)
    return { x: p0.x + (s.x - p0.x) * t, y: p0.y + (s.y - p0.y) * t }
  }
  if (c.op === 'Q') {
    const [x1, y1, x, y] = c.pts as [number, number, number, number]
    return {
      x: u * u * p0.x + 2 * u * t * x1 + t * t * x,
      y: u * u * p0.y + 2 * u * t * y1 + t * t * y,
    }
  }
  const [x1, y1, x2, y2, x, y] = c.pts as [number, number, number, number, number, number]
  return {
    x: u * u * u * p0.x + 3 * u * u * t * x1 + 3 * u * t * t * x2 + t * t * t * x,
    y: u * u * u * p0.y + 3 * u * u * t * y1 + 3 * u * t * t * y2 + t * t * t * y,
  }
}

/** Closest point on any segment (Z counts as the closing line); segments are sampled then refined. */
export function nearestOnPath(
  cmds: PathCmd[],
  x: number,
  y: number,
): { cmd: number; t: number; dist: number } | null {
  const acc: { best: { cmd: number; t: number; dist: number } | null } = { best: null }
  const consider = (cmd: number, t: number) => {
    const p = pointAt(cmds, cmd, t)
    const d = Math.hypot(p.x - x, p.y - y)
    if (!acc.best || d < acc.best.dist) acc.best = { cmd, t, dist: d }
  }
  const stepsOf = (c: PathCmd) => (c.op === 'L' || c.op === 'Z' ? 16 : 32)
  for (let i = 1; i < cmds.length; i++) {
    const c = cmds[i]!
    if (c.op === 'M') continue
    if (c.op === 'Z' && near(startOf(cmds, i), subpathStart(cmds, i))) continue
    const steps = stepsOf(c)
    for (let k = 0; k <= steps; k++) consider(i, k / steps)
  }
  const b = acc.best
  if (!b) return null
  const span = 1 / stepsOf(cmds[b.cmd]!)
  const t0 = b.t
  for (let k = -8; k <= 8; k++) {
    const t = t0 + (k / 8) * span
    if (t >= 0 && t <= 1) consider(b.cmd, t)
  }
  return acc.best
}

/** Split segment `i` at parameter t, adding a vertex there (de Casteljau for curves). */
export function insertVertex(cmds: PathCmd[], i: number, t: number): PathCmd[] {
  const out = clone(cmds)
  const c = out[i]!
  const p0 = startOf(out, i)
  const u = 1 - t
  const lerp = (a: number, b: number) => a * u + b * t
  if (c.op === 'M') return out
  if (c.op === 'L' || c.op === 'Z') {
    const end = c.op === 'Z' ? subpathStart(out, i) : endOf(c)
    const m = { op: 'L' as const, pts: [lerp(p0.x, end.x), lerp(p0.y, end.y)] }
    if (c.op === 'Z') out.splice(i, 0, m)
    else out.splice(i, 1, m, { op: 'L', pts: [end.x, end.y] })
    return out
  }
  if (c.op === 'Q') {
    const [x1, y1, x, y] = c.pts as [number, number, number, number]
    const ax = lerp(p0.x, x1)
    const ay = lerp(p0.y, y1)
    const bx = lerp(x1, x)
    const by = lerp(y1, y)
    const mx = lerp(ax, bx)
    const my = lerp(ay, by)
    out.splice(i, 1, { op: 'Q', pts: [ax, ay, mx, my] }, { op: 'Q', pts: [bx, by, x, y] })
    return out
  }
  const [x1, y1, x2, y2, x, y] = c.pts as [number, number, number, number, number, number]
  const ax = lerp(p0.x, x1)
  const ay = lerp(p0.y, y1)
  const bx = lerp(x1, x2)
  const by = lerp(y1, y2)
  const cx = lerp(x2, x)
  const cy = lerp(y2, y)
  const abx = lerp(ax, bx)
  const aby = lerp(ay, by)
  const bcx = lerp(bx, cx)
  const bcy = lerp(by, cy)
  const mx = lerp(abx, bcx)
  const my = lerp(aby, bcy)
  out.splice(
    i,
    1,
    { op: 'C', pts: [ax, ay, abx, aby, mx, my] },
    { op: 'C', pts: [bcx, bcy, cx, cy, x, y] },
  )
  return out
}
