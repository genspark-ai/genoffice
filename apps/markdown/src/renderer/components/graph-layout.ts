/**
 * A small force-directed layout for the wiki graph — no dependency, fixed
 * iteration budget, deterministic from the input alone (nodes start on a
 * circle, so the same vault draws the same picture twice).
 *
 * The forces are the classic trio: pair repulsion (every node pushes every
 * other away, the term that spreads connected clusters apart), edge springs
 * (a link wants its two ends at `IDEAL` distance), and a weak pull to the
 * center (so orphans do not drift off canvas). Positions clamp inside the
 * viewport with a margin.
 */
export interface LayoutNode {
  path: string
}

export interface LayoutEdge {
  from: string
  to: string
}

export interface LayoutPoint {
  x: number
  y: number
}

/** resting length of an edge, in px */
const IDEAL = 110
/** repulsion strength, px² */
const REPULSION = 30_000
/** spring strength, 0..1 */
const SPRING = 0.1
/** pull toward the viewport center, 0..1 */
const GRAVITY = 0.04
/** iterations of the simulation */
const TICKS = 260
/** velocity damping per tick */
const DAMPING = 0.82
/** px kept clear of the viewport edges */
const MARGIN = 60

export function layoutGraph(
  nodes: readonly LayoutNode[],
  edges: readonly LayoutEdge[],
  width: number,
  height: number,
  ticks = TICKS,
): Map<string, LayoutPoint> {
  const out = new Map<string, LayoutPoint>()
  if (nodes.length === 0 || width <= 0 || height <= 0) return out

  const cx = width / 2
  const cy = height / 2
  const radius = Math.max(40, Math.min(cx, cy) - MARGIN)
  const pos = new Map<string, LayoutPoint>()
  nodes.forEach((node, i) => {
    const angle = (2 * Math.PI * i) / nodes.length - Math.PI / 2
    pos.set(node.path, {
      x: cx + radius * Math.cos(angle),
      y: cy + radius * Math.sin(angle),
    })
  })

  const vx = new Map<string, number>()
  const vy = new Map<string, number>()

  for (let tick = 0; tick < ticks; tick++) {
    const cooling = 1 - tick / ticks
    // repulsion: every pair pushes apart
    for (let i = 0; i < nodes.length; i++) {
      const a = nodes[i]!.path
      const p = pos.get(a)!
      let fx = (cx - p.x) * GRAVITY
      let fy = (cy - p.y) * GRAVITY
      for (let j = 0; j < nodes.length; j++) {
        if (j === i) continue
        const q = pos.get(nodes[j]!.path)!
        let dx = p.x - q.x
        let dy = p.y - q.y
        let d2 = dx * dx + dy * dy
        if (d2 < 1) {
          // same-spot nodes push apart on a stable diagonal, not randomly
          dx = 1
          dy = i % 2 === 0 ? 1 : -1
          d2 = 1
        }
        const force = REPULSION / d2
        const d = Math.sqrt(d2)
        fx += (dx / d) * force
        fy += (dy / d) * force
      }
      vx.set(a, ((vx.get(a) ?? 0) + fx) * DAMPING * cooling)
      vy.set(a, ((vy.get(a) ?? 0) + fy) * DAMPING * cooling)
    }
    // springs: a link pulls its two ends toward IDEAL distance
    for (const edge of edges) {
      const pa = pos.get(edge.from)
      const pb = pos.get(edge.to)
      if (!pa || !pb) continue
      const dx = pb.x - pa.x
      const dy = pb.y - pa.y
      const d = Math.max(1, Math.hypot(dx, dy))
      const force = (SPRING * (d - IDEAL)) / 2
      const ux = (dx / d) * force
      const uy = (dy / d) * force
      vx.set(edge.from, (vx.get(edge.from) ?? 0) + ux)
      vy.set(edge.from, (vy.get(edge.from) ?? 0) + uy)
      vx.set(edge.to, (vx.get(edge.to) ?? 0) - ux)
      vy.set(edge.to, (vy.get(edge.to) ?? 0) - uy)
    }
    for (const node of nodes) {
      const p = pos.get(node.path)!
      p.x = Math.min(width - MARGIN / 2, Math.max(MARGIN / 2, p.x + (vx.get(node.path) ?? 0)))
      p.y = Math.min(height - MARGIN / 2, Math.max(MARGIN / 2, p.y + (vy.get(node.path) ?? 0)))
    }
  }

  for (const [path, p] of pos) out.set(path, { x: p.x, y: p.y })
  return out
}
