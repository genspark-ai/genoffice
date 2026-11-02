import { describe, expect, it } from 'vitest'
import type { MindMapTreeNode } from '../src/types'
import {
  countMindMapNodes,
  labelUnits,
  layoutMindMap,
  MIND_MAP_TOOL,
  MIND_MAP_LIMITS,
  MIND_MAP_METRICS,
  mindMapToMermaid,
  parseMindMapTree,
  wrapLabel,
} from '../src/mind-map'

interface Point {
  x: number
  y: number
}

interface EdgeCurve {
  start: Point
  c1: Point
  c2: Point
  end: Point
}

const CURVE_RE =
  /^M (-?[\d.]+) (-?[\d.]+) C (-?[\d.]+) (-?[\d.]+), (-?[\d.]+) (-?[\d.]+), (-?[\d.]+) (-?[\d.]+)$/

/** Geometry of one cubic connector, straight out of the path string. */
function parseCurve(d: string): EdgeCurve {
  const m = CURVE_RE.exec(d)
  if (!m) throw new Error(`unparsable edge path: ${d}`)
  const n = m.slice(1).map(Number)
  return {
    start: { x: n[0], y: n[1] },
    c1: { x: n[2], y: n[3] },
    c2: { x: n[4], y: n[5] },
    end: { x: n[6], y: n[7] },
  }
}

/** Point at t (0..1) on a cubic bezier. */
function cubicAt(curve: EdgeCurve, t: number): Point {
  const { start: p0, c1: p1, c2: p2, end: p3 } = curve
  const u = 1 - t
  const a = u * u * u
  const b = 3 * u * u * t
  const c = 3 * u * t * t
  const d = t * t * t
  return {
    x: a * p0.x + b * p1.x + c * p2.x + d * p3.x,
    y: a * p0.y + b * p1.y + c * p2.y + d * p3.y,
  }
}

describe('parseMindMapTree', () => {
  it('parses a valid tree and keeps in-range block citations', () => {
    const result = parseMindMapTree(
      {
        root: {
          label: 'Doc',
          children: [{ label: 'Suite', block: 2, children: [{ label: 'Suite' }] }],
        },
      },
      4,
    )
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.tree.children?.[0].children?.[0].label).toBe('Suite')
  })

  it('drops out-of-range block citations instead of failing', () => {
    const result = parseMindMapTree(
      { root: { label: 'Doc', children: [{ label: 'Ghost', block: 99 }] } },
      4,
    )
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.tree.children?.[0].block).toBeUndefined()
  })

  it('rejects a root without branches so the model retries with structure', () => {
    const result = parseMindMapTree({ root: { label: 'Doc' } }, 4)
    expect(result.ok).toBe(false)
    if (result.ok) return
    expect(result.error).toContain('no branches')
  })

  it('rejects a non-object root', () => {
    expect(parseMindMapTree({ root: 'nope' }, 4).ok).toBe(false)
    expect(parseMindMapTree({}, 4).ok).toBe(false)
  })

  it('trims trees beyond the node limit', () => {
    const wide = {
      label: 'Root',
      children: Array.from({ length: MIND_MAP_LIMITS.maxNodes + 20 }, (_, i) => ({
        label: `n${i}`,
      })),
    }
    const result = parseMindMapTree({ root: wide }, 4)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.trimmed).toBe(true)
  })

  it('trims trees deeper than the depth limit', () => {
    let node: Record<string, unknown> = { label: 'leaf' }
    for (let i = 0; i < MIND_MAP_LIMITS.maxDepth + 3; i++) {
      node = { label: `L${i}`, children: [node] }
    }
    const result = parseMindMapTree({ root: node }, 4)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.trimmed).toBe(true)
  })

  it('slices overlong labels', () => {
    const result = parseMindMapTree(
      { root: { label: 'x'.repeat(200), children: [{ label: 'ok' }] } },
      4,
    )
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.tree.label).toHaveLength(MIND_MAP_LIMITS.maxLabelChars)
  })
})

describe('countMindMapNodes', () => {
  it('counts every level', () => {
    expect(
      countMindMapNodes({
        label: 'a',
        children: [{ label: 'b', children: [{ label: 'c' }] }, { label: 'd' }],
      }),
    ).toBe(4)
  })
})

describe('labelUnits / wrapLabel', () => {
  it('counts CJK characters as double width', () => {
    expect(labelUnits('ab')).toBe(2)
    expect(labelUnits('风险')).toBe(4)
    expect(labelUnits('a风')).toBe(3)
  })

  it('wraps latin labels at word boundaries', () => {
    expect(wrapLabel('Task Management Automation', 16)).toEqual(['Task Management', 'Automation'])
  })

  it('keeps short labels on one line', () => {
    expect(wrapLabel('Risks', 20)).toEqual(['Risks'])
  })

  it('ellipsizes overflow on the second line', () => {
    const lines = wrapLabel('one two three four five six seven eight', 12)
    expect(lines).toHaveLength(2)
    expect(lines[1].endsWith('…')).toBe(true)
  })

  it('wraps CJK per character', () => {
    expect(wrapLabel('风险管理要点总结', 8)).toEqual(['风险管理', '要点总结'])
  })
})

describe('layoutMindMap: top-down', () => {
  const tree: MindMapTreeNode = {
    label: 'Root',
    children: [
      {
        label: 'A',
        children: [
          { label: 'A1', block: 1 },
          { label: 'A2', block: 3 },
        ],
      },
      { label: 'B', block: 0 },
    ],
  }

  it('centers parents over their children and stacks leaves left to right', () => {
    const { nodes, edges, width, height } = layoutMindMap(tree)
    const byId = new Map(nodes.map((n) => [n.id, n]))
    const root = byId.get('0')!
    const a = byId.get('0.0')!
    const b = byId.get('0.1')!
    const a1 = byId.get('0.0.0')!
    const a2 = byId.get('0.0.1')!
    const { nodeWidth: w, nodeHeight: h, depthGap, siblingGap, padding } = MIND_MAP_METRICS

    expect(a1.y).toBe(2 * (h + depthGap) + padding)
    expect(a2.y).toBe(a1.y)
    // B is a direct child of the root: one level up from A1/A2, level with A
    expect(b.y).toBe(h + depthGap + padding)
    expect(a.y).toBe(h + depthGap + padding)
    expect(root.y).toBe(padding)
    expect(a1.x).toBe(padding)

    // leaves are sequential; A centers over A1/A2; the root centers over A..B
    expect(a2.x).toBe(a1.x + w + siblingGap)
    expect(b.x).toBe(a2.x + w + siblingGap)
    expect(a.x).toBeCloseTo((a1.x + a2.x + w) / 2 - w / 2)
    expect(root.x).toBeCloseTo((a.x + b.x + w) / 2 - w / 2)

    // one edge per visible parent-child pair
    expect(edges).toHaveLength(4)
    expect(edges.every((e) => e.d.startsWith('M '))).toBe(true)

    expect(width).toBe(b.x + w + padding)
    expect(height).toBe(3 * (h + depthGap) - depthGap + padding * 2)
  })

  it('hides collapsed subtrees and keeps the hidden count', () => {
    const layout = layoutMindMap(tree, new Set(['0.0']))
    const a = layout.nodes.find((n) => n.id === '0.0')!
    expect(a.collapsed).toBe(true)
    expect(a.descendants).toBe(2)
    expect(layout.nodes).toHaveLength(3) // root + A(collapsed) + B
    expect(layout.edges).toHaveLength(2)
    expect(layout.nodes.some((n) => n.id === '0.0.0')).toBe(false)
  })

  it('lays out a single node without children', () => {
    const layout = layoutMindMap({ label: 'solo' })
    expect(layout.nodes).toHaveLength(1)
    expect(layout.edges).toHaveLength(0)
    expect(layout.width).toBe(MIND_MAP_METRICS.nodeWidth + MIND_MAP_METRICS.padding * 2)
  })

  it('is the default when no orientation is given', () => {
    const plain = layoutMindMap(tree)
    const explicit = layoutMindMap(tree, new Set(), { orientation: 'top-down' })
    expect(explicit).toEqual(plain)
  })
})

describe('layoutMindMap edge geometry', () => {
  // The reviewer's reproduction: Root -> {A -> {A1, A2}, B -> {B1}}.
  const repro: MindMapTreeNode = {
    label: 'Root',
    children: [
      { label: 'A', children: [{ label: 'A1' }, { label: 'A2' }] },
      { label: 'B', children: [{ label: 'B1' }] },
    ],
  }
  const other: MindMapTreeNode = {
    label: 'Root',
    children: [
      {
        label: 'A',
        children: [
          { label: 'A1', block: 1 },
          { label: 'A2', block: 3 },
        ],
      },
      { label: 'B', block: 0 },
    ],
  }

  it('anchors both ends on the node borders facing each other', () => {
    const { nodeWidth: w, nodeHeight: h } = MIND_MAP_METRICS
    for (const tree of [repro, other]) {
      const layout = layoutMindMap(tree)
      const byId = new Map(layout.nodes.map((n) => [n.id, n]))
      expect(layout.edges.length).toBeGreaterThan(0)

      for (const edge of layout.edges) {
        const [parentId, childId] = edge.id.split('->')
        const parent = byId.get(parentId)!
        const child = byId.get(childId)!
        const { start, end } = parseCurve(edge.d)

        // The layout is top-down (depth -> y), so a connector must leave the
        // parent's bottom border and arrive on the child's top border.
        expect(start.x).toBe(parent.x + w / 2)
        expect(start.y).toBe(parent.y + h)
        expect(end.x).toBe(child.x + w / 2)
        expect(end.y).toBe(child.y)
        expect(start.y).toBeLessThan(end.y) // the parent sits above the child
      }
    }
  })

  it('keeps every connector inside the vertical gap, crossing no node box', () => {
    const { nodeWidth: w, nodeHeight: h } = MIND_MAP_METRICS
    for (const tree of [repro, other]) {
      const layout = layoutMindMap(tree)

      for (const edge of layout.edges) {
        const curve = parseCurve(edge.d)
        const [parentId, childId] = edge.id.split('->')
        const parent = layout.nodes.find((n) => n.id === parentId)!
        const child = layout.nodes.find((n) => n.id === childId)!

        // control points stay between the two borders: the curve never overshoots
        for (const cp of [curve.c1, curve.c2]) {
          expect(cp.y).toBeGreaterThanOrEqual(curve.start.y)
          expect(cp.y).toBeLessThanOrEqual(curve.end.y)
        }

        // and no sample of the curve lands on top of any node (endpoints are the
        // only contact, by construction on the borders themselves)
        for (let t = 0.05; t <= 0.951; t += 0.05) {
          const p = cubicAt(curve, t)
          expect(p.y).toBeGreaterThan(parent.y + h)
          expect(p.y).toBeLessThan(child.y)
          for (const n of layout.nodes) {
            const hits = p.x > n.x && p.x < n.x + w && p.y > n.y && p.y < n.y + h
            expect(hits, `${edge.id} crosses node ${n.id} at t=${t.toFixed(2)}`).toBe(false)
          }
        }
      }
    }
  })
})

describe('layoutMindMap: left-right orientation', () => {
  const { nodeWidth: W, nodeHeight: H, depthGap, padding } = MIND_MAP_METRICS

  const tree: MindMapTreeNode = {
    label: 'Root',
    children: [
      { label: 'A', children: [{ label: 'A1' }, { label: 'A2' }] },
      { label: 'B', children: [{ label: 'B1' }] },
    ],
  }

  const layoutOf = (t: MindMapTreeNode, collapsed = new Set<string>()) =>
    layoutMindMap(t, collapsed, { orientation: 'left-right' })

  it('puts the root between the two halves', () => {
    const layout = layoutOf(tree)
    const byId = new Map(layout.nodes.map((n) => [n.id, n]))
    const root = byId.get('0')!
    const a = byId.get('0.0')!
    const b = byId.get('0.1')!
    const a1 = byId.get('0.0.0')!
    const a2 = byId.get('0.0.1')!
    const b1 = byId.get('0.1.0')!

    // the first branch goes right, the second one left
    expect(a.x).toBeGreaterThan(root.x)
    expect(a1.x).toBeGreaterThan(a.x)
    expect(a2.x).toBe(a1.x)
    expect(b.x).toBeLessThan(root.x)
    expect(b1.x).toBeLessThan(b.x)

    // every column is one node width plus one gap apart
    expect(a.x - root.x).toBe(W / 2 + depthGap + W / 2)
    expect(a1.x - a.x).toBe(W + depthGap)
    expect(root.x - b.x).toBe(W / 2 + depthGap + W / 2)
    expect(b.x - b1.x).toBe(W + depthGap)
  })

  it('keeps the root centred on the outermost leaves', () => {
    const layout = layoutOf(tree)
    const byId = new Map(layout.nodes.map((n) => [n.id, n]))
    const root = byId.get('0')!
    const leaves = layout.nodes.filter((n) => !n.hasChildren)
    const top = Math.min(...leaves.map((n) => n.y))
    const bottom = Math.max(...leaves.map((n) => n.y))
    expect(root.y).toBe((top + bottom) / 2)
  })

  it('splits the branches by leaf count, so a heavy branch does not tilt the map', () => {
    // A carries five leaves, B only one: A goes right, B goes left, and the
    // root stays in the middle of the map instead of following the heavy side
    const heavy: MindMapTreeNode = {
      label: 'Root',
      children: [
        {
          label: 'A',
          children: Array.from({ length: 5 }, (_, i) => ({ label: `A${i}` })),
        },
        { label: 'B' },
      ],
    }
    const layout = layoutOf(heavy)
    const byId = new Map(layout.nodes.map((n) => [n.id, n]))
    const root = byId.get('0')!
    const a = byId.get('0.0')!
    const b = byId.get('0.1')!

    expect(a.x).toBeGreaterThan(root.x)
    expect(b.x).toBeLessThan(root.x)
    // the root is not dragged to the heavy side
    expect(Math.abs(root.x - a.x)).toBeLessThan(Math.abs(root.x - b.x) + 1)
    // but the heavy side is the taller one
    const aLeaves = layout.nodes.filter((n) => n.id.startsWith('0.0.') && !n.hasChildren)
    expect(aLeaves.length).toBe(5)
  })

  it('anchors every connector on the borders that face each other', () => {
    const layout = layoutOf(tree)
    const byId = new Map(layout.nodes.map((n) => [n.id, n]))

    for (const edge of layout.edges) {
      const [parentId, childId] = edge.id.split('->')
      const parent = byId.get(parentId)!
      const child = byId.get(childId)!
      const { start, end } = parseCurve(edge.d)

      const parentOnLeft = parent.x < child.x
      // the connector leaves the facing border and arrives on the other one
      expect(start.x).toBe(parentOnLeft ? parent.x + W : parent.x)
      expect(start.y).toBe(parent.y + H / 2)
      expect(end.x).toBe(parentOnLeft ? child.x : child.x + W)
      expect(end.y).toBe(child.y + H / 2)
      // the horizontal span is exactly one gap
      expect(Math.abs(end.x - start.x)).toBe(depthGap)
    }
  })

  it('keeps every connector inside the gap, crossing no node box', () => {
    const layout = layoutOf(tree)

    for (const edge of layout.edges) {
      const curve = parseCurve(edge.d)

      // control points stay inside the gap
      for (const cp of [curve.c1, curve.c2]) {
        expect(cp.x).toBeGreaterThanOrEqual(Math.min(curve.start.x, curve.end.x))
        expect(cp.x).toBeLessThanOrEqual(Math.max(curve.start.x, curve.end.x))
      }

      for (let t = 0.05; t <= 0.951; t += 0.05) {
        const p = cubicAt(curve, t)
        for (const n of layout.nodes) {
          const hits = p.x > n.x && p.x < n.x + W && p.y > n.y && p.y < n.y + H
          expect(hits, `${edge.id} crosses node ${n.id} at t=${t.toFixed(2)}`).toBe(false)
        }
      }
    }
  })

  it('never overlaps two node boxes', () => {
    const layout = layoutOf(tree)
    for (let i = 0; i < layout.nodes.length; i++) {
      for (let j = i + 1; j < layout.nodes.length; j++) {
        const a = layout.nodes[i]
        const b = layout.nodes[j]
        const overlaps = a.x < b.x + W && b.x < a.x + W && a.y < b.y + H && b.y < a.y + H
        expect(overlaps, `${a.id} overlaps ${b.id}`).toBe(false)
      }
    }
  })

  it('reports a width and a height that fit the boxes', () => {
    const layout = layoutOf(tree)
    const byId = new Map(layout.nodes.map((n) => [n.id, n]))
    const root = byId.get('0')!
    const leftMost = Math.min(...layout.nodes.map((n) => n.x))
    const rightMost = Math.max(...layout.nodes.map((n) => n.x + W))
    const bottomMost = Math.max(...layout.nodes.map((n) => n.y + H))

    expect(layout.width).toBe(rightMost - leftMost + padding * 2)
    expect(layout.height).toBeGreaterThanOrEqual(bottomMost + padding)
    // the root is not glued to an edge
    expect(root.x).toBeGreaterThan(leftMost)
    expect(root.x + W).toBeLessThan(rightMost)
  })

  it('hides a collapsed subtree and reflows the map around it', () => {
    const collapsed = new Set(['0.0'])
    const layout = layoutOf(tree, collapsed)
    const byId = new Map(layout.nodes.map((n) => [n.id, n]))

    expect(layout.nodes.some((n) => n.id === '0.0.0')).toBe(false)
    expect(byId.get('0.0')!.collapsed).toBe(true)
    expect(byId.get('0.0')!.descendants).toBe(2)
    // the collapsed branch still occupies its column, the root stays centred
    expect(byId.get('0.0')!.x).toBeGreaterThan(byId.get('0')!.x)
    // root -> A (collapsed), root -> B, B -> B1
    expect(layout.edges).toHaveLength(3)
  })

  it('degrades to a single half when the root has one branch', () => {
    const single: MindMapTreeNode = {
      label: 'Root',
      children: [{ label: 'A', children: [{ label: 'A1' }] }],
    }
    const layout = layoutOf(single)
    const byId = new Map(layout.nodes.map((n) => [n.id, n]))
    expect(byId.get('0.1')).toBeUndefined()
    expect(byId.get('0.0')!.x).toBeGreaterThan(byId.get('0')!.x)
    expect(layout.edges).toHaveLength(2)
  })

  it('lays a single node out on its own', () => {
    const layout = layoutOf({ label: 'solo' })
    expect(layout.nodes).toHaveLength(1)
    expect(layout.edges).toHaveLength(0)
    expect(layout.width).toBe(W + padding * 2)
  })

  it('is deterministic', () => {
    expect(layoutOf(tree)).toEqual(layoutOf(tree))
  })
})

describe('model-facing mind map text', () => {
  // The public-repo hygiene gate rejects Han characters outside the i18n tables,
  // and "mind map" in English matches the user's request in any language.
  const HAN = /[\u3400-\u9FFF\uF900-\uFAFF\u3000-\u303F\uFF00-\uFFEF]/

  it('keeps the tool description free of CJK', () => {
    expect(MIND_MAP_TOOL.description).not.toMatch(HAN)
    expect(JSON.stringify(MIND_MAP_TOOL.inputSchema)).not.toMatch(HAN)
  })
})

describe('mindMapToMermaid', () => {
  it('emits the tree as an indented mindmap diagram, root first', () => {
    const tree: MindMapTreeNode = {
      label: 'Doc',
      children: [{ label: 'A', children: [{ label: 'A1' }] }, { label: 'B' }],
    }
    expect(mindMapToMermaid(tree)).toBe('mindmap\n  Doc\n    A\n      A1\n    B')
  })

  it('strips node-shape delimiters so mermaid reads labels as plain text', () => {
    expect(mindMapToMermaid({ label: 'Q1 (sales) [draft]' })).toBe('mindmap\n  Q1 sales draft')
  })

  it('strips fence and directive syntax that would break out of the mermaid block', () => {
    expect(mindMapToMermaid({ label: '```%% x ::: y' })).toBe('mindmap\n  x y')
  })

  it('falls back to Topic when a label is only delimiters', () => {
    expect(mindMapToMermaid({ label: '()' })).toBe('mindmap\n  Topic')
  })
})
