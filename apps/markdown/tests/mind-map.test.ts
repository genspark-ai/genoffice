import { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it } from 'vitest'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { executeTool } from '../src/renderer/ai/tools'
import { createMarkdownSkill } from '../src/renderer/ai/markdown-skill'
import {
  countMindMapNodes,
  labelUnits,
  layoutMindMap,
  MIND_MAP_LIMITS,
  MIND_MAP_METRICS,
  MIND_MAP_TOOL,
  parseMindMapTree,
  wrapLabel,
} from '../src/renderer/ai/mind-map'
import type { MindMapTreeNode } from '@genoffice/agent-core'

// Undestroyed views leave DOMObserver flush timers that fire after jsdom teardown
// ("document is not defined" unhandled error) — destroy every editor we create.
const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

/** 0 h1 | 1 p | 2 h2 | 3 p */
function createEditor(): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content: '',
  })
  editor.commands.setContent(
    '# Chapter 1 Overview\n\nGenSpark is an AI office suite.\n\n## Risk Notes\n\nThis document is for reference only.',
    { contentType: 'markdown' },
  )
  editors.push(editor)
  return editor
}

const call = (name: string, input: Record<string, unknown> = {}) => ({
  id: 't1',
  name,
  input,
})

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
          children: [
            { label: 'Chapter 1', block: 0, children: [{ label: 'Suite', block: 1 }] },
            { label: 'Risks', block: 2 },
          ],
        },
      },
      4,
    )
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.tree.label).toBe('Doc')
    expect(result.nodeCount).toBe(4)
    expect(result.depth).toBe(2)
    expect(result.tree.children?.[0].block).toBe(0)
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
        label: `Node ${i}`,
      })),
    }
    const result = parseMindMapTree({ root: wide }, 4)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.nodeCount).toBe(MIND_MAP_LIMITS.maxNodes)
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
    expect(result.depth).toBe(MIND_MAP_LIMITS.maxDepth)
    expect(result.trimmed).toBe(true)
  })

  it('slices overlong labels', () => {
    const result = parseMindMapTree(
      { root: { label: 'x'.repeat(200), children: [{ label: 'ok' }] } },
      4,
    )
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.tree.label.length).toBe(MIND_MAP_LIMITS.maxLabelChars)
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
    expect(labelUnits('\u98CE\u9669')).toBe(4)
    expect(labelUnits('a\u98CE')).toBe(3)
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
    expect(wrapLabel('\u98CE\u9669\u7BA1\u7406\u8981\u70B9\u603B\u7ED3', 8)).toEqual([
      '\u98CE\u9669\u7BA1\u7406',
      '\u8981\u70B9\u603B\u7ED3',
    ])
  })
})

describe('layoutMindMap', () => {
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

describe('model-facing mind map text', () => {
  // The public-repo hygiene gate rejects Han characters outside the i18n tables,
  // and "mind map" in English matches the user's request in any language.
  const HAN = /[\u3400-\u9FFF\uF900-\uFAFF\u3000-\u303F\uFF00-\uFFEF]/

  it('keeps the tool description free of CJK', () => {
    expect(MIND_MAP_TOOL.description).not.toMatch(HAN)
    expect(JSON.stringify(MIND_MAP_TOOL.inputSchema)).not.toMatch(HAN)
  })

  it('keeps the system prompt free of CJK', () => {
    const systemPrompt = createMarkdownSkill(() => null).systemPrompt
    expect(systemPrompt).toContain('generate_mindmap')
    for (const line of systemPrompt.split('\n')) {
      expect(line, `CJK in prompt line: ${line}`).not.toMatch(HAN)
    }
  })
})

describe('generate_mindmap tool', () => {
  it('returns the tree on the display side channel without mutating the document', async () => {
    const editor = createEditor()
    const before = editor.state.doc.childCount
    expect(before).toBe(4)
    const exec = await executeTool(
      editor,
      call('generate_mindmap', {
        focus: 'risk controls',
        root: {
          label: 'Quarterly Report',
          children: [
            { label: 'Overview', block: 0, children: [{ label: 'Suite', block: 1 }] },
            { label: 'Risks', block: 2 },
          ],
        },
      }),
    )
    expect(exec.isError).toBeUndefined()
    expect(exec.mutated).toBe(false)
    expect(exec.summary).toBe('\u751F\u6210\u601D\u7EF4\u5BFC\u56FE')
    expect(exec.display?.kind).toBe('mindmap')
    expect(exec.display?.tree?.label).toBe('Quarterly Report')
    expect(exec.output).toContain('4 nodes')
    expect(exec.output).toContain('risk controls')
    expect(editor.state.doc.childCount).toBe(before)
  })

  it('reports an error when the model sends no branches', async () => {
    const editor = createEditor()
    const exec = await executeTool(editor, call('generate_mindmap', { root: { label: 'Lonely' } }))
    expect(exec.isError).toBe(true)
    expect(exec.display).toBeUndefined()
  })
})

describe('MIND_MAP_TOOL schema', () => {
  it('gives every array an items schema (Gemini refuses the whole request without one)', () => {
    // genoffice#2059: the leaf `children` array carried only a description,
    // and the API answered "...properties[children].items: missing field".
    const walk = (node: unknown): void => {
      if (typeof node !== 'object' || node === null) return
      if (Array.isArray(node)) {
        node.forEach(walk)
        return
      }
      const record = node as Record<string, unknown>
      if (record.type === 'array') expect(record.items).toBeDefined()
      for (const value of Object.values(record)) walk(value)
    }
    walk(MIND_MAP_TOOL.inputSchema)
  })
})
