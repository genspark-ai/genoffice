import { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it } from 'vitest'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { executeTool } from '../src/renderer/ai/tools'
import {
  countMindMapNodes,
  labelUnits,
  layoutMindMap,
  MIND_MAP_LIMITS,
  MIND_MAP_METRICS,
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
    expect(labelUnits('风险')).toBe(4)
    expect(labelUnits('a风')).toBe(3)
  })

  it('wraps latin labels at word boundaries', () => {
    expect(wrapLabel('Task Management Automation', 16)).toEqual([
      'Task Management',
      'Automation',
    ])
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

    expect(a1.y).toBe(2 * (h + depthGap))
    expect(a2.y).toBe(a1.y)
    // B is a direct child of the root: one level up from A1/A2, level with A
    expect(b.y).toBe(h + depthGap)
    expect(a.y).toBe(h + depthGap)
    expect(root.y).toBe(0)

    // leaves are sequential; A centers over A1/A2; the root centers over A..B
    expect(a2.x).toBe(a1.x + w + siblingGap)
    expect(b.x).toBe(a2.x + w + siblingGap)
    expect(a.x).toBeCloseTo((a1.x + a2.x + w) / 2 - w / 2)
    expect(root.x).toBeCloseTo((a.x + b.x + w) / 2 - w / 2)

    // one edge per visible parent-child pair
    expect(edges).toHaveLength(4)
    expect(edges.every((e) => e.d.startsWith('M '))).toBe(true)

    expect(width).toBe(b.x + w + padding * 2)
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
    expect(exec.summary).toBe('生成思维导图')
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
