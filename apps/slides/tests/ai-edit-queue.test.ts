/**
 * AI edit queue anchoring: a queued item stores element ids plus a page hint,
 * so it must survive page reordering, follow the durable id after a re-render,
 * and be reported as unrunnable when its element is gone or too deeply nested.
 */
import { describe, it, expect } from 'vitest'
import type { PlacedBox, RenderSlide } from '@genoffice/pptx-render'
import {
  buildPageInstruction,
  buildSelectionInstruction,
  describeNode,
  groupByPage,
  resolveQueueItem,
  type EditQueueItem,
} from '../src/renderer/ai/edit-queue'

const box = (x: number, y: number, w: number, h: number): PlacedBox => ({
  x,
  y,
  w,
  h,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
  centerX: x + w / 2,
  centerY: y + h / 2,
})

const text = (s: string) => ({ lines: [{ runs: [{ text: s, fontSizePx: 24 }] }] })

function slide(nodes: unknown[]): RenderSlide {
  return { widthPx: 1280, heightPx: 720, nodes } as unknown as RenderSlide
}

const title = {
  id: 't1',
  sourceId: 't1',
  durableId: 'dur-title',
  type: 'shape',
  box: box(80, 60, 600, 120),
  text: text('Quarterly review'),
}
const picture = { id: 'p1', sourceId: 'p1', type: 'picture', box: box(700, 60, 400, 300) }
const deepGroup = {
  id: 'g1',
  sourceId: 'g1',
  type: 'group',
  box: box(0, 400, 400, 200),
  children: [
    {
      id: 'g2',
      sourceId: 'g2',
      type: 'group',
      box: box(0, 0, 200, 200),
      children: [{ id: 'inner', sourceId: 'inner', type: 'shape', box: box(0, 0, 100, 100) }],
    },
  ],
}

const item = (over: Partial<EditQueueItem> = {}): EditQueueItem => ({
  key: 'k1',
  slideIndex: 0,
  targets: [{ id: 'dur-title', sourceId: 't1', type: 'shape' }],
  instruction: 'Break into two lines',
  status: 'pending',
  ...over,
})

describe('resolveQueueItem', () => {
  it('resolves through the durable id on the cached page', () => {
    const r = resolveQueueItem([slide([title, picture])], item())
    expect(r.ok).toBe(true)
    if (r.ok) {
      expect(r.slideIndex).toBe(0)
      expect(r.nodes[0]?.sourceId).toBe('t1')
    }
  })

  it('re-derives the page after slides are reordered', () => {
    const decks = [slide([picture]), slide([]), slide([title])]
    const r = resolveQueueItem(decks, item())
    expect(r.ok).toBe(true)
    if (r.ok) expect(r.slideIndex).toBe(2)
  })

  it('reports a deleted element instead of guessing', () => {
    const r = resolveQueueItem([slide([picture])], item())
    expect(r).toMatchObject({ ok: false, reason: 'deleted' })
  })

  it('flags a durable id that now matches on more than one page', () => {
    // Copies keep the slide-scoped durable id but get their own render-tree id
    const copy = (sourceId: string) => ({ ...title, id: sourceId, sourceId })
    const r = resolveQueueItem([slide([picture]), slide([copy('x1')]), slide([copy('x2')])], item())
    expect(r).toMatchObject({ ok: false, reason: 'ambiguous' })
  })

  it('breaks a durable-id tie in favour of the annotated page', () => {
    const copy = (sourceId: string) => ({ ...title, id: sourceId, sourceId })
    const r = resolveQueueItem([slide([copy('x1')]), slide([copy('x2')])], item())
    expect(r).toMatchObject({ ok: true, slideIndex: 0 })
  })

  it('rejects elements nested deeper than one group level', () => {
    const r = resolveQueueItem(
      [slide([deepGroup])],
      item({ targets: [{ id: 'inner', sourceId: 'inner', type: 'shape' }] }),
    )
    expect(r).toMatchObject({ ok: false, reason: 'nested' })
  })

  it('will not match an id whose element kind changed', () => {
    const other = {
      id: 't1',
      sourceId: 't1',
      durableId: 'dur-title',
      type: 'picture',
      box: box(0, 0, 10, 10),
    }
    const r = resolveQueueItem([slide([other])], item())
    expect(r).toMatchObject({ ok: false, reason: 'deleted' })
  })

  it('accepts a direct child of a top-level group', () => {
    const group = {
      id: 'g0',
      sourceId: 'g0',
      type: 'group',
      box: box(0, 0, 200, 200),
      children: [{ id: 'c1', sourceId: 'c1', type: 'shape', box: box(0, 0, 100, 100) }],
    }
    const r = resolveQueueItem(
      [slide([group])],
      item({ targets: [{ id: 'c1', sourceId: 'c1', type: 'shape' }] }),
    )
    expect(r.ok).toBe(true)
  })
})

describe('grouping and prompt assembly', () => {
  it('groups same-page edits together and orders pages', () => {
    const decks = [slide([picture]), slide([title])]
    const resolved = [
      resolveQueueItem(decks, item({ key: 'a' })),
      resolveQueueItem(decks, item({ key: 'b', instruction: 'Shorten it' })),
      resolveQueueItem(
        decks,
        item({
          key: 'c',
          slideIndex: 0,
          targets: [{ id: 'p1', sourceId: 'p1', type: 'picture' }],
        }),
      ),
    ]
    const groups = groupByPage(resolved)
    expect(groups.map((g) => g.slideIndex)).toEqual([0, 1])
    expect(groups[1]?.entries).toHaveLength(2)
  })

  it('names every target and pins the run to one page', () => {
    const decks = [slide([title])]
    const groups = groupByPage([resolveQueueItem(decks, item())])
    const prompt = buildPageInstruction(groups[0]!, 1, 1)
    expect(prompt).toContain('slideIndex=0')
    expect(prompt).toContain('dur-title')
    expect(prompt).toContain('Quarterly review')
    expect(prompt).toContain('Break into two lines')
  })

  it('freezes Send now to the selected durable id and visible text', () => {
    const prompt = buildSelectionInstruction(
      0,
      [{ id: 'dur-title', desc: describeNode(title as never) }],
      'Make the font red',
    )
    expect(prompt).toContain('slideIndex=0')
    expect(prompt).toContain('dur-title')
    expect(prompt).toContain('Quarterly review')
    expect(prompt).toContain('Make the font red')
    expect(prompt).toContain('do not modify unlisted elements')
  })
})

describe('describeNode', () => {
  it('summarizes text content and table geometry', () => {
    expect(describeNode(title as never)).toMatchObject({
      type: 'shape',
      text: 'Quarterly review',
    })
    const table = {
      id: 'tb',
      sourceId: 'tb',
      type: 'table',
      box: box(0, 0, 300, 200),
      cells: [{ x: 0, y: 0, text: text('Region') }],
      gridX: [0, 100, 200, 300],
      gridY: [0, 100, 200],
    }
    expect(describeNode(table as never)).toMatchObject({ type: 'table', rows: 2, cols: 3 })
  })
})
