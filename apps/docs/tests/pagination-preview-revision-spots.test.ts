import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { revGroupsOf } from '../src/renderer/editor/margin-annotations'
import {
  measureRevisionSpots,
  mergeBalloonLists,
  stackBalloons,
  type BalloonItem,
} from '../src/renderer/components/PaginationPreview'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const text = (t: string, marks?: JsonNode['marks']): JsonNode => ({
  type: 'text',
  text: t,
  ...(marks ? { marks } : {}),
})
const del = { type: 'del', attrs: { author: 'Bob', date: '2026-07-02T10:00:00Z' } }
const para = (attrs: Record<string, unknown>, ...content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, ...attrs },
  content,
})

describe('revGroupsOf', () => {
  it('merges a fully deleted paragraph with the next paragraph leading deletion across the deleted mark', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          para({}, text('keep '), text('gone', [del]), text(' keep')),
          para({ paraMarkDel: JSON.stringify({ author: 'Bob' }) }, text('whole', [del])),
          para({}, text('lead', [del]), text(' rest')),
        ],
      },
    })
    const groups = revGroupsOf(editor.state.doc)
    expect(groups.map((g) => [g.kind, g.text])).toEqual([
      ['del', 'gone'],
      ['del', 'whole¶lead'],
    ])
    editor.destroy()
  })
})

type Rect = { top: number; bottom: number; left: number; right: number; height: number }
const rect = (top: number, left: number, w = 40, h = 16): Rect => ({
  top,
  bottom: top + h,
  left,
  right: left + w,
  height: h,
})

function block(idx: number, rects: Rect[]): HTMLElement {
  const el = document.createElement('div')
  el.dataset.idx = String(idx)
  el.getClientRects = () => rects as unknown as DOMRectList
  return el
}

describe('measureRevisionSpots', () => {
  it('balloons deleted runs of protected blocks from their XML; a collapsed block anchors at its visible neighbour', () => {
    const pm = document.createElement('div')
    pm.append(block(3, [rect(100, 0, 300)]), block(4, []), block(5, [rect(140, 0, 300)]))
    pm.getClientRects = () => [rect(0, 0, 300, 400)] as unknown as DOMRectList
    const xml = (t: string) =>
      `<w:p><w:r><w:t>live</w:t></w:r><w:del w:id="1" w:author="A"><w:r><w:delText>${t}</w:delText></w:r></w:del></w:p>`
    const blocks = [
      { type: 'passthrough', docxIndex: 3, originalXml: xml('6 &amp; 7') },
      {
        type: 'passthrough',
        docxIndex: 4,
        originalXml: xml('Introduction</w:delText></w:r><w:r><w:tab/><w:delText>1'),
      },
      { type: 'paragraph', docxIndex: 5, originalXml: xml('marked in the editor instead') },
    ]
    const spots = measureRevisionSpots(null, pm, blocks, 0, 1, 0)
    expect(spots).toEqual([
      { key: 'x3', kind: 'del', text: '6 & 7', top: 100, endX: 300, endY: 115 },
      { key: 'x4', kind: 'del', text: 'Introduction 1', top: 100, endX: 300, endY: 115 },
    ])
  })
})

describe('stackBalloons', () => {
  const item = (over: Partial<BalloonItem> & { key: string }): BalloonItem => ({
    anchorTop: 0,
    seq: 0,
    endX: 0,
    height: 20,
    sticky: false,
    ...over,
  })

  it('stacks by sequence (not anchor X), pushing overlaps down by the gap', () => {
    const placed = stackBalloons(
      [
        item({ key: 'b', seq: 1, endX: 20, anchorTop: 10 }),
        item({ key: 'a', seq: 0, endX: 50, anchorTop: 10 }),
        item({ key: 'c', seq: 2, anchorTop: 100 }),
      ],
      0,
      1000,
      () => 4,
    )
    expect(placed.map((p) => [p.key, p.top])).toEqual([
      ['a', 10],
      ['b', 34],
      ['c', 100],
    ])
  })

  it('pulls a stack that runs past the band bottom up above its anchors', () => {
    const placed = stackBalloons(
      [
        item({ key: 'a', anchorTop: 900, seq: 0, height: 60 }),
        item({ key: 'b', anchorTop: 950, seq: 1, height: 60 }),
      ],
      0,
      1000,
      () => 4,
    )
    expect(placed.map((p) => [p.key, p.top])).toEqual([
      ['a', 876],
      ['b', 940],
    ])
  })

  it('compacts revision balloons, then drops the last ones, before letting comments spill', () => {
    const placed = stackBalloons(
      [
        item({ key: 'comment', anchorTop: 0, seq: 0, height: 80, sticky: true }),
        item({ key: 'rev1', anchorTop: 10, seq: 1, height: 40, compactHeight: 20 }),
        item({ key: 'rev2', anchorTop: 20, seq: 2, height: 40, compactHeight: 20 }),
      ],
      0,
      130,
      () => 4,
    )
    expect(placed.map((p) => [p.key, p.top, p.compact])).toEqual([
      ['comment', 0, false],
      ['rev1', 84, true],
      ['rev2', 108, true],
    ])
    const dropped = stackBalloons(
      [
        item({ key: 'comment', anchorTop: 0, seq: 0, height: 80, sticky: true }),
        item({ key: 'rev1', anchorTop: 10, seq: 1, height: 40, compactHeight: 20 }),
        item({ key: 'rev2', anchorTop: 20, seq: 2, height: 40, compactHeight: 20 }),
      ],
      0,
      110,
      () => 4,
    )
    expect(dropped.map((p) => p.key)).toEqual(['comment', 'rev1'])
  })
})

describe('mergeBalloonLists', () => {
  it('interleaves comments and revisions by order line, same line left to right, keeping each list order', () => {
    const comments = [
      { key: 'c1', orderTop: 100, endX: 300 },
      // document-later comment on the same line with a smaller range end
      { key: 'c2', orderTop: 100, endX: 120 },
      // next-column comment: visually higher (anchorTop 40) but ordered after by its floored orderTop
      { key: 'c3', orderTop: 300, anchorTop: 40, endX: 50 },
    ]
    const revs = [
      { key: 'r1', orderTop: 100, endX: 200 },
      { key: 'r2', orderTop: 200, endX: 10 },
    ]
    const merged = mergeBalloonLists(comments, revs)
    expect(merged.map((m) => [m.key, m.seq])).toEqual([
      ['r1', 0],
      ['c1', 1],
      ['c2', 2],
      ['r2', 3],
      ['c3', 4],
    ])
  })
})
