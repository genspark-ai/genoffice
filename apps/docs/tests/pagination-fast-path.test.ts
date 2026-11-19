import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blocksKeepSlicing,
  crossesPage,
  singleInlineEditIndex,
} from '../src/renderer/editor/pagination-fast-path'
import type { BlockBox, PageSlice } from '../src/renderer/pagination-types'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}

const text = (t: string): JsonNode => ({ type: 'text', text: t })
const para = (t: string): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: [text(t)],
})

const liveEditors: Editor[] = []
afterEach(() => {
  for (const e of liveEditors.splice(0)) e.destroy()
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  liveEditors.push(editor)
  return editor
}

/** doc position just inside the top-level block `index`, offset by `offset` chars */
function inside(editor: Editor, index: number, offset = 1): number {
  let pos = 0
  for (let i = 0; i < index; i++) pos += editor.state.doc.child(i).nodeSize
  return pos + offset
}

describe('singleInlineEditIndex', () => {
  it('returns the block index for typing inside one paragraph', () => {
    const editor = createEditor([para('first'), para('second'), para('third')])
    const tr = editor.state.tr.insertText('x', inside(editor, 1, 3))
    expect(singleInlineEditIndex(tr)).toBe(1)
  })

  it('accepts a deletion and a mark toggle inside one paragraph', () => {
    const editor = createEditor([para('hello world'), para('tail')])
    const del = editor.state.tr.delete(inside(editor, 0, 1), inside(editor, 0, 4))
    expect(singleInlineEditIndex(del)).toBe(0)
    const bold = editor.state.schema.marks.bold
    if (bold) {
      const mark = editor.state.tr.addMark(
        inside(editor, 0, 1),
        inside(editor, 0, 6),
        bold.create(),
      )
      expect(singleInlineEditIndex(mark)).toBe(0)
    }
  })

  it('rejects edits spanning two paragraphs', () => {
    const editor = createEditor([para('first'), para('second')])
    const tr = editor.state.tr.delete(inside(editor, 0, 4), inside(editor, 1, 3))
    expect(singleInlineEditIndex(tr)).toBeNull()
  })

  it('rejects a paragraph split (block count changes)', () => {
    const editor = createEditor([para('first'), para('second')])
    const tr = editor.state.tr.split(inside(editor, 0, 3))
    expect(singleInlineEditIndex(tr)).toBeNull()
  })

  it('rejects inserting a non-text inline node', () => {
    const editor = createEditor([para('first')])
    const br = editor.state.schema.nodes.hardBreak
    const tr = editor.state.tr.insert(inside(editor, 0, 3), br.create())
    expect(singleInlineEditIndex(tr)).toBeNull()
  })

  it('rejects attribute changes and decoration-only transactions', () => {
    const editor = createEditor([para('first'), para('second')])
    const attrs = editor.state.tr.setNodeMarkup(inside(editor, 1, 0), undefined, {
      docxIndex: null,
      align: 'center',
    })
    expect(singleInlineEditIndex(attrs)).toBeNull()
    const sel = editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 2))
    expect(singleInlineEditIndex(sel)).toBeNull()
  })
})

describe('crossesPage', () => {
  const slices: PageSlice[] = [
    { start: 0, end: 900, section: 0 },
    { start: 900, end: 1800, section: 0 },
  ]
  it('is false for a block inside one page and for a page-leading block', () => {
    expect(crossesPage({ top: 100, height: 200 }, slices)).toBe(false)
    expect(crossesPage({ top: 900, height: 50 }, slices)).toBe(false)
    expect(crossesPage({ top: 850, height: 50 }, slices)).toBe(false)
  })
  it('is true when a boundary falls inside the block', () => {
    expect(crossesPage({ top: 850, height: 100 }, slices)).toBe(true)
  })
})

describe('blocksKeepSlicing', () => {
  const slices: PageSlice[] = [{ start: 0, end: 900, section: 0 }]
  const el = (h: number, html = 'text'): HTMLElement => {
    const p = document.createElement('p')
    p.innerHTML = html
    p.getBoundingClientRect = () =>
      ({ height: h, width: 500, top: 0, left: 0, bottom: h, right: 500 }) as DOMRect
    return p
  }
  const box = (domHeight: number): BlockBox => ({ top: 100, height: domHeight + 8, domHeight })

  it('keeps the slicing when the edited block has the same height', () => {
    expect(blocksKeepSlicing([{ el: el(40), prev: box(40) }], slices, 1)).toBe(true)
    expect(blocksKeepSlicing([{ el: el(80), prev: box(40) }], slices, 2)).toBe(true)
  })
  it('rejects a paragraph turning empty or non-empty (column-balance quota)', () => {
    expect(
      blocksKeepSlicing([{ el: el(40), prev: { ...box(40), emptyPara: true } }], slices, 1),
    ).toBe(false)
    expect(blocksKeepSlicing([{ el: el(40, ''), prev: box(40) }], slices, 1)).toBe(false)
    expect(
      blocksKeepSlicing([{ el: el(40, ''), prev: { ...box(40), emptyPara: true } }], slices, 1),
    ).toBe(true)
  })
  it('rejects a height change, geometry-bearing content and unmeasured blocks', () => {
    expect(blocksKeepSlicing([{ el: el(58), prev: box(40) }], slices, 1)).toBe(false)
    expect(blocksKeepSlicing([{ el: el(40, 'a<img>'), prev: box(40) }], slices, 1)).toBe(false)
    expect(blocksKeepSlicing([{ el: el(40), prev: { top: 100, height: 48 } }], slices, 1)).toBe(
      false,
    )
    expect(blocksKeepSlicing([], slices, 1)).toBe(false)
  })
})
