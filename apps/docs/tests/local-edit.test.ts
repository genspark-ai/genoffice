import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  containsNode,
  localEditBlocks,
  touchedNeedsRecompute,
} from '../src/renderer/editor/local-edit'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
}
const text = (t: string): JsonNode => ({ type: 'text', text: t })
const para = (t: string, attrs: Record<string, unknown> = {}): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, ...attrs },
  content: [text(t)],
})
const BORDER = { borders: 'top,bottom', borderLines: null, shadingFill: null }

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
function inside(editor: Editor, index: number, offset = 1): number {
  let pos = 0
  for (let i = 0; i < index; i++) pos += editor.state.doc.child(i).nodeSize
  return pos + offset
}
const suppressed = (editor: Editor) =>
  Array.from(editor.view.dom.querySelectorAll('p')).map((p) =>
    [...p.classList]
      .filter((c) => c.startsWith('pbdr-suppress'))
      .sort()
      .join(' '),
  )

describe('localEditBlocks / touchedNeedsRecompute', () => {
  it('reports the touched top-level block of an inline edit and rejects block-count changes', () => {
    const editor = createEditor([para('one'), para('two'), para('three')])
    const tr = editor.state.tr.insertText('x', inside(editor, 1, 2))
    expect([...(localEditBlocks(tr) ?? [])]).toEqual([inside(editor, 1, 0)])
    const split = editor.state.tr.split(inside(editor, 1, 2))
    expect(localEditBlocks(split)).toBeNull()
  })
  it('asks for a rebuild when the touched block matters or carried a decoration', () => {
    const editor = createEditor([para('one'), para('two', BORDER), para('three')])
    const tr = editor.state.tr.insertText('x', inside(editor, 2, 2))
    const touched = localEditBlocks(tr)!
    const empty = DecorationSet.empty
    expect(touchedNeedsRecompute(tr, touched, empty, () => false)).toBe(false)
    expect(touchedNeedsRecompute(tr, touched, empty, (n) => !!n.attrs.borders)).toBe(false)
    const third = tr.doc.childAfter(inside(editor, 2, 0))
    const decorated = DecorationSet.create(tr.doc, [
      Decoration.node(third.offset, third.offset + third.node!.nodeSize, { class: 'x' }),
    ])
    expect(touchedNeedsRecompute(tr, touched, decorated, () => false)).toBe(true)
    // what the block stopped being counts too: dropping the borders of an
    // undecorated bordered block must rebuild (its neighbours may regroup)
    const trUnborder = editor.state.tr.setNodeMarkup(inside(editor, 1, 0), undefined, {
      ...editor.state.doc.child(1).attrs,
      borders: null,
    })
    expect(
      touchedNeedsRecompute(
        trUnborder,
        localEditBlocks(trUnborder)!,
        empty,
        (n) => !!n.attrs.borders,
      ),
    ).toBe(true)
    const trBordered = editor.state.tr.insertText('x', inside(editor, 1, 2))
    expect(
      touchedNeedsRecompute(
        trBordered,
        localEditBlocks(trBordered)!,
        empty,
        (n) => !!n.attrs.borders,
      ),
    ).toBe(true)
  })
})

describe('border groups under local edits', () => {
  it('keeps the group classes while typing elsewhere and regroups when borders change', () => {
    const editor = createEditor([
      para('a', BORDER),
      para('b', BORDER),
      para('plain'),
      para('c', BORDER),
    ])
    expect(suppressed(editor)).toEqual(['pbdr-suppress-bottom', 'pbdr-suppress-top', '', ''])
    editor.view.dispatch(editor.state.tr.insertText('!', inside(editor, 2, 3)))
    expect(suppressed(editor)).toEqual(['pbdr-suppress-bottom', 'pbdr-suppress-top', '', ''])
    // the plain paragraph joins the group: a, b, plain, c become one run
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(inside(editor, 2, 0), undefined, {
        ...editor.state.doc.child(2).attrs,
        ...BORDER,
      }),
    )
    expect(suppressed(editor)).toEqual([
      'pbdr-suppress-bottom',
      'pbdr-suppress-bottom pbdr-suppress-top',
      'pbdr-suppress-bottom pbdr-suppress-top',
      'pbdr-suppress-top',
    ])
    // typing into a grouped paragraph keeps its classes
    editor.view.dispatch(editor.state.tr.insertText('!', inside(editor, 1, 2)))
    expect(suppressed(editor)[1]).toBe('pbdr-suppress-bottom pbdr-suppress-top')
  })
})

describe('containsNode', () => {
  it('sees list items nested inside a table cell', () => {
    const editor = createEditor([para('a')])
    editor.commands.setContent({
      type: 'doc',
      content: [
        para('a'),
        {
          type: 'docTable',
          content: [
            {
              type: 'docTableRow',
              content: [
                {
                  type: 'docTableCell',
                  content: [
                    {
                      type: 'docListItem',
                      attrs: { docxIndex: null, numId: '1', ilvl: 0 },
                      content: [text('item')],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    })
    const table = editor.state.doc.child(1)
    expect(table.type.name).toBe('docTable')
    expect(containsNode(table, (n) => n.type.name === 'docListItem')).toBe(true)
    expect(containsNode(editor.state.doc.child(0), (n) => n.type.name === 'docListItem')).toBe(
      false,
    )
  })
})
