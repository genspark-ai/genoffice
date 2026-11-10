import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions, pendingCommentPluginKey } from '../src/renderer/editor/extensions'
import {
  addCommentToSelection,
  nextCommentId,
  removeCommentFromDoc,
} from '../src/renderer/editor/comments'

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
const para = (...content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
})
const paraWithAttrs = (attrs: Record<string, unknown>, ...content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, ...attrs },
  content,
})

function createEditor(content: JsonNode[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
}

function select(editor: Editor, from: number, to: number) {
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)),
  )
}

/** ids attr of the comment mark covering each text node, in doc order */
function commentIdsPerNode(editor: Editor): Array<string | null> {
  const out: Array<string | null> = []
  editor.state.doc.descendants((node) => {
    if (!node.isText) return
    const mark = node.marks.find((m) => m.type.name === 'comment')
    out.push(mark ? String(mark.attrs.ids) : null)
  })
  return out
}

describe('nextCommentId', () => {
  it('allocates one past the max numeric id', () => {
    expect(nextCommentId([])).toBe('1')
    expect(
      nextCommentId([
        { id: '3', author: 'a', text: '' },
        { id: '7', author: 'b', text: '' },
      ]),
    ).toBe('8')
  })
})

describe('addCommentToSelection', () => {
  it('marks the selected text with the comment id', () => {
    const editor = createEditor([para(text('abcd'))])
    select(editor, 2, 4)
    expect(addCommentToSelection(editor, '1')).toBe(true)
    expect(commentIdsPerNode(editor)).toEqual([null, '1', null])
    editor.destroy()
  })

  it('merges with an existing overlapping comment', () => {
    const editor = createEditor([
      para(text('ab'), text('cd', [{ type: 'comment', attrs: { ids: '1' } }])),
    ])
    select(editor, 1, 5) // covers ab + cd
    addCommentToSelection(editor, '2')
    expect(commentIdsPerNode(editor)).toEqual(['2', '1 2'])
    editor.destroy()
  })

  it('returns false for an empty selection', () => {
    const editor = createEditor([para(text('abc'))])
    select(editor, 2, 2)
    expect(addCommentToSelection(editor, '9')).toBe(false)
    editor.destroy()
  })
})

describe('removeCommentFromDoc', () => {
  it('removes the mark when the id was the last one', () => {
    const editor = createEditor([
      para(text('x'), text('marked', [{ type: 'comment', attrs: { ids: '1' } }])),
    ])
    removeCommentFromDoc(editor, '1')
    expect(commentIdsPerNode(editor)).toEqual([null])
    editor.destroy()
  })

  it('keeps other ids on shared ranges', () => {
    const editor = createEditor([
      para(text('shared', [{ type: 'comment', attrs: { ids: '1 2' } }])),
    ])
    removeCommentFromDoc(editor, '1')
    expect(commentIdsPerNode(editor)).toEqual(['2'])
    editor.destroy()
  })

  it('clears cross-paragraph endpoint attrs while retaining other comment ids', () => {
    const editor = createEditor([
      paraWithAttrs({ commentStarts: ['1', '2'] }, text('range starts')),
      paraWithAttrs({ commentEnds: ['1'] }, text('range ends')),
    ])
    removeCommentFromDoc(editor, '1')
    expect(editor.state.doc.child(0).attrs.commentStarts).toEqual(['2'])
    expect(editor.state.doc.child(0).attrs.commentEnds).toBeNull()
    expect(editor.state.doc.child(1).attrs.commentStarts).toBeNull()
    expect(editor.state.doc.child(1).attrs.commentEnds).toBeNull()
    editor.destroy()
  })
})

describe('pending comment highlight', () => {
  it('decorates the pushed range and clears on null', () => {
    const editor = createEditor([para(text('pending text'))])
    const decosAt = () => {
      const set = pendingCommentPluginKey.getState(editor.state)
      return set ? set.find() : []
    }

    editor.view.dispatch(editor.state.tr.setMeta(pendingCommentPluginKey, { from: 2, to: 5 }))
    expect(decosAt()).toHaveLength(1)
    expect(decosAt()[0].from).toBe(2)
    expect(decosAt()[0].to).toBe(5)

    editor.view.dispatch(editor.state.tr.setMeta(pendingCommentPluginKey, null))
    expect(decosAt()).toHaveLength(0)
    editor.destroy()
  })

  it('survives unrelated transactions by mapping positions', () => {
    const editor = createEditor([para(text('abcdef'))])
    editor.view.dispatch(editor.state.tr.setMeta(pendingCommentPluginKey, { from: 3, to: 6 }))
    editor.view.dispatch(editor.state.tr.insertText('XX', 1, 1)) // insert before the range
    const set = pendingCommentPluginKey.getState(editor.state)
    const decos = set ? set.find() : []
    expect(decos).toHaveLength(1)
    expect(decos[0].from).toBe(5)
    expect(decos[0].to).toBe(8)
    editor.destroy()
  })
})
