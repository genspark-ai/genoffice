import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions, pendingCommentPluginKey } from '../src/renderer/editor/extensions'
import {
  addCommentToSelection,
  addReplyToCommentRange,
  canCommentAt,
  commentAnchors,
  commentIdsAt,
  nextCommentId,
  removeCommentFromDoc,
} from '../src/renderer/editor/comments'
import { pmDocToSavePlan } from '../src/renderer/editor/convert'

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

describe('a comment on an empty paragraph', () => {
  const emptyBetween = () => createEditor([para(text('above')), para(), para(text('below'))])
  const rangesOf = (editor: Editor, index: number) => {
    const { commentStarts, commentEnds } = editor.state.doc.child(index).attrs
    return [commentStarts, commentEnds]
  }

  it('is held by the paragraph, opening and closing around nothing', () => {
    const editor = emptyBetween()
    select(editor, 8, 8)
    expect(canCommentAt(editor)).toBe(true)
    expect(addCommentToSelection(editor, '1')).toBe(true)
    expect(rangesOf(editor, 1)).toEqual([['1'], ['1']])
    expect(commentIdsPerNode(editor)).toEqual([null, null])
    expect(commentAnchors(editor.state.doc)).toEqual([{ id: '1', from: 8, to: 8 }])
    expect(commentIdsAt(editor.state, 8)).toEqual(['1'])
    expect(commentIdsAt(editor.state, 3)).toEqual([])
    editor.destroy()
  })

  it('takes a reply where its parent is, and goes with one delete', () => {
    const editor = emptyBetween()
    select(editor, 8, 8)
    addCommentToSelection(editor, '1')
    expect(addReplyToCommentRange(editor, '1', '2')).toBe(true)
    expect(rangesOf(editor, 1)).toEqual([
      ['1', '2'],
      ['1', '2'],
    ])
    removeCommentFromDoc(editor, '2')
    removeCommentFromDoc(editor, '1')
    expect(rangesOf(editor, 1)).toEqual([null, null])
    editor.destroy()
  })

  it('is written once when the paragraph was split before the first save', () => {
    const editor = emptyBetween()
    select(editor, 8, 8)
    addCommentToSelection(editor, '1')
    editor.commands.splitBlock()
    const plan = pmDocToSavePlan(editor.getJSON() as never, [])
    const ranged = plan.saveBlocks.filter(
      (b) => b.kind === 'generated' && (b.block.commentStarts || b.block.commentEnds),
    )
    expect(ranged).toHaveLength(1)
    editor.destroy()
  })

  it('made on the empty line a split left under a saved paragraph, it is written with that line', () => {
    const original = {
      type: 'paragraph',
      docxIndex: 0,
      originalXml: '<w:p><w:r><w:t>saved</w:t></w:r></w:p>',
      commentStarts: ['7'],
      runs: [{ text: 'saved' }],
    }
    const editor = createEditor([
      paraWithAttrs({ docxIndex: 0, commentStarts: ['7'] }, text('saved')),
    ])
    select(editor, 6, 6)
    editor.commands.splitBlock()
    expect(editor.state.doc.child(1).content.size).toBe(0)
    expect(addCommentToSelection(editor, '8')).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as never, [original] as never)
    const ranges = plan.saveBlocks.map((b) =>
      b.kind === 'generated' ? [b.block.commentStarts, b.block.commentEnds] : b.kind,
    )
    // the saved paragraph is written as it was, with the range it opened; the new one goes with the new line
    expect(ranges).toEqual(['original', [['8'], ['8']]])
    editor.destroy()
  })

  it('takes the number of a deleted comment the saved paragraph used to hold', () => {
    const original = {
      type: 'paragraph',
      docxIndex: 0,
      originalXml: '<w:p><w:r><w:t>saved</w:t></w:r></w:p>',
      commentStarts: ['7'],
      commentEnds: ['7'],
      runs: [{ text: 'saved' }],
    }
    const editor = createEditor([
      paraWithAttrs({ docxIndex: 0, commentStarts: ['7'], commentEnds: ['7'] }, text('saved')),
    ])
    removeCommentFromDoc(editor, '7')
    select(editor, 6, 6)
    editor.commands.splitBlock()
    expect(addCommentToSelection(editor, '7')).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as never, [original] as never)
    const ranges = plan.saveBlocks.map((b) =>
      b.kind === 'generated' ? [b.block.commentStarts, b.block.commentEnds] : b.kind,
    )
    expect(ranges).toEqual([
      [undefined, undefined],
      [['7'], ['7']],
    ])
    editor.destroy()
  })

  it('has nowhere to go at a caret beside no word, or in a table cell', () => {
    const editor = createEditor([para(text('... ...'))])
    select(editor, 5, 5)
    expect(canCommentAt(editor)).toBe(false)
    expect(addCommentToSelection(editor, '1')).toBe(false)
    editor.destroy()
    const table = createEditor([
      {
        type: 'docTable',
        content: [
          {
            type: 'docTableRow',
            content: [{ type: 'docTableCell', content: [{ type: 'docParagraph' }] }],
          },
        ],
      },
    ])
    table.commands.focus('start')
    expect(table.state.selection.$from.parent.content.size).toBe(0)
    expect(canCommentAt(table)).toBe(false)
    table.destroy()
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
