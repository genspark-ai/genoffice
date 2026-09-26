/**
 * Review > Comments: Previous / Next walk the open threads in document order
 * and light the thread up in the pane; Delete ▾ removes the thread under the
 * caret, every resolved thread, or all of them in one document pass.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import type { CommentInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  commentAnchors,
  commentIdsAt,
  removeCommentsFromDoc,
} from '../src/renderer/editor/comments'
import {
  commentThreadAtCaret,
  deleteAllComments,
  deleteCommentAtCaret,
  gotoComment,
  type ReviewContext,
} from '../src/renderer/review-actions'

const live: Editor[] = []
afterEach(() => {
  for (const e of live.splice(0)) e.destroy()
})

const marked = (text: string, ids: string) => ({
  type: 'text',
  text,
  marks: [{ type: 'comment', attrs: { ids } }],
})

function makeEditor() {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'alpha ' },
            marked('beta', '1 4'),
            { type: 'text', text: ' gamma ' },
            marked('delta', '2'),
          ],
        },
        {
          type: 'docParagraph',
          content: [{ type: 'text', text: 'eps ' }, marked('zeta', '3')],
        },
      ],
    },
  })
  live.push(editor)
  return editor
}

const info = (id: string, extra: Partial<CommentInfo> = {}): CommentInfo => ({
  id,
  author: 'a',
  text: `c${id}`,
  ...extra,
})

function makeCtx(editor: Editor, comments: CommentInfo[]) {
  const state = { comments }
  const ctx = {
    editor,
    dirtyRef: { current: false },
    setStatus: vi.fn(),
    get comments() {
      return state.comments
    },
    setComments: vi.fn((update: (prev: CommentInfo[]) => CommentInfo[]) => {
      state.comments = update(state.comments)
    }),
    setCommentsDirty: vi.fn(),
    setShowComments: vi.fn(),
    setCommentFocus: vi.fn(),
  } as unknown as ReviewContext
  return { ctx, state }
}

describe('comment anchors', () => {
  it('lists the first range of every id in document order', () => {
    const editor = makeEditor()
    // "alpha " 1..7, beta 7..11, " gamma " 11..18, delta 18..23; p2 starts at 24, "eps " 25..29, zeta 29..33
    expect(commentAnchors(editor.state.doc)).toEqual([
      { id: '1', from: 7, to: 11 },
      { id: '4', from: 7, to: 11 },
      { id: '2', from: 18, to: 23 },
      { id: '3', from: 29, to: 33 },
    ])
    expect(commentIdsAt(editor.state, 9)).toEqual(['1', '4'])
    expect(commentIdsAt(editor.state, 11)).toEqual(['1', '4'])
    expect(commentIdsAt(editor.state, 14)).toEqual([])
  })

  it('removes a batch of ids in one pass', () => {
    const editor = makeEditor()
    removeCommentsFromDoc(editor, ['1', '3'])
    expect(commentAnchors(editor.state.doc)).toEqual([
      { id: '4', from: 7, to: 11 },
      { id: '2', from: 18, to: 23 },
    ])
    expect(editor.state.doc.textContent).toBe('alpha beta gamma deltaeps zeta')
  })
})

describe('Previous / Next comment', () => {
  it('cycles through open threads, selects the anchor and focuses the pane', () => {
    const editor = makeEditor()
    const { ctx } = makeCtx(editor, [
      info('1'),
      info('4', { parentId: '1' }),
      info('2', { done: true }),
      info('3'),
    ])
    editor.commands.setTextSelection(1)
    expect(gotoComment(ctx, 1)).toBe(true)
    expect([editor.state.selection.from, editor.state.selection.to]).toEqual([7, 11])
    expect(ctx.setShowComments).toHaveBeenCalledWith(true)
    expect(ctx.setCommentFocus).toHaveBeenLastCalledWith(expect.objectContaining({ id: '1' }))
    // the resolved thread (2) is skipped, replies never count as stops
    gotoComment(ctx, 1)
    expect([editor.state.selection.from, editor.state.selection.to]).toEqual([29, 33])
    expect(ctx.setCommentFocus).toHaveBeenLastCalledWith(expect.objectContaining({ id: '3' }))
    gotoComment(ctx, 1)
    expect(editor.state.selection.from).toBe(7)
    gotoComment(ctx, -1)
    expect(editor.state.selection.from).toBe(29)
    // from inside a thread's anchor, Previous leaves it (wraps to the last one here)
    editor.commands.setTextSelection(9)
    gotoComment(ctx, -1)
    expect(editor.state.selection.from).toBe(29)
    editor.commands.setTextSelection(31)
    gotoComment(ctx, -1)
    expect(editor.state.selection.from).toBe(7)
  })

  it('leaves a zero-width block anchor in both directions', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'one ' }, marked('two', '1')] },
          {
            type: 'docParagraph',
            attrs: { commentStarts: ['2'] },
            content: [{ type: 'text', text: 'x' }],
          },
        ],
      },
    })
    live.push(editor)
    const { ctx } = makeCtx(editor, [info('1'), info('2')])
    expect(commentAnchors(editor.state.doc).map((a) => [a.id, a.from, a.to])).toEqual([
      ['1', 5, 8],
      ['2', 10, 10],
    ])
    editor.commands.setTextSelection(1)
    gotoComment(ctx, 1)
    gotoComment(ctx, 1)
    expect(ctx.setCommentFocus).toHaveBeenLastCalledWith(expect.objectContaining({ id: '2' }))
    gotoComment(ctx, -1)
    expect(ctx.setCommentFocus).toHaveBeenLastCalledWith(expect.objectContaining({ id: '1' }))
    gotoComment(ctx, 1)
    expect(ctx.setCommentFocus).toHaveBeenLastCalledWith(expect.objectContaining({ id: '2' }))
  })

  it('does nothing without open threads', () => {
    const editor = makeEditor()
    const { ctx } = makeCtx(editor, [info('1', { done: true })])
    expect(gotoComment(ctx, 1)).toBe(false)
    expect(ctx.setShowComments).not.toHaveBeenCalled()
  })
})

describe('Delete ▾', () => {
  it('deletes the thread under the caret, resolving a reply to its parent', () => {
    const editor = makeEditor()
    const { ctx, state } = makeCtx(editor, [info('1'), info('4', { parentId: '1' }), info('2')])
    editor.commands.setTextSelection(9)
    expect(commentThreadAtCaret(ctx)).toBe('1')
    deleteCommentAtCaret(ctx)
    expect(state.comments.map((c) => c.id)).toEqual(['2'])
    expect(commentAnchors(editor.state.doc).map((a) => a.id)).toEqual(['2', '3'])
    editor.commands.setTextSelection(3)
    expect(commentThreadAtCaret(ctx)).toBeNull()
  })

  it('deletes all resolved threads, then everything', () => {
    const editor = makeEditor()
    const { ctx, state } = makeCtx(editor, [
      info('1', { done: true }),
      info('4', { parentId: '1' }),
      info('2'),
      info('3'),
    ])
    // the reply of a resolved thread goes with it even though it carries no done flag
    deleteAllComments(ctx, true)
    expect(state.comments.map((c) => c.id)).toEqual(['2', '3'])
    expect(commentAnchors(editor.state.doc).map((a) => a.id)).toEqual(['2', '3'])
    deleteAllComments(ctx, false)
    expect(state.comments).toEqual([])
    expect(commentAnchors(editor.state.doc)).toEqual([])
    expect(ctx.dirtyRef.current).toBe(true)
  })
})
