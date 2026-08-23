import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { CommentInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  addCommentToSelection,
  addReplyToCommentRange,
  nextCommentId,
} from '../src/renderer/editor/comments'
import { buildCommentsContext, buildDocContext, commentAnchors } from '../src/renderer/ai/protocol'
import { executeTool, type AiCommentsAccess } from '../src/renderer/ai/tools'
import { replyToComment, type ReviewContext } from '../src/renderer/review-actions'

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

const editors = new Set<Editor>()
afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

function createEditor(content: JsonNode[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.add(editor)
  return editor
}

const NUM_IDS = { bullet: null, ordered: null }
const DATE = '2026-01-01T00:00:00Z'

/** editor with a comment (id 1) anchored inside block 1 */
function setup(): { editor: Editor; comments: CommentInfo[] } {
  const editor = createEditor([
    para('Intro paragraph.'),
    para('The commented paragraph with enough text.'),
    para('Closing paragraph.'),
  ])
  const block0 = editor.state.doc.child(0).nodeSize
  editor.commands.setTextSelection({ from: block0 + 2, to: block0 + 12 })
  addCommentToSelection(editor, '1')
  const comments: CommentInfo[] = [
    { id: '1', author: 'User', date: DATE, text: 'Please shorten this' },
  ]
  return { editor, comments }
}

function makeAccess(editor: Editor, comments: CommentInfo[]): AiCommentsAccess {
  return {
    list: () => comments,
    reply: (parentId, text) => {
      const id = nextCommentId(comments)
      if (!addReplyToCommentRange(editor, parentId, id)) return false
      comments.push({ id, author: 'AI Assistant', date: DATE, text, parentId })
      return true
    },
    resolve: (id) => {
      if (!comments.some((c) => c.id === id)) return false
      for (let i = 0; i < comments.length; i++) {
        const c = comments[i]
        if (c.id === id || c.parentId === id) comments[i] = { ...c, done: true }
      }
      return true
    },
  }
}

describe('comments context', () => {
  it('anchors carry the block index and the anchored text', () => {
    const { editor } = setup()
    const anchors = commentAnchors(editor)
    expect(anchors.get('1')).toEqual({ blockIndex: 1, excerpt: 'he comment' })
  })

  it('unresolved threads ride along in the per-turn context, resolved ones drop out', () => {
    const { editor, comments } = setup()
    const context = buildDocContext(editor, undefined, comments)
    expect(context).toContain('Unresolved comments')
    expect(context).toContain('id 1 by User (block 1')
    expect(context).toContain('Please shorten this')
    const resolvedContext = buildDocContext(editor, undefined, [{ ...comments[0], done: true }])
    expect(resolvedContext).not.toContain('Unresolved comments')
  })

  it('the read_comments variant lists resolved threads and replies', () => {
    const { editor, comments } = setup()
    comments[0] = { ...comments[0], done: true }
    comments.push({ id: '2', author: 'AI Assistant', date: DATE, text: 'Done.', parentId: '1' })
    const full = buildCommentsContext(editor, comments, true)
    expect(full).toContain('[resolved]')
    expect(full).toContain('reply id 2 by AI Assistant: Done.')
  })
})

describe('comment tools', () => {
  it('reply_comment attaches to the thread root even when given a reply id', async () => {
    const { editor, comments } = setup()
    const access = makeAccess(editor, comments)
    access.reply('1', 'first reply')
    const exec = await executeTool(
      editor,
      { id: 't', name: 'reply_comment', input: { parentId: '2', text: 'second reply' } },
      NUM_IDS,
      undefined,
      undefined,
      undefined,
      access,
    )
    expect(exec.isError).toBeFalsy()
    expect(exec.mutated).toBe(true)
    const last = comments.at(-1)!
    expect(last.parentId).toBe('1')
    expect(last.text).toBe('second reply')
  })

  it('resolve_comment resolves the whole thread via the root', async () => {
    const { editor, comments } = setup()
    const access = makeAccess(editor, comments)
    const exec = await executeTool(
      editor,
      { id: 't', name: 'resolve_comment', input: { id: '1' } },
      NUM_IDS,
      undefined,
      undefined,
      undefined,
      access,
    )
    expect(exec.isError).toBeFalsy()
    expect(exec.mutated).toBe(false)
    expect(comments[0].done).toBe(true)
  })

  it('back-to-back replies in one tick mint distinct ids (App wiring: getter + array-replacing setter)', () => {
    const { editor, comments } = setup()
    // mirrors App's live-mirror wiring exactly: setComments REPLACES the array
    // (a functional React update), and ctx.comments is a getter into the live
    // reference — a render-time snapshot here would mint duplicate ids
    let live = comments
    const ctx = {
      editor,
      get comments() {
        return live
      },
      setComments: (action: unknown) => {
        live =
          typeof action === 'function'
            ? (action as (prev: CommentInfo[]) => CommentInfo[])(live)
            : (action as CommentInfo[])
      },
      setCommentsDirty: () => {},
      setStatus: () => {},
      dirtyRef: { current: false },
    } as unknown as ReviewContext
    expect(replyToComment(ctx, '1', 'first', 'AI Assistant')).toBe(true)
    expect(replyToComment(ctx, '1', 'second', 'AI Assistant')).toBe(true)
    const ids = live.map((c) => c.id)
    expect(new Set(ids).size).toBe(ids.length)
    expect(live.filter((c) => c.parentId === '1')).toHaveLength(2)
  })

  it('unknown ids and a missing access fail with guidance', async () => {
    const { editor, comments } = setup()
    const access = makeAccess(editor, comments)
    const unknown = await executeTool(
      editor,
      { id: 't', name: 'resolve_comment', input: { id: '99' } },
      NUM_IDS,
      undefined,
      undefined,
      undefined,
      access,
    )
    expect(unknown.isError).toBe(true)
    expect(unknown.output).toContain('read_comments')
    const noAccess = await executeTool(
      editor,
      {
        id: 't',
        name: 'read_comments',
        input: {},
      } as never,
      NUM_IDS,
    )
    expect(noAccess.isError).toBe(true)
  })
})
