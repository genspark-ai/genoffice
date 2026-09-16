import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { CommentInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  addCommentToRange,
  addCommentToSelection,
  addReplyToCommentRange,
  nextCommentId,
  removeCommentFromDoc,
} from '../src/renderer/editor/comments'
import { buildCommentsContext, buildDocContext, commentAnchors } from '../src/renderer/ai/protocol'
import { executeTool, type AiCommentsAccess } from '../src/renderer/ai/tools'
import { resolveCommentAnchor } from '../src/renderer/ai/comment-ops'
import { replyToComment, type ReviewContext } from '../src/renderer/review-actions'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: { type: string; attrs?: Record<string, unknown> }[]
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
    add: (range, text, meta) => {
      const id = nextCommentId(comments)
      if (!addCommentToRange(editor, range.from, range.to, id)) return null
      comments.push({ id, author: meta.author ?? 'AI Assistant', date: DATE, text })
      return id
    },
    remove: (id) => {
      if (!comments.some((c) => c.id === id)) return false
      const victims = [id, ...comments.filter((c) => c.parentId === id).map((c) => c.id)]
      for (const v of victims) removeCommentFromDoc(editor, v)
      for (let i = comments.length - 1; i >= 0; i--) {
        if (victims.includes(comments[i].id)) comments.splice(i, 1)
      }
      return true
    },
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

const call = (
  editor: Editor,
  access: AiCommentsAccess,
  name: string,
  input: Record<string, unknown>,
) => executeTool(editor, { id: 't', name, input }, NUM_IDS, undefined, undefined, undefined, access)

describe('add_comment / delete_comment', () => {
  it('anchors a new thread to an exact text span and reports it in the comments context', async () => {
    const { editor, comments } = setup()
    const access = makeAccess(editor, comments)
    const exec = await call(editor, access, 'add_comment', {
      blockIndex: 2,
      text: 'Closing',
      comment: 'Consider a stronger ending.',
    })
    expect(exec.isError).toBeFalsy()
    expect(exec.mutated).toBe(true)
    expect(exec.output).toContain('Added comment 2 on block 2')
    expect(commentAnchors(editor).get('2')).toEqual({ blockIndex: 2, excerpt: 'Closing' })
    expect(comments.at(-1)).toMatchObject({
      id: '2',
      author: 'AI Assistant',
      text: 'Consider a stronger ending.',
    })
    expect(buildCommentsContext(editor, comments, true)).toContain(
      'id 2 by AI Assistant (block 2, anchored text "Closing")',
    )
  })

  it('covers the whole block when text is omitted and honours author/initials', async () => {
    const { editor, comments } = setup()
    const access = makeAccess(editor, comments)
    const exec = await call(editor, access, 'add_comment', {
      blockIndex: 0,
      comment: 'Whole block.',
      author: 'Reviewer',
      initials: 'RV',
    })
    expect(exec.isError).toBeFalsy()
    expect(commentAnchors(editor).get('2')?.excerpt).toBe('Intro paragraph.')
    expect(comments.at(-1)?.author).toBe('Reviewer')
  })

  it('refuses ambiguous, missing and out-of-range anchors with guidance', () => {
    const editor = createEditor([para('again and again and again')])
    expect(resolveCommentAnchor(editor, { blockIndex: 0, text: 'again' })).toEqual({
      error: '"again" occurs 3 times in block 0; set occurrence (1-3) to pick one',
    })
    const third = resolveCommentAnchor(editor, { blockIndex: 0, text: 'again', occurrence: 3 })
    expect(third).toMatchObject({ from: 1 + 'again and again and '.length, excerpt: 'again' })
    expect(resolveCommentAnchor(editor, { blockIndex: 0, text: 'nope' })).toMatchObject({
      error: expect.stringContaining('does not contain "nope"'),
    })
    expect(resolveCommentAnchor(editor, { blockIndex: 4 })).toMatchObject({
      error: expect.stringContaining('between 0 and 0'),
    })
    expect(
      resolveCommentAnchor(editor, { blockIndex: 0, text: 'again', occurrence: 9 }),
    ).toMatchObject({
      error: expect.stringContaining('out of range'),
    })
  })

  it('matches against the live text: tracked deletions are skipped, positions stay exact', () => {
    const REV = { author: 'Bob', date: DATE }
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: null },
        content: [
          text('Keep '),
          { type: 'text', text: 'gone target ', marks: [{ type: 'del', attrs: REV }] },
          text('target here'),
        ],
      },
    ])
    const hit = resolveCommentAnchor(editor, { blockIndex: 0, text: 'target' })
    const from = 1 + 'Keep gone target '.length
    expect(hit).toEqual({ from, to: from + 'target'.length, excerpt: 'target' })
    expect(resolveCommentAnchor(editor, { blockIndex: 0, text: 'gone' })).toMatchObject({
      error: expect.stringContaining('does not contain'),
    })
    expect(resolveCommentAnchor(editor, { blockIndex: 0, text: 'Keep target' })).toMatchObject({
      from: 1,
      to: from + 'target'.length,
    })
    expect(resolveCommentAnchor(editor, { blockIndex: 0 })).toMatchObject({
      excerpt: 'Keep target here',
    })
  })

  it('a table anchors to its first paragraph only; blocks without text fail the same way either way', () => {
    const cell = (label: string) => ({
      type: 'docTableCell',
      content: [{ type: 'docParagraph', attrs: { docxIndex: null }, content: [text(label)] }],
    })
    const editor = createEditor([
      {
        type: 'docTable',
        attrs: { docxIndex: null },
        content: [
          { type: 'docTableRow', content: [cell('City'), cell('GDP')] },
          { type: 'docTableRow', content: [cell('Paris'), cell('900')] },
        ],
      },
      { type: 'docParagraph', attrs: { docxIndex: null }, content: [] },
    ])
    const table = resolveCommentAnchor(editor, { blockIndex: 0 })
    // table(1) + row(1) + cell(1) + paragraph(1) = first cell text starts at 4
    expect(table).toEqual({ from: 4, to: 4 + 'City'.length, excerpt: 'City' })
    expect(resolveCommentAnchor(editor, { blockIndex: 0, text: 'GDP' })).toMatchObject({
      error: expect.stringContaining('omit text to anchor the comment to its first paragraph'),
    })
    const empty =
      'block 1 (docParagraph) has no text to anchor a comment to; pick a block with text'
    expect(resolveCommentAnchor(editor, { blockIndex: 1 })).toEqual({ error: empty })
    expect(resolveCommentAnchor(editor, { blockIndex: 1, text: 'x' })).toEqual({ error: empty })
  })

  it('delete_comment refuses a root with replies unless withReplies, then cascades', async () => {
    const { editor, comments } = setup()
    const access = makeAccess(editor, comments)
    access.reply('1', 'a reply')
    const refused = await call(editor, access, 'delete_comment', { id: '1' })
    expect(refused.isError).toBe(true)
    expect(refused.output).toContain('withReplies')
    expect(comments).toHaveLength(2)

    const replyOnly = await call(editor, access, 'delete_comment', { id: '2' })
    expect(replyOnly.isError).toBeFalsy()
    expect(comments.map((c) => c.id)).toEqual(['1'])
    expect(commentAnchors(editor).has('2')).toBe(false)
    expect(commentAnchors(editor).has('1')).toBe(true)

    access.reply('1', 'another')
    const cascade = await call(editor, access, 'delete_comment', { id: '1', withReplies: true })
    expect(cascade.isError).toBeFalsy()
    expect(cascade.output).toBe('Deleted comments 1, 2.')
    expect(comments).toHaveLength(0)
    expect(commentAnchors(editor).size).toBe(0)

    const unknown = await call(editor, access, 'delete_comment', { id: '7' })
    expect(unknown.isError).toBe(true)
    expect(unknown.output).toContain('read_comments')
  })
})
