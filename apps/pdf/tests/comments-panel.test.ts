/**
 * Pure logic of the all-document comments panel: the text filter (matches
 * thread roots, replies and authors) and the page/time ordering.
 */
import { describe, expect, it } from 'vitest'
import {
  countReplies,
  filterCommentEntries,
  type CommentEntry,
} from '../src/renderer/CommentsPanel'
import type { NoteThreadItem } from '../src/renderer/note-threads'

const note = (over: Partial<NoteThreadItem> = {}): NoteThreadItem => ({
  key: 'S1',
  saved: null,
  pendingId: null,
  author: 'Ann',
  contents: '',
  timeMs: null,
  color: null,
  at: [0, 0],
  replies: [],
  ...over,
})

const entry = (pageNo: number, root: NoteThreadItem, origIdx = pageNo - 1): CommentEntry => ({
  origIdx,
  pageNo,
  root,
})

describe('filterCommentEntries', () => {
  it('matches the root contents, reply contents and authors, case-insensitively', () => {
    const es = [
      entry(1, note({ key: 'S1', contents: 'invoice missing' })),
      entry(
        2,
        note({
          key: 'S2',
          contents: 'ok',
          replies: [note({ key: 'S3', contents: 'the LOGO is blurry' })],
        }),
      ),
      entry(3, note({ key: 'S4', author: 'Bob', contents: 'fine' })),
    ]
    expect(filterCommentEntries(es, 'logo').map((e) => e.root.key)).toEqual(['S2'])
    expect(filterCommentEntries(es, 'ann').map((e) => e.root.key)).toEqual(['S1', 'S2'])
    expect(filterCommentEntries(es, '').map((e) => e.root.key)).toEqual(['S1', 'S2', 'S4'])
  })

  it('orders by page then time, unknown times last within a page', () => {
    const es = [
      entry(2, note({ key: 'S5', timeMs: 100 })),
      entry(1, note({ key: 'S1', timeMs: 200 })),
      entry(1, note({ key: 'S2', timeMs: null })),
      entry(1, note({ key: 'S3', timeMs: 100 })),
    ]
    expect(filterCommentEntries(es, '').map((e) => e.root.key)).toEqual(['S3', 'S1', 'S2', 'S5'])
  })

  it('does not mutate the input array', () => {
    const es = [entry(2, note({ key: 'S2' })), entry(1, note({ key: 'S1' }))]
    const snapshot = es.map((e) => e.root.key)
    filterCommentEntries(es, '')
    expect(es.map((e) => e.root.key)).toEqual(snapshot)
  })
})

describe('countReplies', () => {
  it('counts nested replies at every depth', () => {
    const root = note({
      replies: [note({ replies: [note()] }), note()],
    })
    expect(countReplies(root)).toBe(3)
    expect(countReplies(note())).toBe(0)
  })
})
