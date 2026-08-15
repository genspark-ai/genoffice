import { describe, expect, it } from 'vitest'
import {
  buildNoteThreads,
  findThreadRoot,
  flattenThread,
  parsePdfDate,
  threadSubtree,
  toSavedNote,
} from '../src/renderer/note-threads'
import type { NoteInput, PdfJsAnnotData, SavedNoteAnnot } from '../src/renderer/note-threads'

const saved = (over: Partial<SavedNoteAnnot>): SavedNoteAnnot => ({
  pageIndex: 0,
  objNum: 1,
  type: 'note',
  rect: [100, 682, 120, 700],
  color: [1, 0.78, 0.13],
  author: 'Alice',
  contents: 'root',
  timeMs: 1000,
  inReplyTo: null,
  ...over,
})

const pending = (id: string, over: Partial<NoteInput> = {}): { id: string; input: NoteInput } => ({
  id,
  input: {
    kind: 'note',
    pageIndex: 0,
    color: [1, 0, 0],
    at: [50, 50],
    contents: `pending ${id}`,
    localId: id,
    ...over,
  },
})

describe('parsePdfDate', () => {
  it('parses a full PDF date with timezone offset', () => {
    // 2026-08-12 17:30:00 +08:00 == 09:30:00 UTC
    expect(parsePdfDate("D:20260812173000+08'00'")).toBe(Date.UTC(2026, 7, 12, 9, 30, 0))
  })

  it('parses Z as UTC and tolerates missing trailing fields', () => {
    expect(parsePdfDate('D:20260812Z')).toBe(Date.UTC(2026, 7, 12))
    expect(parsePdfDate('D:2026')).toBe(new Date(2026, 0, 1).getTime())
  })

  it('rejects garbage and out-of-range fields', () => {
    expect(parsePdfDate(undefined)).toBeNull()
    expect(parsePdfDate('yesterday')).toBeNull()
    expect(parsePdfDate('D:20261399')).toBeNull()
  })
})

describe('toSavedNote', () => {
  const raw = (over: Partial<PdfJsAnnotData> = {}): PdfJsAnnotData => ({
    id: '12R',
    annotationType: 1,
    rect: [100, 682, 120, 700],
    titleObj: { str: 'WPS User' },
    contentsObj: { str: 'A comment' },
    creationDate: "D:20260812173000+08'00'",
    inReplyTo: null,
    ...over,
  })

  it('converts a Text annotation with a reply link', () => {
    const note = toSavedNote(raw({ inReplyTo: '9R' }), 3)
    expect(note).toMatchObject({
      pageIndex: 3,
      objNum: 12,
      author: 'WPS User',
      contents: 'A comment',
      inReplyTo: 9,
    })
    expect(note!.timeMs).toBe(Date.UTC(2026, 7, 12, 9, 30, 0))
  })

  it('skips non-Text, hidden, Group and ref-less annotations', () => {
    expect(toSavedNote(raw({ annotationType: 9 }), 0)).toBeNull()
    expect(toSavedNote(raw({ hidden: true }), 0)).toBeNull()
    expect(toSavedNote(raw({ replyType: 'Group' }), 0)).toBeNull()
    expect(toSavedNote(raw({ id: 'annot_x' }), 0)).toBeNull()
  })
})

describe('buildNoteThreads', () => {
  it('chains saved replies under their root sorted by time', () => {
    const roots = buildNoteThreads(
      [
        saved({ objNum: 1, contents: 'root', timeMs: 1000 }),
        saved({ objNum: 3, contents: 'late', inReplyTo: 1, timeMs: 3000 }),
        saved({ objNum: 2, contents: 'early', inReplyTo: 1, timeMs: 2000 }),
      ],
      [],
    )
    expect(roots).toHaveLength(1)
    expect(roots[0]!.replies.map((r) => r.contents)).toEqual(['early', 'late'])
  })

  it('promotes orphaned replies to roots instead of dropping them', () => {
    const roots = buildNoteThreads([saved({ objNum: 5, inReplyTo: 99, contents: 'orphan' })], [])
    expect(roots).toHaveLength(1)
    expect(roots[0]!.contents).toBe('orphan')
  })

  it('keeps every note visible when /IRT forms a cycle', () => {
    const roots = buildNoteThreads(
      [saved({ objNum: 1, inReplyTo: 2 }), saved({ objNum: 2, inReplyTo: 1 })],
      [],
    )
    const seen = roots.flatMap((r) => flattenThread(r)).map(({ item }) => item.key)
    expect(new Set(seen).size).toBe(2)
  })

  it('attaches pending replies to saved parents and to pending parents', () => {
    const roots = buildNoteThreads(
      [saved({ objNum: 7 })],
      [
        pending('a'),
        pending('b', {
          replyToSaved: { objNum: 7, rect: [100, 682, 120, 700], contents: 'root' },
        }),
        pending('c', { replyToLocalId: 'a' }),
      ],
    )
    const rootS = roots.find((r) => r.key === 'S7')!
    const rootA = roots.find((r) => r.key === 'Pa')!
    expect(roots).toHaveLength(2)
    expect(rootS.replies.map((r) => r.key)).toEqual(['Pb'])
    expect(rootA.replies.map((r) => r.key)).toEqual(['Pc'])
  })

  it('threadSubtree collects the mixed saved/pending cascade of a delete', () => {
    const roots = buildNoteThreads(
      [saved({ objNum: 1 }), saved({ objNum: 2, inReplyTo: 1 })],
      [pending('x', { replyToSaved: { objNum: 2, rect: [100, 682, 120, 700], contents: 'root' } })],
    )
    const { saved: savedHits, pendingIds } = threadSubtree(roots[0]!)
    expect(savedHits.map((s) => s.objNum).sort()).toEqual([1, 2])
    expect(pendingIds).toEqual(['x'])
  })

  it('findThreadRoot resolves the root from any member key', () => {
    const roots = buildNoteThreads([saved({ objNum: 1 }), saved({ objNum: 2, inReplyTo: 1 })], [])
    expect(findThreadRoot(roots, 'S2')?.key).toBe('S1')
    expect(findThreadRoot(roots, 'S99')).toBeNull()
  })
})
