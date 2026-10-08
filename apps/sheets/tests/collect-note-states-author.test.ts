import { describe, expect, it } from 'vitest'

import { threadStore } from '../src/renderer/threaded-comments'
import { collectNoteStates } from '../src/renderer/univer-sync'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'

interface Comment {
  row: number
  column: number
  author: string
  text: string
  thread?: {
    id: string
    personId: string
    author: string
    dT: string
    done: boolean
    text: string
    replies: never[]
  }
}

const encodeNoteText = (author: string, text: string): string =>
  author ? `${author}:\n${text}` : text

function harness(comments: Comment[], live: { row: number; col: number; note: string }[]) {
  const worksheet = { getNotes: () => live }
  const workbook = { getSheetBySheetId: (id: string) => (id === 'sheet1' ? worksheet : undefined) }
  const runtime = {
    univerAPI: { getActiveWorkbook: () => workbook },
  } as unknown as UniverRuntime
  const state = {
    editJournal: { noteDirty: new Set(['sheet1']), sheets: { removed: new Set<string>() } },
    file: { sheets: [{ id: 'sheet1', comments }] },
  } as unknown as LazyWorkbookState
  return { runtime, state }
}

describe('collectNoteStates author marker', () => {
  it('keeps a note that was authored this session whole', () => {
    const authored = { row: 0, col: 0, note: 'Status:\nOn track' }
    const { runtime, state } = harness(
      [{ row: 0, column: 0, author: '', text: authored.note }],
      [authored],
    )
    const [sheet] = collectNoteStates(runtime, state)
    expect(sheet?.notes).toEqual([{ row: 0, column: 0, author: '', text: 'Status:\nOn track' }])
  })

  it('still recovers the author of a note the file installed', () => {
    const comment = { row: 1, column: 2, author: 'Dana', text: 'check Q3\nand Q4' }
    const live = { row: 1, col: 2, note: encodeNoteText(comment.author, comment.text) }
    const { runtime, state } = harness([comment], [live])
    const [sheet] = collectNoteStates(runtime, state)
    expect(sheet?.notes).toEqual([{ row: 1, column: 2, author: 'Dana', text: 'check Q3\nand Q4' }])
  })

  it('keeps an author-less note the file installed verbatim', () => {
    const comment = { row: 0, column: 0, author: '', text: 'TODO:\n- call Bob' }
    const live = { row: 0, col: 0, note: comment.text }
    const { runtime, state } = harness([comment], [live])
    const [sheet] = collectNoteStates(runtime, state)
    expect(sheet?.notes).toEqual([{ row: 0, column: 0, author: '', text: 'TODO:\n- call Bob' }])
  })

  it('keeps the author of a session-edited note and does not fold it into the text', () => {
    const comment = { row: 1, column: 2, author: 'Dana', text: 'original' }
    const live = { row: 1, col: 2, note: 'Dana:\nedited this session' }
    const { runtime, state } = harness([comment], [live])
    const [sheet] = collectNoteStates(runtime, state)
    expect(sheet?.notes).toEqual([
      { row: 1, column: 2, author: 'Dana', text: 'edited this session' },
    ])
  })

  it('does not mistake a label first line for the author of an authored note', () => {
    // the author's name must be the one the file recorded, not a shape match
    const comment = { row: 1, column: 2, author: 'Dana', text: 'original' }
    const live = { row: 1, col: 2, note: 'Status:\nOn track' }
    const { runtime, state } = harness([comment], [live])
    const [sheet] = collectNoteStates(runtime, state)
    expect(sheet?.notes).toEqual([{ row: 1, column: 2, author: '', text: 'Status:\nOn track' }])
  })
})

describe('collectNoteStates threaded comments', () => {
  const thread = {
    id: '{R1}',
    personId: '{P1}',
    author: 'Ada',
    dT: '2024-01-01T00:00:00.00',
    done: false,
    text: 'Root',
    replies: [] as never[],
  }

  it('emits store threads next to the live notes and skips a note on a threaded cell', () => {
    const comment = { row: 0, column: 0, author: 'Ada', text: 'Root', thread }
    const { runtime, state } = harness(
      [comment],
      [
        { row: 0, col: 0, note: 'stale' },
        { row: 2, col: 2, note: 'plain' },
      ],
    )
    threadStore.load({ sheets: [{ id: 'sheet1', comments: [comment] }] } as never)
    threadStore.setAuthor('Grace')
    threadStore.reply('sheet1', 0, 0, 'Second')
    const [sheet] = collectNoteStates(runtime, state)
    expect(sheet?.notes).toEqual([
      { row: 2, column: 2, author: '', text: 'plain' },
      {
        row: 0,
        column: 0,
        author: 'Ada',
        text: 'Root',
        thread: {
          id: '{R1}',
          personId: '{P1}',
          author: 'Ada',
          dT: '2024-01-01T00:00:00.00',
          done: false,
          replies: [expect.objectContaining({ author: 'Grace', text: 'Second' })],
        },
      },
    ])
  })

  it('marks a plain note on a cell the file had a thread on as converted (thread: null)', () => {
    const comment = { row: 0, column: 0, author: 'Ada', text: 'Root', thread }
    const { runtime, state } = harness([comment], [{ row: 0, col: 0, note: 'Ada:\nRoot' }])
    threadStore.load({ sheets: [] } as never)
    const [sheet] = collectNoteStates(runtime, state)
    expect(sheet?.notes).toEqual([{ row: 0, column: 0, author: 'Ada', text: 'Root', thread: null }])
  })

  it('splits the author off a note converted from a thread created this session', () => {
    const { runtime, state } = harness([], [{ row: 4, col: 4, note: 'Grace:\nfresh thread' }])
    threadStore.load({ sheets: [] } as never)
    threadStore.setAuthor('Grace')
    threadStore.add('sheet1', 4, 4, 'fresh thread')
    threadStore.convertToNote('sheet1', 4, 4)
    const [sheet] = collectNoteStates(runtime, state)
    expect(sheet?.notes).toEqual([
      { row: 4, column: 4, author: 'Grace', text: 'fresh thread', thread: null },
    ])
  })
})
