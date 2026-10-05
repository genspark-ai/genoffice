import { describe, expect, it } from 'vitest'

import { collectNoteStates } from '../src/renderer/univer-sync'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'

interface Comment {
  row: number
  column: number
  author: string
  text: string
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
    // Neither the AI set_note op nor the note editor writes an "Author:\n"
    // marker, so an author-less note whose first line ends in a colon used to be
    // read as author="Status" and lost that line on save.
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
    // The note editor rewrites the note string in place, so an authored note the
    // user edited still starts with "Author:\n". Taking the note whole would
    // fold "Dana:" into the text and lose the author column permanently.
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
