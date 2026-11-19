import { describe, expect, it } from 'vitest'

import { collectNoteStates } from '../src/renderer/univer-sync'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'
import type { NotePictureAdd } from '../src/renderer/edit-journal'

/**
 * The picture a reader put into a note this session.
 *
 * Univer's note model has no picture of its own — the note popup is a plain
 * text box — so the picture travels beside the note in the save request, and
 * the saver learns about it from the journal.
 */

const DATA_URL = 'data:image/png;base64,AAAA'

function harness(
  live: { row: number; col: number; note: string }[],
  notePictures: NotePictureAdd[] = [],
) {
  const worksheet = { getNotes: () => live }
  const workbook = { getSheetBySheetId: (id: string) => (id === 'sheet1' ? worksheet : undefined) }
  const runtime = {
    univerAPI: { getActiveWorkbook: () => workbook },
  } as unknown as UniverRuntime
  const state = {
    editJournal: {
      noteDirty: new Set(['sheet1']),
      notePictures,
      sheets: { removed: new Set<string>() },
    },
    file: { sheets: [{ id: 'sheet1', comments: [] }] },
  } as unknown as LazyWorkbookState
  return { runtime, state }
}

const picture: NotePictureAdd['picture'] = {
  mediaType: 'image/png',
  dataUrl: DATA_URL,
  widthPx: 128,
  heightPx: 106,
  name: 'pedigree.png',
}

describe('a picture put into a note', () => {
  it('rides along with the note it belongs to', () => {
    const { runtime, state } = harness(
      [{ row: 0, col: 1, note: 'the pedigree' }],
      [{ sheetId: 'sheet1', row: 0, column: 1, picture }],
    )
    const [sheet] = collectNoteStates(runtime, state)
    expect(sheet?.notes).toEqual([
      {
        row: 0,
        column: 1,
        author: '',
        text: 'the pedigree',
        picture: {
          mediaType: 'image/png',
          base64: 'AAAA',
          widthPx: 128,
          heightPx: 106,
          name: 'pedigree.png',
        },
      },
    ])
  })

  it('says nothing about the picture of a note nobody touched', () => {
    // a note that says nothing keeps the picture the file already had, so the
    // saver must not be told about a picture here
    const { runtime, state } = harness([{ row: 0, col: 1, note: 'plain' }])
    const [sheet] = collectNoteStates(runtime, state)
    expect(sheet?.notes[0]?.picture).toBeUndefined()
  })

  it('takes the picture away when the note says so', () => {
    const { runtime, state } = harness(
      [{ row: 0, col: 1, note: 'plain' }],
      [{ sheetId: 'sheet1', row: 0, column: 1, picture: null }],
    )
    const [sheet] = collectNoteStates(runtime, state)
    expect(sheet?.notes[0]?.picture).toBeNull()
  })

  it('leaves a note without a payload alone', () => {
    // a data URL without its payload carries no picture
    const { runtime, state } = harness(
      [{ row: 0, col: 1, note: 'plain' }],
      [
        {
          sheetId: 'sheet1',
          row: 0,
          column: 1,
          picture: { ...picture, dataUrl: 'data:image/png;base64,' },
        },
      ],
    )
    const [sheet] = collectNoteStates(runtime, state)
    expect(sheet?.notes[0]?.picture).toBeUndefined()
  })

  it('keeps a picture off a note in another sheet', () => {
    const { runtime, state } = harness(
      [{ row: 0, col: 1, note: 'the pedigree' }],
      [{ sheetId: 'sheet2', row: 0, column: 1, picture }],
    )
    const [sheet] = collectNoteStates(runtime, state)
    expect(sheet?.notes[0]?.picture).toBeUndefined()
  })
})
