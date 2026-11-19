import { describe, expect, it } from 'vitest'

import type { WorkbookFile } from '../src/shared/desktop-api'
import { collectNotePictures } from '../src/renderer/note-pictures'

/** Only the fields `collectNotePictures` reads; the file type is far wider. */
function fileWith(sheets: unknown[]): WorkbookFile {
  return {
    sessionId: 'session-1',
    sheets,
  } as unknown as WorkbookFile
}

const sheet = (id: string, name: string, comments: unknown[]) => ({ id, name, comments })

const note = (row: number, column: number, image?: string) => ({
  row,
  column,
  author: 'Ada',
  text: `note at ${row},${column}`,
  ...(image === undefined ? {} : { image: { id: image } }),
})

describe('collectNotePictures', () => {
  it('lists only the notes that are showing a picture', () => {
    const pictures = collectNotePictures(
      fileWith([
        sheet('s1', 'Data', [
          note(0, 0),
          note(1, 2, 'note-picture-1'),
          note(4, 0, 'note-picture-2'),
        ]),
      ]),
    )
    expect(pictures.map((picture) => [picture.row, picture.column])).toEqual([
      [1, 2],
      [4, 0],
    ])
    expect(pictures[0]?.imageId).toBe('note-picture-1')
    expect(pictures[0]?.sheetName).toBe('Data')
    expect(pictures[0]?.author).toBe('Ada')
  })

  it('carries a picture across sheets, keeping sheet order', () => {
    const pictures = collectNotePictures(
      fileWith([
        sheet('s1', 'First', [note(0, 0, 'note-picture-1')]),
        sheet('s2', 'Second', [note(0, 0, 'note-picture-2')]),
      ]),
    )
    expect(pictures.map((picture) => picture.sheetName)).toEqual(['First', 'Second'])
    expect(pictures.map((picture) => picture.imageId)).toEqual(['note-picture-1', 'note-picture-2'])
  })

  it('is empty for a workbook with no pictures, and for no workbook at all', () => {
    expect(collectNotePictures(null)).toEqual([])
    expect(collectNotePictures(fileWith([sheet('s1', 'Data', [note(0, 0)])]))).toEqual([])
    expect(collectNotePictures(fileWith([]))).toEqual([])
  })
})

describe('a picture put into a note this session', () => {
  const pending = (row: number, column: number, dataUrl: string | null) => ({
    sheetId: 's1',
    row,
    column,
    dataUrl,
  })

  it('shows a picture the file did not have yet', () => {
    const pictures = collectNotePictures(fileWith([sheet('s1', 'Data', [note(0, 0)])]), [
      pending(3, 2, 'data:image/png;base64,AAAA'),
    ])
    expect(pictures).toHaveLength(1)
    expect(pictures[0]?.dataUrl).toBe('data:image/png;base64,AAAA')
    expect(pictures[0]?.sheetName).toBe('Data')
    expect([pictures[0]?.row, pictures[0]?.column]).toEqual([3, 2])
  })

  it('replaces the picture the file had for that note', () => {
    const pictures = collectNotePictures(
      fileWith([sheet('s1', 'Data', [note(1, 2, 'note-picture-1')])]),
      [pending(1, 2, 'data:image/png;base64,BBBB')],
    )
    expect(pictures).toHaveLength(1)
    expect(pictures[0]?.dataUrl).toBe('data:image/png;base64,BBBB')
    expect(pictures[0]?.imageId).toBeUndefined()
  })

  it('stops showing a picture that was taken away', () => {
    const pictures = collectNotePictures(
      fileWith([sheet('s1', 'Data', [note(1, 2, 'note-picture-1')])]),
      [pending(1, 2, null)],
    )
    expect(pictures).toEqual([])
  })
})
