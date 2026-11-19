import { describe, expect, it } from 'vitest'
import { draftOwnedEditIds, previewsTextEdit } from '../src/renderer/text-edit-preview'
import type { LocalTextEdit } from '../src/renderer/text-edit-preview'

const edit = (id: string, pageIndex: number, newText: string): LocalTextEdit => ({
  id,
  input: {
    pageIndex,
    rect: [72, 500, 400, 700],
    oldText: 'Line 1 of page 3, plain text to edit.',
    newText,
    fontSize: 11,
  },
})

const paragraph = edit('paragraph', 2, 'Line 1 of page 3, plain text to edit. FIRST')
const lineA = edit('line-a', 2, 'Line 2 of page 3 A')
const lineB = edit('line-b', 2, 'Line 3 of page 3 B')
const otherPage = edit('other-page', 1, 'Line 1 of page 2 C')
const edits = [paragraph, lineA, lineB, otherPage]
const previewed = (pageIndex: number, owned: ReadonlySet<string>) =>
  edits.filter((te) => previewsTextEdit(te, pageIndex, owned))

describe('previewsTextEdit', () => {
  it('previews every pending edit of the page while no draft is open', () => {
    expect(previewed(2, draftOwnedEditIds(null))).toEqual([paragraph, lineA, lineB])
  })

  it('does not preview the pending edit a reopened draft edits', () => {
    const owned = draftOwnedEditIds({ editId: 'paragraph' })
    expect(previewed(2, owned)).toEqual([lineA, lineB])
  })

  it('does not preview the line edits a block draft folded in', () => {
    const owned = draftOwnedEditIds({ editId: 'line-a', foldedIds: ['line-b'] })
    expect(previewed(2, owned)).toEqual([paragraph])
  })

  it('a new draft owns no pending edit', () => {
    expect([...draftOwnedEditIds({})]).toEqual([])
    expect(previewed(1, draftOwnedEditIds({}))).toEqual([otherPage])
  })
})
