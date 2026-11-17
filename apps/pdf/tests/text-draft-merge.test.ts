import { describe, expect, it } from 'vitest'
import { mergeTextDraft } from '../src/renderer/text-edit-preview'
import type { LocalTextEdit, TextDraft } from '../src/renderer/text-edit-preview'
import type { TextBlock } from '../src/renderer/text-block'

const ctx = { pageBlocks: new Map<number, TextBlock[]>(), newId: () => 'd-test' }

const draft = (over: Partial<TextDraft> = {}): TextDraft => ({
  origIdx: 3,
  rect: [10, 20, 200, 40],
  oldText: 'Scanned words',
  fontSize: 12,
  value: 'Scanned words',
  ...over,
})

const edit = (inputOver: Partial<LocalTextEdit['input']> = {}, id = 'e1'): LocalTextEdit => ({
  id,
  input: {
    pageIndex: 3,
    rect: [10, 20, 200, 40],
    oldText: 'Scanned words',
    newText: 'Scanned words',
    fontSize: 12,
    ...inputOver,
  },
})

// A paragraph-sized block draft: one line at the block's first baseline (the
// jsdom text measure returns 0 width, so the rewrap keeps the text on one line)
const paraBlock = {
  leftPt: 10,
  firstBaseline: 30,
  widthPt: 190,
  lineHeight: 14,
  align: 'left' as const,
  bottomPt: 20,
}

const editsOf = (m: ReturnType<typeof mergeTextDraft>): LocalTextEdit[] =>
  Array.isArray(m) ? m : []

describe('mergeTextDraft: the reveal toggle on the OCR bake layer', () => {
  it('a pure reveal on a fresh draft commits — the eye toggle is itself a change', () => {
    const merged = editsOf(mergeTextDraft([], draft({ reveal: true, invisible: true }), ctx))
    expect(merged).toHaveLength(1)
    expect(merged[0]!.input.reveal).toBe(true)
  })

  it('an untouched fresh draft still merges nothing', () => {
    expect(mergeTextDraft([], draft(), ctx)).toBeNull()
  })

  it('reopening a reveal edit keeps the reveal through a later edit', () => {
    const e = edit({ newText: 'Scanned wordz', reveal: true })
    const d = draft({
      value: 'Scanned wordz2',
      reveal: true,
      invisible: true,
      editId: 'e1',
    })
    const merged = editsOf(mergeTextDraft([e], d, ctx))
    expect(merged[0]!.input.reveal).toBe(true)
    expect(merged[0]!.input.newText).toBe('Scanned wordz2')
  })

  it('reopening a reveal edit and committing it untouched still merges nothing', () => {
    // The reopen restores reveal from the pending edit; with no other change the
    // commit is a no-op (the edit already carries the reveal)
    const e = edit({ newText: 'Scanned words!', reveal: true })
    const d = draft({ value: 'Scanned words!', reveal: true, invisible: true, editId: 'e1' })
    expect(mergeTextDraft([e], d, ctx)).toBeNull()
  })

  it('reopening a reveal-only edit and toggling the eye off reverts the edit', () => {
    const e = edit({ reveal: true })
    const d = draft({ invisible: true, editId: 'e1' })
    expect(mergeTextDraft([e], d, ctx)).toEqual([])
  })

  it('a reopened edit without the eye keeps the text change invisible', () => {
    const e = edit({ newText: 'Scanned wordz' })
    const d = draft({ value: 'Scanned wordz3', editId: 'e1' })
    const merged = editsOf(mergeTextDraft([e], d, ctx))
    expect(merged[0]!.input.reveal).toBeUndefined()
    expect(merged[0]!.input.newText).toBe('Scanned wordz3')
  })

  it('a pending edit committed with the eye offered keeps the eye offered on reopen', () => {
    const e = edit({ newText: 'Scanned wordz' })
    const d = draft({ value: 'Scanned wordz2', invisible: true, editId: 'e1' })
    const merged = editsOf(mergeTextDraft([e], d, ctx))
    expect(merged[0]!.invisible).toBe(true)
  })

  it('a folded paragraph draft with the eye toggled on commits — a pure reveal is not "untouched"', () => {
    // The folded line edit stays unrevealed; the user opens the eye on the
    // paragraph draft and commits with the text untouched: the reveal must
    // reach the engine instead of reading as an untouched fold (null)
    const e1 = edit({ newText: 'Scanned wordz' })
    const e2 = edit({ oldText: 'more', newText: 'more' }, 'e2')
    const d = draft({
      value: 'Scanned wordz',
      reveal: true,
      invisible: true,
      editId: 'e1',
      foldedIds: ['e2'],
      foldBase: 'Scanned wordz',
      block: paraBlock,
    })
    const merged = editsOf(mergeTextDraft([e1, e2], d, ctx))
    expect(merged).toHaveLength(1)
    expect(merged[0]!.id).toBe('e1')
    expect(merged[0]!.input.reveal).toBe(true)
    expect(merged[0]!.invisible).toBe(true)
  })

  it('a folded paragraph draft whose eye the fold seeded commits untouched — the folded edits stay', () => {
    // The fold seeded the eye (a revealed line edit folded in): an untouched
    // commit still keeps the folded edits instead of rebuilding the paragraph
    const e1 = edit({ newText: 'Scanned wordz', reveal: true })
    const d = draft({
      value: 'Scanned wordz',
      reveal: true,
      foldReveal: true,
      invisible: true,
      editId: 'e1',
      foldBase: 'Scanned wordz',
      block: paraBlock,
    })
    expect(mergeTextDraft([e1], d, ctx)).toBeNull()
  })

  it('a revealed line edit folded into a paragraph draft keeps the reveal through a later text edit', () => {
    const e1 = edit({ newText: 'Scanned wordz', reveal: true })
    const e2 = edit({ oldText: 'more', newText: 'more' }, 'e2')
    const d = draft({
      value: 'Scanned wordz now',
      reveal: true,
      foldReveal: true,
      invisible: true,
      editId: 'e1',
      foldedIds: ['e2'],
      foldBase: 'Scanned wordz',
      block: paraBlock,
    })
    const merged = editsOf(mergeTextDraft([e1, e2], d, ctx))
    expect(merged).toHaveLength(1)
    expect(merged[0]!.input.reveal).toBe(true)
    expect(merged[0]!.input.newText).toBe('Scanned wordz now')
  })
})
