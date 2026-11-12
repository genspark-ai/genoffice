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

const edit = (inputOver: Partial<LocalTextEdit['input']> = {}): LocalTextEdit => ({
  id: 'e1',
  input: {
    pageIndex: 3,
    rect: [10, 20, 200, 40],
    oldText: 'Scanned words',
    newText: 'Scanned words',
    fontSize: 12,
    ...inputOver,
  },
})

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
})
