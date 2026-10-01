import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { nextNoteId, type NoteInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { submitNote, type ReviewContext } from '../src/renderer/review-actions'

/** the in-text reference marks of one kind, in document order */
function marks(editor: Editor, kind: 'footnote' | 'endnote'): Array<{ id: string; num: number }> {
  const out: Array<{ id: string; num: number }> = []
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'docNoteRef' || node.attrs.kind !== kind) return true
    out.push({ id: String(node.attrs.id), num: Number(node.attrs.num) })
    return false
  })
  return out
}

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [
            { type: 'text', text: 'alpha' },
            { type: 'docNoteRef', attrs: { kind: 'footnote', id: '3', num: 1 } },
          ],
        },
        {
          type: 'docParagraph',
          attrs: { docxIndex: 1 },
          content: [
            { type: 'text', text: 'beta' },
            { type: 'docNoteRef', attrs: { kind: 'footnote', id: '5', num: 2 } },
          ],
        },
      ],
    },
  })
}

function makeCtx(editor: Editor, footnotes: NoteInfo[]): { ctx: ReviewContext; state: NoteInfo[] } {
  const state = [...footnotes]
  const ctx = {
    editor,
    notePrompt: { kind: 'footnote' },
    setFootnotes: (value: unknown) => {
      const next =
        typeof value === 'function'
          ? (value as (prev: NoteInfo[]) => NoteInfo[])(state)
          : (value as NoteInfo[])
      state.splice(0, state.length, ...next)
    },
    setNotesDirty: () => undefined,
    footnotes,
  } as unknown as ReviewContext
  return { ctx, state }
}

describe('inserting a note above the existing ones', () => {
  it('renumbers the marks already in the body into document order', () => {
    const editor = makeEditor()
    // the caret sits before both existing marks, so the new one comes first
    editor.commands.setTextSelection(1)
    const { ctx, state } = makeCtx(editor, [
      { id: '3', text: 'first note' },
      { id: '5', text: 'second note' },
    ])

    submitNote(ctx, 'inserted note')

    const newId = nextNoteId([
      { id: '3', text: '' },
      { id: '5', text: '' },
    ])
    // document order is the new mark, then the two that were already there
    expect(marks(editor, 'footnote')).toEqual([
      { id: newId, num: 1 },
      { id: '3', num: 2 },
      { id: '5', num: 3 },
    ])
    // the note list itself still appends
    expect(state.map((n) => n.id)).toEqual(['3', '5', newId])
    editor.destroy()
  })

  it('leaves the marks of the other kind alone', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0 },
            content: [
              { type: 'text', text: 'alpha' },
              { type: 'docNoteRef', attrs: { kind: 'endnote', id: '9', num: 1 } },
              { type: 'docNoteRef', attrs: { kind: 'footnote', id: '3', num: 1 } },
            ],
          },
        ],
      },
    })
    editor.commands.setTextSelection(1)
    const { ctx } = makeCtx(editor, [{ id: '3', text: 'a footnote' }])

    submitNote(ctx, 'inserted note')

    expect(marks(editor, 'endnote')).toEqual([{ id: '9', num: 1 }])
    expect(marks(editor, 'footnote').map((m) => m.num)).toEqual([1, 2])
    editor.destroy()
  })
})
