/**
 * Plain-text paste takes the insertion point's formatting like typing (Word
 * Keep Text Only), including the pilcrow memory of a paragraph emptied by
 * Delete/Backspace/Cut — pasting a new title over a deleted line must not
 * fall back to the theme font (alpha ledger r172).
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { pasteTextSlice } from '../src/renderer/editor/paste-text'

// jsdom has no ClipboardEvent; pasteText only needs an Event shell
class FakeClipboardEvent extends Event {
  clipboardData = null
}
;(globalThis as Record<string, unknown>).ClipboardEvent = FakeClipboardEvent

const CALIBRI = { font: 'Calibri', fontAscii: 'Calibri', sizeHalfPoints: 24 }

function makeEditor(marks: { type: string; attrs?: Record<string, unknown> }[] = []) {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    // same wiring as App.tsx editorProps
    editorProps: {
      clipboardTextParser: (text, $context, _plain, view) => pasteTextSlice(text, $context, view),
    },
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            {
              type: 'text',
              text: 'Old title here',
              marks: [{ type: 'docTextStyle', attrs: CALIBRI }, ...marks],
            },
          ],
        },
        { type: 'docParagraph' },
      ],
    },
  })
}

const TEXT_LEN = 'Old title here'.length

function selectLine(editor: Editor) {
  const { state, dispatch } = editor.view
  dispatch(state.tr.setSelection(TextSelection.create(state.doc, 1, 1 + TEXT_LEN)))
}

function landedStyle(editor: Editor, paraIndex = 0): Record<string, unknown> | undefined {
  const para = editor.state.doc.child(paraIndex)
  const first = para.childCount > 0 ? para.child(0) : null
  return first?.marks.find((m) => m.type.name === 'docTextStyle')?.attrs
}

describe('plain-text paste formatting (r172)', () => {
  it('paste over a selection keeps the replaced text font', () => {
    const editor = makeEditor()
    selectLine(editor)
    editor.view.pasteText('John Lee Hooker, Hard times')
    expect(landedStyle(editor)).toMatchObject(CALIBRI)
    expect(editor.state.doc.child(0).textContent).toBe('John Lee Hooker, Hard times')
    editor.destroy()
  })

  it('select → Delete → paste into the emptied paragraph keeps the font', () => {
    const editor = makeEditor()
    selectLine(editor)
    editor.commands.deleteSelection()
    // the deletion seeds the pilcrow memory, so typing would inherit too
    expect(
      editor.state.storedMarks?.find((m) => m.type.name === 'docTextStyle')?.attrs,
    ).toMatchObject(CALIBRI)
    editor.view.pasteText('John Lee Hooker, Hard times')
    expect(landedStyle(editor)).toMatchObject(CALIBRI)
    editor.destroy()
  })

  it('multi-line paste after Delete marks every line', () => {
    const editor = makeEditor()
    selectLine(editor)
    editor.commands.deleteSelection()
    editor.view.pasteText('line one\nline two')
    expect(landedStyle(editor, 0)).toMatchObject(CALIBRI)
    expect(landedStyle(editor, 1)).toMatchObject(CALIBRI)
    editor.destroy()
  })

  it('pending ribbon format (storedMarks) wins over neighbors, like typing', () => {
    const editor = makeEditor()
    // caret at the end of the Calibri text, then a pending font change
    const { state, dispatch } = editor.view
    dispatch(state.tr.setSelection(TextSelection.create(state.doc, 1 + TEXT_LEN)))
    editor.commands.setMark('docTextStyle', { font: 'Arial', fontAscii: 'Arial' })
    editor.view.pasteText('XYZ')
    const para = editor.state.doc.child(0)
    const last = para.child(para.childCount - 1)
    expect(last.marks.find((m) => m.type.name === 'docTextStyle')?.attrs).toMatchObject({
      fontAscii: 'Arial',
    })
    editor.destroy()
  })

  it('annotation marks never ride the deletion carry onto pasted text', () => {
    const editor = makeEditor([{ type: 'comment', attrs: { id: 'c1' } }])
    selectLine(editor)
    editor.commands.deleteSelection()
    editor.view.pasteText('replacement')
    const first = editor.state.doc.child(0).child(0)
    expect(first.marks.find((m) => m.type.name === 'comment')).toBeUndefined()
    expect(first.marks.find((m) => m.type.name === 'docTextStyle')?.attrs).toMatchObject(CALIBRI)
    editor.destroy()
  })

  it('paste into a never-formatted empty paragraph stays unmarked (theme default)', () => {
    const editor = makeEditor()
    const { state, dispatch } = editor.view
    // caret into the trailing empty paragraph
    dispatch(state.tr.setSelection(TextSelection.create(state.doc, state.doc.content.size - 1)))
    editor.view.pasteText('bare')
    expect(landedStyle(editor, 1)).toBeUndefined()
    editor.destroy()
  })
})
