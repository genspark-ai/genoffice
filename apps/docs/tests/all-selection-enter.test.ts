/**
 * r125: Enter with Ctrl+A's AllSelection was a silent no-op — every command in
 * the default split chain declines at doc depth 0. Word replaces the selection
 * with a paragraph break. Exercised through a real keydown: the
 * keyboardShortcut() helper wraps the handler in one command transaction,
 * which masks the multi-dispatch behavior under test.
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { AllSelection, TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [{ type: 'text', text: 'premier paragraphe' }] },
        { type: 'docParagraph', content: [{ type: 'text', text: 'deuxieme paragraphe' }] },
      ],
    },
  })
}

const pressEnter = (editor: Editor): boolean => {
  const event = new KeyboardEvent('keydown', {
    key: 'Enter',
    code: 'Enter',
    bubbles: true,
    cancelable: true,
  })
  editor.view.dom.dispatchEvent(event)
  return event.defaultPrevented
}

describe('Enter replaces a cross-block start-to-start selection (r129)', () => {
  it('deletes the selection and splits, instead of throwing mid-dispatch', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'aaaa bbbb' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'cccc dddd' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'eeee ffff' }] },
        ],
      },
    })
    // both ends at block starts (Shift+Right from a paragraph end, then
    // Shift+Down): para2 starts at 12, para3 at 23
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 12, 23)),
    )
    expect(pressEnter(editor)).toBe(true)
    expect(editor.state.doc.textContent).toBe('aaaa bbbbeeee ffff')
    editor.destroy()
  })
})

describe('Enter keeps the replaced selection start marks (r133)', () => {
  const calibri = { type: 'docTextStyle', attrs: { font: 'Calibri', fontAscii: 'Calibri' } }
  const makeMarked = () =>
    new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'ligne un', marks: [calibri] }] },
          {
            type: 'docParagraph',
            content: [{ type: 'text', text: 'ligne deux', marks: [calibri] }],
          },
        ],
      },
    })

  it('cross-block replace stores the first deleted char marks for typing', () => {
    const editor = makeMarked()
    // Home..next block start + end: whole para 1 + into para 2 (both blocks consumed)
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, 21)),
    )
    pressEnter(editor)
    const stored = editor.state.storedMarks ?? []
    expect(stored.some((m) => m.type.name === 'docTextStyle' && m.attrs.font === 'Calibri')).toBe(
      true,
    )
    editor.destroy()
  })

  it('does NOT carry annotation marks (comment/link/revisions) onto typing', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'ligne un',
                marks: [calibri, { type: 'comment', attrs: { ids: 'c1' } }],
              },
            ],
          },
          {
            type: 'docParagraph',
            content: [{ type: 'text', text: 'ligne deux', marks: [calibri] }],
          },
        ],
      },
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, 21)),
    )
    pressEnter(editor)
    const stored = editor.state.storedMarks ?? []
    expect(stored.some((m) => m.type.name === 'docTextStyle')).toBe(true)
    expect(stored.some((m) => m.type.name === 'comment')).toBe(false)
    editor.destroy()
  })

  it('AllSelection replace stores the first char marks', () => {
    const editor = makeMarked()
    editor.view.dispatch(editor.state.tr.setSelection(new AllSelection(editor.state.doc)))
    pressEnter(editor)
    const stored = editor.state.storedMarks ?? []
    expect(stored.some((m) => m.type.name === 'docTextStyle' && m.attrs.font === 'Calibri')).toBe(
      true,
    )
    editor.destroy()
  })
})

describe('Enter-replace keeps the format on the FIRST emptied line (r133 residual)', () => {
  it('restores the carried marks when arrowing up into the first empty paragraph', () => {
    const CALIBRI = { font: 'Calibri', sizeHalfPoints: 24 }
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'texte calibri',
                marks: [{ type: 'docTextStyle', attrs: CALIBRI }],
              },
            ],
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'suite' }] },
        ],
      } as never,
    })
    // Home + Shift+Down shape: block starts of para1 and para2
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, 16)),
    )
    expect(pressEnter(editor)).toBe(true)
    // arrow up into the FIRST emptied paragraph (navigation clears storedMarks;
    // the pilcrow memory must restore them from the stamp)
    editor.commands.setTextSelection(1)
    const restored = editor.state.storedMarks
    expect(restored).not.toBeNull()
    const ts = restored!.find((m) => m.type.name === 'docTextStyle')
    expect(ts?.attrs).toMatchObject(CALIBRI)
    editor.destroy()
  })
})

describe('Enter replaces an AllSelection (r125)', () => {
  it('deletes everything and leaves a paragraph break, like Word', () => {
    const editor = makeEditor()
    editor.view.dispatch(editor.state.tr.setSelection(new AllSelection(editor.state.doc)))
    expect(pressEnter(editor)).toBe(true)
    expect(editor.state.doc.textContent).toBe('')
    expect(editor.state.doc.childCount).toBe(2)
    editor.destroy()
  })

  it('is undoable back to the original content', () => {
    const editor = makeEditor()
    editor.view.dispatch(editor.state.tr.setSelection(new AllSelection(editor.state.doc)))
    pressEnter(editor)
    editor.commands.undo()
    // the split and the delete may be separate history steps
    if (editor.state.doc.textContent === '') editor.commands.undo()
    expect(editor.state.doc.textContent).toContain('premier paragraphe')
    editor.destroy()
  })
})
