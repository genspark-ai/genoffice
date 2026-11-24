/**
 * The caret and selection moves, word deletion and character effects behind
 * the chords Word for Mac has and Chromium does not, or answers differently.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { deleteWord, editText, moveCaret, parseCaretMove } from '../src/renderer/editor/caret-keys'

let editor: Editor

function make(paragraphs: string[]): Editor {
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: paragraphs.map((text) => ({
        type: 'docParagraph',
        content: text ? [{ type: 'text', text }] : [],
      })),
    },
  })
  return editor
}

afterEach(() => editor?.destroy())

/** [paragraph index, offset in it] of a document position */
function place(pos: number): [number, number] {
  const $pos = editor.state.doc.resolve(pos)
  return [$pos.index(0), $pos.parentOffset]
}

function caretAt(paragraph: number, offset: number): void {
  let pos = 0
  editor.state.doc.forEach((_, start, index) => {
    if (index === paragraph) pos = start + 1 + offset
  })
  editor.commands.setTextSelection(pos)
}

describe('caret moves', () => {
  it('reads a move and its extending twin', () => {
    expect(parseCaretMove('para-up')).toEqual({ to: 'para-up', extend: false })
    expect(parseCaretMove('doc-end-extend')).toEqual({ to: 'doc-end', extend: true })
    expect(parseCaretMove('sideways')).toBe(null)
    expect(parseCaretMove('-extend')).toBe(null)
  })

  it('goes to the start of the paragraph, then of the one before', () => {
    make(['first one', 'second one', 'third'])
    caretAt(1, 4)
    moveCaret(editor, 'para-up', false)
    expect(place(editor.state.selection.head)).toEqual([1, 0])
    expect(editor.state.selection.empty).toBe(true)
    moveCaret(editor, 'para-up', false)
    expect(place(editor.state.selection.head)).toEqual([0, 0])
    moveCaret(editor, 'para-up', false)
    expect(place(editor.state.selection.head)).toEqual([0, 0])
  })

  it('goes to the start of the next paragraph, and to the end of the last', () => {
    make(['first one', 'second one', 'third'])
    caretAt(0, 2)
    moveCaret(editor, 'para-down', false)
    expect(place(editor.state.selection.head)).toEqual([1, 0])
    moveCaret(editor, 'para-down', false)
    expect(place(editor.state.selection.head)).toEqual([2, 0])
    moveCaret(editor, 'para-down', false)
    expect(place(editor.state.selection.head)).toEqual([2, 5])
  })

  it('extends by paragraph from where the selection started', () => {
    make(['first one', 'second one', 'third'])
    caretAt(1, 4)
    moveCaret(editor, 'para-down', true)
    expect(place(editor.state.selection.anchor)).toEqual([1, 4])
    expect(place(editor.state.selection.head)).toEqual([2, 0])
    moveCaret(editor, 'para-up', true)
    expect(place(editor.state.selection.head)).toEqual([1, 0])
    moveCaret(editor, 'para-up', true)
    expect(place(editor.state.selection.anchor)).toEqual([1, 4])
    expect(place(editor.state.selection.head)).toEqual([0, 0])
  })

  it('goes and extends to the ends of the document', () => {
    make(['first one', 'second one', 'third'])
    caretAt(1, 4)
    moveCaret(editor, 'doc-end', false)
    expect(place(editor.state.selection.head)).toEqual([2, 5])
    moveCaret(editor, 'doc-start', false)
    expect(place(editor.state.selection.head)).toEqual([0, 0])
    caretAt(1, 4)
    moveCaret(editor, 'doc-start', true)
    expect(place(editor.state.selection.anchor)).toEqual([1, 4])
    expect(place(editor.state.selection.head)).toEqual([0, 0])
    moveCaret(editor, 'doc-end', true)
    expect(place(editor.state.selection.head)).toEqual([2, 5])
  })

  it('at a paragraph edge a word deletion is the Backspace or Delete of that place', () => {
    make(['first', 'second'])
    caretAt(1, 0)
    deleteWord(editor, -1)
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.getText()).toBe('firstsecond')
    make(['first', 'second'])
    caretAt(0, 5)
    deleteWord(editor, 1)
    expect(editor.state.doc.childCount).toBe(1)
  })

  it('leaves a read-only document alone on Cut and Paste', () => {
    make(['first one'])
    editor.commands.selectAll()
    editor.setEditable(false)
    const run = vi.fn(() => true)
    document.execCommand = run
    editText(editor, 'cut')
    editText(editor, 'paste')
    expect(run).not.toHaveBeenCalled()
    editText(editor, 'copy')
    expect(run).toHaveBeenCalledWith('copy')
    expect(editor.getText()).toBe('first one')
  })

  it('deletes a selection as it stands', () => {
    make(['first one'])
    editor.commands.setTextSelection({ from: 1, to: 7 })
    deleteWord(editor, -1)
    expect(editor.getText()).toBe('one')
  })
})
