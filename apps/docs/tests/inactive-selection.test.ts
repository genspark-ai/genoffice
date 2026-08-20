/**
 * r119: while a ribbon combobox holds focus, the document selection must stay
 * visibly highlighted — the decoration mirrors the state selection whenever a
 * control opts in via setInactiveSelectionShown.
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { AllSelection, TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { setInactiveSelectionShown } from '../src/renderer/editor/inactive-selection'

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [{ type: 'text', text: 'bassiste de David Bowie' }] },
      ],
    },
  })
}

const highlighted = (editor: Editor): string | null =>
  editor.view.dom.querySelector('.doc-inactive-selection')?.textContent ?? null

describe('inactive selection highlight (r119)', () => {
  it('decorates the selected range while shown, clears on hide', () => {
    const editor = makeEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, 9)))
    expect(highlighted(editor)).toBeNull()
    setInactiveSelectionShown(editor, true)
    expect(highlighted(editor)).toBe('bassiste')
    setInactiveSelectionShown(editor, false)
    expect(highlighted(editor)).toBeNull()
    editor.destroy()
  })

  it('highlights a ctrl-A AllSelection too (bugbot: non-text selections)', () => {
    const editor = makeEditor()
    editor.view.dispatch(editor.state.tr.setSelection(new AllSelection(editor.state.doc)))
    setInactiveSelectionShown(editor, true)
    expect(highlighted(editor)).toBe('bassiste de David Bowie')
    editor.destroy()
  })

  it('does nothing for a collapsed caret', () => {
    const editor = makeEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 3, 3)))
    setInactiveSelectionShown(editor, true)
    expect(highlighted(editor)).toBeNull()
    editor.destroy()
  })

  it('follows edits while shown (decoration maps with the doc)', () => {
    const editor = makeEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, 9)))
    setInactiveSelectionShown(editor, true)
    editor.view.dispatch(editor.state.tr.insertText('X', 1, 1))
    // decoration recomputes from the mapped selection every render
    expect(highlighted(editor)).toContain('assiste')
    editor.destroy()
  })
})
