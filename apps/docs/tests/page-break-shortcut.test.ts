/**
 * Issue #126: Word's Ctrl/Cmd+Enter inserts a page break. Bound in the editor
 * keymap only — a menu accelerator would steal the key from renderer inputs
 * that use Cmd+Enter to submit (comments panel, prompt modal).
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'

function makeEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [{ type: 'text', text: 'first paragraph' }] },
        { type: 'docParagraph', content: [{ type: 'text', text: 'second paragraph' }] },
      ],
    },
  })
}

const pressModEnter = (editor: Editor): boolean => {
  const event = new KeyboardEvent('keydown', {
    key: 'Enter',
    code: 'Enter',
    ctrlKey: true,
    bubbles: true,
    cancelable: true,
  })
  editor.view.dom.dispatchEvent(event)
  return event.defaultPrevented
}

const breakCount = (editor: Editor): number => {
  let count = 0
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'docParagraph' && node.attrs.pageBreakBefore === true) count++
    return false
  })
  return count
}

describe('Mod-Enter inserts a page break', () => {
  it('adds a pageBreakBefore paragraph at the caret', () => {
    const editor = makeEditor()
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 16)))
    expect(breakCount(editor)).toBe(0)
    expect(pressModEnter(editor)).toBe(true)
    expect(breakCount(editor)).toBe(1)
    expect(editor.state.doc.textContent).toBe('first paragraphsecond paragraph')
    editor.destroy()
  })

  it('replaces a selection with the break, like other typing', () => {
    const editor = makeEditor()
    // span " paragraph" inside the first block
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 6, 16)),
    )
    expect(pressModEnter(editor)).toBe(true)
    expect(breakCount(editor)).toBe(1)
    expect(editor.state.doc.textContent).toBe('firstsecond paragraph')
    editor.destroy()
  })
})
