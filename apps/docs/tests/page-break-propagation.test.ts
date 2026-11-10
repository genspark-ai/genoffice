/**
 * r157: Enter inside a pageBreakBefore paragraph must leave the break on
 * exactly one half — the one starting with the original content. Cloning it
 * onto the new paragraph turned a plain newline on a page-top title into
 * another page jump.
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { insertPageBreak } from '../src/renderer/editor/page-break'

function makeEditor(blocks: Array<{ type?: string; pbb?: boolean; text?: string }>): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: blocks.map((b) => ({
        type: b.type ?? 'docParagraph',
        attrs: b.pbb ? { pageBreakBefore: true } : {},
        content: b.text ? [{ type: 'text', text: b.text }] : [],
      })),
    } as never,
  })
}

const pressEnter = (editor: Editor): void => {
  const event = new KeyboardEvent('keydown', {
    key: 'Enter',
    code: 'Enter',
    bubbles: true,
    cancelable: true,
  })
  editor.view.dom.dispatchEvent(event)
}

const shape = (editor: Editor): Array<[string, boolean, string]> => {
  const out: Array<[string, boolean, string]> = []
  editor.state.doc.forEach((n) => out.push([n.type.name, !!n.attrs.pageBreakBefore, n.textContent]))
  return out
}

/** caret at parentOffset `off` inside doc child 1 (the break paragraph) */
function setCaret(editor: Editor, off: number): void {
  const first = editor.state.doc.child(0).nodeSize
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, first + 1 + off)),
  )
}

describe('r157: Enter does not clone pageBreakBefore', () => {
  it('Enter at the end of a page-top title keeps the break on the title only', () => {
    const editor = makeEditor([{ text: 'before' }, { pbb: true, text: 'Title' }, { text: 'after' }])
    setCaret(editor, 5)
    pressEnter(editor)
    expect(shape(editor)).toEqual([
      ['docParagraph', false, 'before'],
      ['docParagraph', true, 'Title'],
      ['docParagraph', false, ''],
      ['docParagraph', false, 'after'],
    ])
    editor.destroy()
  })

  it('Enter mid-paragraph keeps the break on the first half only', () => {
    const editor = makeEditor([{ text: 'before' }, { pbb: true, text: 'Title' }, { text: 'after' }])
    setCaret(editor, 2)
    pressEnter(editor)
    expect(shape(editor)).toEqual([
      ['docParagraph', false, 'before'],
      ['docParagraph', true, 'Ti'],
      ['docParagraph', false, 'tle'],
      ['docParagraph', false, 'after'],
    ])
    editor.destroy()
  })

  it('Enter at the start moves the break to the content half, not the new empty line', () => {
    const editor = makeEditor([{ text: 'before' }, { pbb: true, text: 'Title' }, { text: 'after' }])
    setCaret(editor, 0)
    pressEnter(editor)
    expect(shape(editor)).toEqual([
      ['docParagraph', false, 'before'],
      ['docParagraph', false, ''],
      ['docParagraph', true, 'Title'],
      ['docParagraph', false, 'after'],
    ])
    editor.destroy()
  })

  it('Enter at the end of a page-top heading behaves the same', () => {
    const editor = makeEditor([
      { text: 'before' },
      { type: 'docHeading', pbb: true, text: 'Title' },
      { text: 'after' },
    ])
    setCaret(editor, 5)
    pressEnter(editor)
    const rows = shape(editor)
    expect(rows[1]).toEqual(['docHeading', true, 'Title'])
    expect(rows[2][1]).toBe(false)
    editor.destroy()
  })

  it('Ctrl+Enter inside an existing break paragraph still yields two breaks (#1103)', () => {
    const editor = makeEditor([
      { text: 'before' },
      { pbb: true, text: 'TitleTail' },
      { text: 'after' },
    ])
    setCaret(editor, 5)
    insertPageBreak(editor)
    expect(shape(editor)).toEqual([
      ['docParagraph', false, 'before'],
      ['docParagraph', true, 'Title'],
      ['docParagraph', true, 'Tail'],
      ['docParagraph', false, 'after'],
    ])
    editor.destroy()
  })
})
