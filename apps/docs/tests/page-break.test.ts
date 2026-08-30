/**
 * Word Ctrl+Enter parity (alpha ledger r154): a page break in front of text
 * moves that text to the next page with NO coupled blank line, and deleting
 * at the break merges the paragraphs (clearing the break) without content
 * loss.
 */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { insertPageBreak } from '../src/renderer/editor/page-break'

const makeEditor = (text: string) =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
    },
  })

describe('insertPageBreak (r154)', () => {
  it('mid-paragraph: splits, second half carries the break, no empty paragraph', () => {
    const editor = makeEditor('hello world')
    editor.commands.setTextSelection(7) // before "world"
    insertPageBreak(editor)
    const { doc } = editor.state
    expect(doc.childCount).toBe(2)
    expect(doc.child(0).textContent).toBe('hello ')
    expect(doc.child(1).textContent).toBe('world')
    expect(doc.child(0).attrs.pageBreakBefore).toBeFalsy()
    expect(doc.child(1).attrs.pageBreakBefore).toBe(true)
    editor.destroy()
  })

  it('document start: a leading break character, not the paragraph attribute', () => {
    // The attribute is Word's paragraph property — a no-op when the block
    // already sits at a page top, so Ctrl+Enter at the document start would
    // do nothing (bugbot). Pagination honors a leading w:br even onto the
    // blank first page.
    const editor = makeEditor('world')
    editor.commands.setTextSelection(1)
    insertPageBreak(editor)
    const { doc } = editor.state
    expect(doc.childCount).toBe(1)
    const para = doc.child(0)
    expect(para.textContent).toBe('world')
    expect(para.attrs.pageBreakBefore).toBeFalsy()
    const first = para.child(0)
    expect(first.type.name).toBe('hardBreak')
    expect(first.attrs.pageBreak).toBe(true)
    editor.destroy()
  })

  it('paragraph start elsewhere: the block takes the attribute, no line box added', () => {
    // Outside the document start a leading <br> would render a blank first
    // line at the new page top (bugbot); the attribute path is artifact-free
    // and pagination honors it on any non-blank page.
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'hello' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'world' }] },
        ],
      },
    })
    editor.commands.setTextSelection(8) // start of "world"
    insertPageBreak(editor)
    const { doc } = editor.state
    expect(doc.childCount).toBe(2)
    expect(doc.child(1).textContent).toBe('world')
    expect(doc.child(1).attrs.pageBreakBefore).toBe(true)
    expect(doc.child(1).child(0).type.name).toBe('text')
    editor.destroy()
  })

  it('mid-paragraph: one undo reverts the whole break', () => {
    const editor = makeEditor('hello world')
    editor.commands.setTextSelection(7)
    insertPageBreak(editor)
    editor.commands.undo()
    const { doc } = editor.state
    expect(doc.childCount).toBe(1)
    expect(doc.child(0).textContent).toBe('hello world')
    expect(doc.child(0).attrs.pageBreakBefore).toBeFalsy()
    editor.destroy()
  })

  it('document start: backspace removes the break character and nothing else', () => {
    const editor = makeEditor('world')
    editor.commands.setTextSelection(1)
    insertPageBreak(editor)
    editor.commands.deleteRange({
      from: editor.state.selection.from - 1,
      to: editor.state.selection.from,
    })
    const { doc } = editor.state
    expect(doc.childCount).toBe(1)
    expect(doc.child(0).textContent).toBe('world')
    expect(doc.child(0).childCount).toBe(1)
    editor.destroy()
  })

  it('paragraph end: empty second half carries the break (Word behavior)', () => {
    const editor = makeEditor('hello')
    editor.commands.setTextSelection(6)
    insertPageBreak(editor)
    const { doc } = editor.state
    expect(doc.childCount).toBe(2)
    expect(doc.child(0).textContent).toBe('hello')
    expect(doc.child(1).textContent).toBe('')
    expect(doc.child(1).attrs.pageBreakBefore).toBe(true)
    editor.destroy()
  })

  it('backspace at the break start merges without losing text (break clears)', () => {
    const editor = makeEditor('hello world')
    editor.commands.setTextSelection(7)
    insertPageBreak(editor)
    // caret sits at the start of "world"; join backward like Backspace
    editor.commands.joinBackward()
    const { doc } = editor.state
    expect(doc.childCount).toBe(1)
    expect(doc.child(0).textContent).toBe('hello world')
    expect(doc.child(0).attrs.pageBreakBefore).toBeFalsy()
    editor.destroy()
  })
})
