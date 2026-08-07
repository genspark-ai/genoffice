/**
 * The per-block --doc-line-factor must follow the live text, not the
 * text the block had when its DOM was first rendered (blockAttrs bakes it into
 * toDOM output, which ProseMirror reuses while typing).
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'

const factorOf = (editor: Editor, index = 0): string => {
  const p = editor.view.dom.querySelectorAll('p')[index] as HTMLElement
  return p.style.getPropertyValue('--doc-line-factor')
}

describe('live line-height factor decorations', () => {
  it('typing CJK into a paragraph created empty switches its factor to 1.3', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [{ type: 'docParagraph' }] } as never,
    })
    // Created empty: no per-block factor, inherits the document-level variable
    expect(factorOf(editor)).toBe('')

    editor.commands.insertContentAt(1, '中文正文')
    expect(factorOf(editor)).toBe('1.3')

    editor.destroy()
  })

  it('replacing CJK text with Western text switches back to the Latin factor', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', content: [{ type: 'text', text: '中文' }] }],
      } as never,
    })
    expect(factorOf(editor)).toBe('1.3')

    editor.commands.setTextSelection({ from: 1, to: 3 })
    editor.commands.insertContent('latin')
    expect(factorOf(editor)).toBe('var(--doc-line-factor-latin,1.2)')

    editor.destroy()
  })
})
