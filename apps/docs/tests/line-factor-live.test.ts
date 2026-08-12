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

// jsdom's CSSOM drops the min() font-size, so assert the --doc-strut custom
// property it references (font-size:min(var(--doc-strut), 1em) rides along)
const fontSizeOf = (editor: Editor, index = 0): string => {
  const p = editor.view.dom.querySelectorAll('p')[index] as HTMLElement
  return p.style.getPropertyValue('--doc-strut')
}

const sizedParagraph = (text: string, sizeHalfPoints: number) =>
  ({
    type: 'doc',
    content: [
      {
        type: 'docParagraph',
        content: [
          { type: 'text', text, marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints } }] },
        ],
      },
    ],
  }) as never

describe('live line-height factor decorations', () => {
  it('typing CJK into a paragraph created empty switches its factor to the CJK var', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [{ type: 'docParagraph' }] } as never,
    })
    // Created empty: no per-block factor, inherits the document-level variable
    expect(factorOf(editor)).toBe('')

    editor.commands.insertContentAt(1, '中文正文')
    expect(factorOf(editor)).toBe('var(--doc-line-factor-cjk,1.7)')

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
    expect(factorOf(editor)).toBe('var(--doc-line-factor-cjk,1.7)')

    editor.commands.setTextSelection({ from: 1, to: 3 })
    editor.commands.insertContent('latin')
    expect(factorOf(editor)).toBe('var(--doc-line-factor-latin,1.2)')

    editor.destroy()
  })
})

describe('live strut font-size decorations', () => {
  it('changing the run font size updates the paragraph strut font-size', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: sizedParagraph('sized', 24),
    })
    expect(fontSizeOf(editor)).toBe('12pt')

    editor.commands.setTextSelection({ from: 1, to: 6 })
    editor.commands.setMark('docTextStyle', { sizeHalfPoints: 28 })
    expect(fontSizeOf(editor)).toBe('14pt')

    editor.destroy()
  })

  it('removes the strut font-size when part of the text loses its explicit size', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: sizedParagraph('sized', 24),
    })
    expect(fontSizeOf(editor)).toBe('12pt')

    editor.commands.setTextSelection({ from: 1, to: 3 })
    editor.commands.unsetMark('docTextStyle')
    expect(fontSizeOf(editor)).toBe('')

    editor.destroy()
  })

  it('deleting the only sized run removes the paragraph strut font-size', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: sizedParagraph('sized', 24),
    })
    expect(fontSizeOf(editor)).toBe('12pt')

    editor.commands.deleteRange({ from: 1, to: 6 })
    expect(fontSizeOf(editor)).toBe('')

    editor.destroy()
  })
})

describe('empty paragraph line size (emptyRunSize attr)', () => {
  it('renders a run-less paragraph at its paragraph-mark size', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', attrs: { emptyRunSize: 2 } }],
      } as never,
    })
    const p = editor.view.dom.querySelector('p') as HTMLElement
    expect(p.style.fontSize).toBe('1pt')
    editor.destroy()
  })
})
