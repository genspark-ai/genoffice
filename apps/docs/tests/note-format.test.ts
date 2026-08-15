import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { noteMarkText, toRoman } from '../src/renderer/note-format'
import { editorExtensions } from '../src/renderer/editor/extensions'

describe('note numbering', () => {
  it('endnotes number in lowercase roman, footnotes stay arabic', () => {
    expect(toRoman(1)).toBe('i')
    expect(toRoman(4)).toBe('iv')
    expect(toRoman(9)).toBe('ix')
    expect(toRoman(14)).toBe('xiv')
    expect(noteMarkText('endnote', 3)).toBe('iii')
    expect(noteMarkText('footnote', 3)).toBe('3')
  })

  it('in-text reference marks render roman for endnotes', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: '正文' },
              { type: 'docNoteRef', attrs: { kind: 'footnote', id: 'f1', num: 2 } },
              { type: 'docNoteRef', attrs: { kind: 'endnote', id: 'e1', num: 2 } },
            ],
          },
        ],
      } as never,
    })
    // bare number/roman text: the editor-only brackets are CSS ::before/::after
    // (hidden in print so they don't leak into exported PDF text)
    const sups = [...editor.view.dom.querySelectorAll('sup[data-note-ref]')]
    expect(sups.map((s) => s.textContent)).toEqual(['2', 'ii'])
    editor.destroy()
  })
})
