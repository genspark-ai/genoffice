import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import {
  formatNoteNumber,
  noteMarkText,
  setNoteNumFmts,
  toRoman,
} from '../src/renderer/note-format'
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

  it('w:numFmt overrides the per-kind default (decimal endnotes, roman footnotes, chicago marks)', () => {
    expect(formatNoteNumber(undefined, 4)).toBe('4')
    expect(formatNoteNumber('upperRoman', 4)).toBe('IV')
    expect(formatNoteNumber('lowerLetter', 27)).toBe('aa')
    expect(formatNoteNumber('upperLetter', 2)).toBe('B')
    expect(formatNoteNumber('chicago', 5)).toBe('**')
    expect(formatNoteNumber('decimalEnclosedCircle', 3)).toBe('3')
    setNoteNumFmts({ endnote: { numFmt: 'decimal' }, footnote: { numFmt: 'lowerRoman' } })
    try {
      expect(noteMarkText('endnote', 1)).toBe('1')
      expect(noteMarkText('footnote', 2)).toBe('ii')
      const editor = new Editor({
        element: document.createElement('div'),
        extensions: editorExtensions,
        content: {
          type: 'doc',
          content: [
            {
              type: 'docParagraph',
              content: [{ type: 'docNoteRef', attrs: { kind: 'endnote', id: 'e1', num: 3 } }],
            },
          ],
        } as never,
      })
      expect(editor.view.dom.querySelector('sup[data-note-ref]')?.textContent).toBe('3')
      editor.destroy()
    } finally {
      setNoteNumFmts({})
    }
    expect(noteMarkText('endnote', 1)).toBe('i')
  })

  it('caps repeat expansion for hostile note numbers', () => {
    expect(formatNoteNumber('lowerLetter', 1)).toBe('a')
    expect(formatNoteNumber('lowerLetter', 27)).toBe('aa')
    expect(formatNoteNumber('lowerLetter', 1e9)).toBe('1000000000')
    expect(formatNoteNumber('chicago', 5)).toBe('**')
    expect(formatNoteNumber('chicago', 1e9)).toBe('1000000000')
    expect(formatNoteNumber('lowerLetter', 1e9).length).toBeLessThan(100)
  })
})
