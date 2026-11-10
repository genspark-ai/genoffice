/**
 * Chromium treats U+2116 and other CJK-codepage symbols as ideographs when
 * distributing justification space; Word stretches only the spaces around them.
 * Non-CJK justified paragraphs wrap them in .doc-justify-symbol (text-justify:none).
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'

const symbols = (editor: Editor): string[] =>
  Array.from(editor.view.dom.querySelectorAll('.doc-justify-symbol'), (el) => el.textContent ?? '')

const paragraphDoc = (text: string, attrs?: Record<string, unknown>) =>
  ({
    type: 'doc',
    content: [
      {
        type: 'docParagraph',
        ...(attrs ? { attrs } : {}),
        content: [{ type: 'text', text }],
      },
    ],
  }) as never

const makeEditor = (content: unknown) =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: content as never,
  })

describe('justification symbol decorations', () => {
  it('wraps the symbols of a justified non-CJK paragraph', () => {
    const editor = makeEditor(paragraphDoc('за № 470-EL, 20 ℃', { align: 'justify' }))
    expect(symbols(editor)).toEqual(['№', '℃'])
    expect(editor.view.dom.textContent).toBe('за № 470-EL, 20 ℃')
    editor.destroy()
  })

  it('leaves left-aligned and CJK paragraphs alone', () => {
    const left = makeEditor(paragraphDoc('за № 470-EL'))
    expect(symbols(left)).toEqual([])
    left.destroy()
    const cjk = makeEditor(paragraphDoc('文件 № 470', { align: 'justify' }))
    expect(symbols(cjk)).toEqual([])
    cjk.destroy()
  })
})
