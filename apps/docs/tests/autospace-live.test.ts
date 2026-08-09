/**
 * CJK-Latin autospace pad widgets: Chromium's text-autospace delivers 1/8em of
 * Word's ~1/4em autoSpaceDE/DN gap; a .doc-autospace-pad widget at each boundary
 * supplies the rest and must follow live edits.
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'

const padCount = (editor: Editor): number =>
  editor.view.dom.querySelectorAll('.doc-autospace-pad').length

const paragraphDoc = (text: string, attrs?: Record<string, unknown>) =>
  ({
    type: 'doc',
    content: [
      {
        type: 'docParagraph',
        ...(attrs ? { attrs } : {}),
        content: text ? [{ type: 'text', text }] : [],
      },
    ],
  }) as never

const makeEditor = (content: unknown) =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: content as never,
  })

describe('autospace pad decorations', () => {
  it('pads each CJK-Latin/digit boundary, skipping space/punctuation seams', () => {
    const editor = makeEditor(paragraphDoc('テスト17.0km ペン'))
    expect(padCount(editor)).toBe(1)
    editor.destroy()
  })

  it('covers hangul boundaries', () => {
    const editor = makeEditor(paragraphDoc('한글A와 B'))
    expect(padCount(editor)).toBe(2)
    editor.destroy()
  })

  it('adds no pads when the paragraph turns autoSpace off', () => {
    const editor = makeEditor(paragraphDoc('ペン12', { autoSpace: false }))
    expect(padCount(editor)).toBe(0)
    editor.destroy()
  })

  it('pads across styled-run boundaries', () => {
    const editor = makeEditor({
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'ペン' },
            {
              type: 'text',
              text: '12',
              marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 24 } }],
            },
          ],
        },
      ],
    })
    expect(padCount(editor)).toBe(1)
    editor.destroy()
  })

  it('follows live edits', () => {
    const editor = makeEditor(paragraphDoc('ペン'))
    expect(padCount(editor)).toBe(0)

    editor.commands.insertContentAt(3, '12')
    expect(padCount(editor)).toBe(1)

    editor.commands.deleteRange({ from: 3, to: 5 })
    expect(padCount(editor)).toBe(0)
    editor.destroy()
  })
})
