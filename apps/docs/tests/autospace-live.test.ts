/**
 * CJK-Latin autospace pads: Chromium's text-autospace delivers 1/8em of Word's
 * ~1/4em autoSpaceDE/DN gap; the character after each boundary is wrapped in a
 * .doc-autospace-pad span whose start margin supplies the rest, and the pads
 * must follow live edits.
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import type { Run } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { runSpanSpecs } from '../src/renderer/editor/protected-render'
import { codePointLengthAt } from '../src/renderer/line-metrics'

const pads = (editor: Editor): string[] =>
  Array.from(editor.view.dom.querySelectorAll('.doc-autospace-pad'), (el) => el.textContent ?? '')
const padCount = (editor: Editor): number => pads(editor).length

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

  // an empty pad span whose margin overflows the line end makes Chromium rewind
  // to the previous pad instead of breaking inside the CJK run (prod 003)
  it('wraps the character after the boundary instead of inserting an empty span', () => {
    const editor = makeEditor(paragraphDoc('累计产生357次表单提交'))
    expect(pads(editor)).toEqual(['3', '次'])
    expect(editor.view.dom.textContent).toBe('累计产生357次表单提交')
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

describe('static-render autospace pads', () => {
  const run = (text: string): Run => ({ text }) as Run

  it('wraps the character after each boundary', () => {
    expect(runSpanSpecs(run('累计产生357次'))).toEqual([
      [
        'span',
        {},
        '累计产生',
        ['span', { class: 'doc-autospace-pad' }, '3'],
        '57',
        ['span', { class: 'doc-autospace-pad' }, '次'],
      ],
    ])
  })

  it('leadPad wraps the first character of a run that continues a seam', () => {
    expect(runSpanSpecs(run('286名'), undefined, true)).toEqual([
      [
        'span',
        {},
        ['span', { class: 'doc-autospace-pad' }, '2'],
        '86',
        ['span', { class: 'doc-autospace-pad' }, '名'],
      ],
    ])
    expect(runSpanSpecs(run('286名'), false, true)).toEqual([
      ['span', { style: 'text-autospace:no-autospace' }, '286名'],
    ])
  })

  it('measures astral code points as two UTF-16 units', () => {
    expect(codePointLengthAt('A\u{20000}', 1)).toBe(2)
    expect(codePointLengthAt('AB', 1)).toBe(1)
  })
})
