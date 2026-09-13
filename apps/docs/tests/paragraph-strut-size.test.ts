/**
 * Word's line strut follows run sizes: a paragraph whose runs all declare an
 * explicit size must carry that size (max of runs) on the block element, or the
 * inherited body size inflates every line box (table cells often run smaller
 * than Normal). Any run inheriting the body size keeps the inherited strut.
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'

const sized = (text: string, sizeHalfPoints: number) => ({
  type: 'text',
  text,
  marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints } }],
})

const editorWith = (content: unknown[]): Editor =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content } as never,
  })

// jsdom's CSSOM drops the min() font-size, so assert the --doc-strut custom
// property it references (font-size:min(var(--doc-strut), 1em) rides along)
const pFontSize = (editor: Editor, index = 0): string => {
  const p = editor.view.dom.querySelectorAll('p')[index] as HTMLElement
  return p.style.getPropertyValue('--doc-strut')
}

describe('paragraph strut font-size from run sizes', () => {
  it('all runs explicit: the paragraph carries the max run size', () => {
    const editor = editorWith([
      { type: 'docParagraph', content: [sized('small ', 17), sized('big', 20)] },
    ])
    expect(pFontSize(editor)).toBe('10pt')
    editor.destroy()
  })

  it('a run inheriting the body size suppresses the strut override', () => {
    const editor = editorWith([
      { type: 'docParagraph', content: [sized('sized ', 17), { type: 'text', text: 'plain' }] },
    ])
    expect(pFontSize(editor)).toBe('')
    editor.destroy()
  })

  it('a docTextStyle mark without sizeHalfPoints counts as inherited', () => {
    const editor = editorWith([
      {
        type: 'docParagraph',
        content: [
          sized('sized ', 17),
          {
            type: 'text',
            text: 'red',
            marks: [{ type: 'docTextStyle', attrs: { color: 'FF0000' } }],
          },
        ],
      },
    ])
    expect(pFontSize(editor)).toBe('')
    editor.destroy()
  })

  it('an empty paragraph emits no font-size', () => {
    const editor = editorWith([{ type: 'docParagraph' }])
    expect(pFontSize(editor)).toBe('')
    editor.destroy()
  })

  // Word probe 2026-09-11: space runs never size a line (a 2pt, 4pt or 17pt
  // space paragraph measures exactly like an 11pt text line)
  it('space-only runs do not size the strut', () => {
    const editor = editorWith([
      { type: 'docParagraph', content: [sized('text', 20), sized('   ', 40)] },
      { type: 'docParagraph', content: [sized(' ', 8)] },
    ])
    expect(pFontSize(editor, 0)).toBe('10pt')
    expect(pFontSize(editor, 1)).toBe('')
    editor.destroy()
  })

  it('a space-only paragraph is sized by the mark like an empty one', () => {
    const editor = editorWith([
      { type: 'docParagraph', attrs: { emptyRunSize: 8 }, content: [sized(' ', 8)] },
      { type: 'docParagraph', content: [sized(' ', 8)] },
    ])
    const [marked, plain] = Array.from(editor.view.dom.querySelectorAll('p')) as HTMLElement[]
    expect(marked.style.fontSize).toBe('4pt')
    expect(plain.style.fontSize).toBe('')
    for (const p of [marked, plain]) {
      expect(p.style.getPropertyValue('--doc-strut')).toBe('')
      expect(p.style.getPropertyValue('--doc-line-factor')).toBe('var(--doc-line-factor-latin,1.2)')
    }
    editor.destroy()
  })
})
