/**
 * Word's line strut follows run sizes: a paragraph whose runs all declare an
 * explicit size must carry the smallest of them on the block element (larger
 * runs lay their own line boxes), or the inherited body size inflates every
 * line box (table cells often run smaller than Normal). Any run inheriting the
 * body size keeps the inherited strut.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
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

// jsdom's CSSOM drops the var() font-size, so assert the --doc-strut custom
// property it references (font-size:var(--doc-strut) rides along)
const pFontSize = (editor: Editor, index = 0): string => {
  const p = editor.view.dom.querySelectorAll('p')[index] as HTMLElement
  return p.style.getPropertyValue('--doc-strut')
}

describe('paragraph strut font-size from run sizes', () => {
  it('all runs explicit: the paragraph carries the smallest run size', () => {
    const editor = editorWith([
      { type: 'docParagraph', content: [sized('small ', 17), sized('big', 20)] },
    ])
    expect(pFontSize(editor)).toBe('8.5pt')
    editor.destroy()
  })

  it('a bare inheriting text node (no element to restore its size) suppresses the strut override', () => {
    const editor = editorWith([
      { type: 'docParagraph', content: [sized('sized ', 17), { type: 'text', text: 'plain' }] },
    ])
    expect(pFontSize(editor)).toBe('')
    editor.destroy()
  })

  // the inheriting run has an element of its own, so .doc-strut-mixed hands it
  // --doc-base-fs back while the strut shrinks to the declared size (Word paces
  // a 9pt sub-line under inherited 11pt text at 9pt)
  it('an inheriting run inside a mark lets the strut take the smaller declared size', () => {
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
    expect(pFontSize(editor)).toBe('min(8.5pt,var(--doc-base-fs,8.5pt))')
    const p = editor.view.dom.querySelector('p') as HTMLElement
    expect(p.classList.contains('doc-strut-mixed')).toBe(true)
    editor.destroy()
  })

  it('all runs explicit: no mixed class', () => {
    const editor = editorWith([
      { type: 'docParagraph', content: [sized('small ', 17), sized('big', 20)] },
    ])
    const p = editor.view.dom.querySelector('p') as HTMLElement
    expect(p.classList.contains('doc-strut-mixed')).toBe(false)
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

// the restore rule sits at zero specificity (:where), so run elements sized by
// a stylesheet rule keep that size: a footnote mark stays 0.65em and an rtl
// run keeps the Cs size painted by .doc-cs; both are direct children without
// an inline font-size, so only specificity keeps --doc-base-fs off them
describe('.doc-strut-mixed restore rule yields to stylesheet sizes', () => {
  const css = readFileSync(
    join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
    'utf8',
  )

  const mountMixed = (extra: unknown) => {
    const editor = editorWith([
      {
        type: 'docParagraph',
        content: [
          sized('small ', 17),
          { type: 'text', text: 'plain', marks: [{ type: 'docTextStyle', attrs: { bold: true } }] },
          extra,
        ],
      },
    ])
    const p = editor.view.dom.querySelector('p') as HTMLElement
    expect(p.classList.contains('doc-strut-mixed')).toBe(true)
    return { editor, p }
  }

  it('the restore rule is wrapped in :where; the mark rule carries a class (.doc-cs: doc-cs-gate test)', () => {
    expect(css).toMatch(
      /^:where\(:is\(\.doc-page, \.page-hf\) \.doc-strut-mixed > :not\(\[style\*='font-size'\]\)\) \{\n\s*font-size: var\(--doc-base-fs, 11pt\);\n\}/m,
    )
    expect(css).toMatch(/^\.doc-note-ref \{[^}]*font-size: 0\.65em;/m)
  })

  it('a footnote mark is a direct child with no inline size', () => {
    const { editor, p } = mountMixed({
      type: 'docNoteRef',
      attrs: { kind: 'footnote', id: '1', num: 1 },
    })
    const mark = p.querySelector('sup.doc-note-ref') as HTMLElement
    expect(mark.parentElement).toBe(p)
    expect(mark.getAttribute('style') ?? '').not.toContain('font-size')
    editor.destroy()
  })

  it('an rtl run is a direct child with no inline size', () => {
    const { editor, p } = mountMixed({
      type: 'text',
      text: '\u0639\u0631\u0628\u064a',
      marks: [{ type: 'docTextStyle', attrs: { cs: true } }],
    })
    const cs = p.querySelector('span.doc-cs') as HTMLElement
    expect(cs.parentElement).toBe(p)
    expect(cs.getAttribute('style') ?? '').not.toContain('font-size')
    editor.destroy()
  })
})
