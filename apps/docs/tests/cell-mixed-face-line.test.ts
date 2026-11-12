/**
 * A cell paragraph mixing declared faces of different line factors: the
 * paragraph carries .doc-mixed-face and every run span its own --doc-run-lh,
 * so styles.css can give the shorter face its own line box (Word lays a mixed
 * line at max ascent + max descent; a Times run centred in a Sakkal Majalla
 * line-height lifted rows by 0.15em).
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'

const run = (text: string, font: string, sizeHalfPoints = 24) => ({
  type: 'text',
  text,
  marks: [{ type: 'docTextStyle', attrs: { fontAscii: font, csFont: font, sizeHalfPoints } }],
})

const paragraph = (runs: unknown[]) => ({ type: 'docParagraph', content: runs })

const mount = (content: unknown) =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [content] } as never,
  })

describe('mixed-face cell paragraphs', () => {
  it('flags a paragraph whose declared faces differ in line factor and tags each run', () => {
    const editor = mount(
      paragraph([run('الانحراف ', 'Sakkal Majalla'), run('σ', 'Times New Roman')]),
    )
    const p = editor.view.dom.querySelector('p') as HTMLElement
    expect(p.classList.contains('doc-mixed-face')).toBe(true)
    expect(p.style.getPropertyValue('--doc-line-factor')).toBe('1.3965')
    const spans = [...p.querySelectorAll('span[style*="--doc-run-lh"]')] as HTMLElement[]
    expect(spans.map((s) => s.style.getPropertyValue('--doc-run-lh'))).toEqual(['1.3965', '1.15'])
    editor.destroy()
  })

  it('leaves a single-face paragraph and an inheriting one alone', () => {
    const same = mount(paragraph([run('abc ', 'Sakkal Majalla'), run('def', 'Sakkal Majalla')]))
    expect(same.view.dom.querySelector('p')!.classList.contains('doc-mixed-face')).toBe(false)
    same.destroy()
    const inherit = mount(paragraph([{ type: 'text', text: 'abc ' }, run('σ', 'Times New Roman')]))
    expect(inherit.view.dom.querySelector('p')!.classList.contains('doc-mixed-face')).toBe(false)
    inherit.destroy()
  })

  it('leaves mixed run sizes alone (a larger shorter-face box could outgrow the strut line)', () => {
    const sizes = mount(
      paragraph([run('abc ', 'Sakkal Majalla'), run('\u03c3', 'Times New Roman', 28)]),
    )
    expect(sizes.view.dom.querySelector('p')!.classList.contains('doc-mixed-face')).toBe(false)
    sizes.destroy()
  })
})
