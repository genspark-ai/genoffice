/**
 * Bold/italic runs of a per-cut size-adjusted alias (Merriweather GO: regular
 * 112% Georgia, bold 97.8%) lay a box of another size than the regular-cut
 * strut; Chromium's baseline-aligned union inflated 10.5pt label/value lines
 * to 13.5pt where Word paces 13.2 (prod corpus 20261020 sample 043).
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'

const run = (text: string, attrs: Record<string, unknown>, bold = false) => ({
  type: 'text',
  text,
  marks: [...(bold ? [{ type: 'bold' }] : []), { type: 'docTextStyle', attrs }],
})

const mount = (runs: unknown[], attrs: Record<string, unknown> = {}) =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', attrs, content: runs }],
    } as never,
  })

const para = (editor: Editor) => editor.view.dom.querySelector('p') as HTMLElement
const M = { fontAscii: 'Merriweather', sizeHalfPoints: 21 }

describe('per-cut size-adjusted alias runs', () => {
  it('flags the bold run and the one-size paragraph', () => {
    const editor = mount([run('Publication:', M), run('Mail & Guardian', M, true)])
    const p = para(editor)
    expect(p.classList.contains('doc-cut-adjusted')).toBe(true)
    const spans = [...p.querySelectorAll('span[data-doc-style]')] as HTMLElement[]
    expect(spans.map((s) => s.hasAttribute('data-cut-adj'))).toEqual([true, true])
    expect(spans[1].closest('strong')).not.toBeNull()
    editor.destroy()
  })

  it('leaves regular-only, mixed-size and exact-rule paragraphs alone', () => {
    const plain = mount([run('Publication:', M), run('Mail', M)])
    expect(para(plain).classList.contains('doc-cut-adjusted')).toBe(false)
    plain.destroy()
    const sizes = mount([run('Publication:', M), run('Mail', { ...M, sizeHalfPoints: 24 }, true)])
    expect(para(sizes).classList.contains('doc-cut-adjusted')).toBe(false)
    sizes.destroy()
    const exact = mount([run('Publication:', M), run('Mail', M, true)], {
      lineRule: 'exact',
      lineRawTwips: 280,
    })
    expect(para(exact).classList.contains('doc-cut-adjusted')).toBe(false)
    exact.destroy()
    const georgia = mount([
      run('a', { fontAscii: 'Georgia' }),
      run('b', { fontAscii: 'Georgia' }, true),
    ])
    const spans = [...para(georgia).querySelectorAll('span[data-doc-style]')]
    expect(spans.some((s) => s.hasAttribute('data-cut-adj'))).toBe(false)
    georgia.destroy()
  })
})
