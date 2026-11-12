/**
 * Complex-script runs and the paragraph strut. Word lays an Arabic-only
 * paragraph at its w:cs face's factor (Arial 1.15, probe 2026-10-13) and a
 * mixed Latin + Arabic line at the tallest face (Calibri 1.22). Chromium
 * centres each face in the inherited line-height, so a Calibri strut under
 * Arial spans paced 1.25em and the mixed line 1.31em.
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import {
  DocDocument,
  DocText,
  TextboxParagraph,
  editorExtensions,
} from '../src/renderer/editor/extensions'
import { TextStyleMark } from '../src/renderer/editor/marks'

const AR = 'صف خلق خود'

const run = (
  text: string,
  attrs: Record<string, unknown>,
): { type: string; text: string; marks: unknown[] } => ({
  type: 'text',
  text,
  marks: [{ type: 'docTextStyle', attrs }],
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

describe('complex-script paragraph strut', () => {
  it('takes the cs face as strut when every run is complex-script text', () => {
    const editor = mount([
      run(AR, { fontAscii: 'Calibri', csFont: 'Arial', sizeHalfPoints: 24 }),
      run(' كمثل', { fontAscii: 'Calibri', csFont: 'Arial', sizeHalfPoints: 24 }),
    ])
    const p = para(editor)
    expect(p.style.fontFamily).toMatch(/^"?Arial/)
    expect(p.style.getPropertyValue('--doc-font-box')).toBe('1.1172')
    expect(p.style.getPropertyValue('--doc-line-factor')).toBe('1.15')
    expect(p.classList.contains('doc-cs-mixed')).toBe(false)
    editor.destroy()
  })

  it('uses the inherited cs face even without a declared ascii face', () => {
    const editor = mount([run(AR, { csFont: 'Arial' })])
    const p = para(editor)
    expect(p.style.fontFamily).toMatch(/^"?Arial/)
    editor.destroy()
  })

  it('keeps the declared ascii strut when the runs differ in cs face', () => {
    const editor = mount([
      run(AR, { fontAscii: 'Calibri', csFont: 'Arial' }),
      run(' ك', { fontAscii: 'Calibri', csFont: 'Sakkal Majalla' }),
    ])
    expect(para(editor).style.fontFamily).toMatch(/^"?Calibri/)
    editor.destroy()
  })
})

describe('mixed Latin + complex-script paragraphs', () => {
  it('collapses cs spans into the strut line at one run size', () => {
    const editor = mount([
      run(AR, { fontAscii: 'Calibri', csFont: 'Arial', sizeHalfPoints: 24 }),
      run(' Latin', { fontAscii: 'Calibri', sizeHalfPoints: 24 }),
    ])
    const p = para(editor)
    expect(p.classList.contains('doc-cs-mixed')).toBe(true)
    expect(p.style.fontFamily).toMatch(/^"?Calibri/)
    expect(p.style.getPropertyValue('--doc-line-factor')).toBe('1.22')
    const spans = [...p.querySelectorAll('span[data-doc-style]')] as HTMLElement[]
    expect(spans.map((s) => s.hasAttribute('data-cs'))).toEqual([true, false])
    editor.destroy()
  })

  it('leaves mixed run sizes and fixed line rules alone', () => {
    const sizes = mount([
      run(AR, { fontAscii: 'Calibri', csFont: 'Arial', sizeHalfPoints: 24 }),
      run(' Latin', { fontAscii: 'Calibri', sizeHalfPoints: 28 }),
    ])
    expect(para(sizes).classList.contains('doc-cs-mixed')).toBe(false)
    sizes.destroy()
    const exact = mount(
      [run(AR, { fontAscii: 'Calibri', csFont: 'Arial' }), run(' Latin', { fontAscii: 'Calibri' })],
      { lineRule: 'exact', lineRawTwips: 280 },
    )
    expect(para(exact).classList.contains('doc-cs-mixed')).toBe(false)
    exact.destroy()
  })

  it('ignores a Latin-only paragraph whose runs merely carry a cs face', () => {
    const editor = mount([run('Latin only', { fontAscii: 'Calibri', csFont: 'Arial' })])
    expect(para(editor).classList.contains('doc-cs-mixed')).toBe(false)
    editor.destroy()
  })
})

describe('textbox paragraphs', () => {
  it('carries doc-cs-mixed on the rendered class attribute', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: [DocDocument, DocText, TextboxParagraph, TextStyleMark],
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              run(AR, { fontAscii: 'Calibri', csFont: 'Arial' }),
              run(' Latin', { fontAscii: 'Calibri' }),
            ],
          },
        ],
      } as never,
    })
    const p = editor.view.dom.querySelector('.doc-textbox-para') as HTMLElement | null
    expect(p?.classList.contains('doc-cs-mixed')).toBe(true)
    editor.destroy()
  })
})
