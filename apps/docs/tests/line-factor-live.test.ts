/**
 * The per-block --doc-line-factor must follow the live text, not the
 * text the block had when its DOM was first rendered (blockAttrs bakes it into
 * toDOM output, which ProseMirror reuses while typing).
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'

const factorOf = (editor: Editor, index = 0): string => {
  const p = editor.view.dom.querySelectorAll('p')[index] as HTMLElement
  return p.style.getPropertyValue('--doc-line-factor')
}

// jsdom's CSSOM drops the min() font-size, so assert the --doc-strut custom
// property it references (font-size:min(var(--doc-strut), 1em) rides along)
const fontSizeOf = (editor: Editor, index = 0): string => {
  const p = editor.view.dom.querySelectorAll('p')[index] as HTMLElement
  return p.style.getPropertyValue('--doc-strut')
}

const sizedParagraph = (text: string, sizeHalfPoints: number) =>
  ({
    type: 'doc',
    content: [
      {
        type: 'docParagraph',
        content: [
          { type: 'text', text, marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints } }] },
        ],
      },
    ],
  }) as never

describe('live line-height factor decorations', () => {
  it('typing CJK into a paragraph created empty switches its factor to the CJK var', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [{ type: 'docParagraph' }] } as never,
    })
    // Created empty: no per-block factor, inherits the document-level variable
    expect(factorOf(editor)).toBe('')

    editor.commands.insertContentAt(1, '中文正文')
    expect(factorOf(editor)).toBe('var(--doc-line-factor-cjk,1.7)')

    editor.destroy()
  })

  it('replacing CJK text with Western text switches back to the Latin factor', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', content: [{ type: 'text', text: '中文' }] }],
      } as never,
    })
    expect(factorOf(editor)).toBe('var(--doc-line-factor-cjk,1.7)')

    editor.commands.setTextSelection({ from: 1, to: 3 })
    editor.commands.insertContent('latin')
    expect(factorOf(editor)).toBe('var(--doc-line-factor-latin,1.2)')

    editor.destroy()
  })
})

describe('live strut font-size decorations', () => {
  it('changing the run font size updates the paragraph strut font-size', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: sizedParagraph('sized', 24),
    })
    expect(fontSizeOf(editor)).toBe('12pt')

    editor.commands.setTextSelection({ from: 1, to: 6 })
    editor.commands.setMark('docTextStyle', { sizeHalfPoints: 28 })
    expect(fontSizeOf(editor)).toBe('14pt')

    editor.destroy()
  })

  it('removes the strut font-size when part of the text loses its explicit size', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: sizedParagraph('sized', 24),
    })
    expect(fontSizeOf(editor)).toBe('12pt')

    editor.commands.setTextSelection({ from: 1, to: 3 })
    editor.commands.unsetMark('docTextStyle')
    expect(fontSizeOf(editor)).toBe('')

    editor.destroy()
  })

  it('deleting the only sized run removes the paragraph strut font-size', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: sizedParagraph('sized', 24),
    })
    expect(fontSizeOf(editor)).toBe('12pt')

    editor.commands.deleteRange({ from: 1, to: 6 })
    expect(fontSizeOf(editor)).toBe('')

    editor.destroy()
  })
})

describe('declared run fonts drive the Latin factor and the paragraph face', () => {
  const fontParagraph = (runs: Array<{ text: string; font?: string }>) =>
    ({
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: runs.map((r) => ({
            type: 'text',
            text: r.text,
            ...(r.font
              ? { marks: [{ type: 'docTextStyle', attrs: { font: r.font, fontAscii: r.font } }] }
              : {}),
          })),
        },
      ],
    }) as never

  it('all runs declaring Calibri: run-font factor and the run face on the paragraph', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: fontParagraph([{ text: 'declared', font: 'Calibri' }]),
    })
    expect(factorOf(editor)).toBe('1.22')
    const p = editor.view.dom.querySelector('p') as HTMLElement
    expect(p.style.fontFamily).toContain('Calibri')
    editor.destroy()
  })

  it('a run inheriting the body font keeps the doc var via max() and no paragraph face', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: fontParagraph([{ text: 'declared ', font: 'Calibri' }, { text: 'inherited' }]),
    })
    expect(factorOf(editor)).toBe('max(var(--doc-line-factor-latin,1.2), 1.22)')
    const p = editor.view.dom.querySelector('p') as HTMLElement
    expect(p.style.fontFamily).toBe('')
    editor.destroy()
  })

  it('CJK runs declaring a JP-variant Noto face take the JA substitution factor', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: fontParagraph([{ text: '日本語本文', font: 'Noto Sans CJK JP' }]),
    })
    expect(factorOf(editor)).toBe('1.3029')
    editor.destroy()
  })

  it('unsetting the font mark live drops the factor override and the paragraph face', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: fontParagraph([{ text: 'declared', font: 'Calibri' }]),
    })
    expect(factorOf(editor)).toBe('1.22')

    editor.commands.setTextSelection({ from: 1, to: 9 })
    editor.commands.unsetMark('docTextStyle')
    expect(factorOf(editor)).toBe('var(--doc-line-factor-latin,1.2)')
    const p = editor.view.dom.querySelector('p') as HTMLElement
    expect(p.style.fontFamily).toBe('')
    editor.destroy()
  })
})

describe('SimSun-substitution ・/〜 line lift decorations', () => {
  const gapDoc = (text: string, font: string, attrs?: Record<string, unknown>) =>
    ({
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          ...(attrs ? { attrs } : {}),
          content: [{ type: 'text', text, marks: [{ type: 'docTextStyle', attrs: { font } }] }],
        },
      ],
    }) as never

  const liftSpans = (editor: Editor): HTMLElement[] =>
    Array.from(editor.view.dom.querySelectorAll('span')).filter(
      (s) => (s as HTMLElement).style.lineHeight !== '',
    ) as HTMLElement[]

  it('wraps ・ and 〜 in a 1.7143 line-height span under a SimSun-substituted face', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: gapDoc('項目・内容〜まで', 'Noto Sans CJK JP'),
    })
    const spans = liftSpans(editor)
    expect(spans.map((s) => s.textContent)).toEqual(['・', '〜'])
    expect(spans[0].style.lineHeight).toBe(
      'round(up, calc(1.7143 * 1em - var(--doc-grid-pitch,0.0001px) * 0.001), var(--doc-grid-pitch,0.0001px))',
    )
    editor.destroy()
  })

  it('scales the lift by an explicit line-spacing multiple', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: gapDoc('項目・内容', 'Noto Sans JP', { lineRule: 'auto', lineSpacing: 2 }),
    })
    expect(liftSpans(editor)[0].style.lineHeight).toBe(
      'calc(round(up, calc(1.7143 * 1em - var(--doc-grid-pitch,0.0001px) * 0.001), var(--doc-grid-pitch,0.0001px)) * 2)',
    )
    editor.destroy()
  })

  it('does not lift under non-SimSun faces, without ・/〜, or with an exact rule', () => {
    for (const content of [
      gapDoc('項目・内容', 'Microsoft YaHei'),
      gapDoc('바탕・본문', 'Batang'),
      gapDoc('項目と内容', 'Noto Sans CJK JP'),
      gapDoc('項目・内容', 'Noto Sans CJK JP', { lineRule: 'exact', lineRawTwips: 240 }),
    ]) {
      const editor = new Editor({
        element: document.createElement('div'),
        extensions: editorExtensions,
        content,
      })
      expect(liftSpans(editor)).toEqual([])
      editor.destroy()
    }
  })
})

describe('empty paragraph line size (emptyRunSize attr)', () => {
  it('renders a run-less paragraph at its paragraph-mark size', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', attrs: { emptyRunSize: 2 } }],
      } as never,
    })
    const p = editor.view.dom.querySelector('p') as HTMLElement
    expect(p.style.fontSize).toBe('1pt')
    editor.destroy()
  })
})
