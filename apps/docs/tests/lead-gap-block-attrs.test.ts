/**
 * blockAttrs glyph-shift inputs (styles.css --doc-lead-gap / --doc-lead-exact):
 * Word rests the descent on the single line's bottom and puts an exact baseline
 * at 0.8 h (probe 2026-09-30); the paragraph carries its strut face box, CJK
 * lines opt out of the gap, and a direct atLeast drops an inherited exact cap.
 */
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions, textboxSubExtensions } from '../src/renderer/editor/extensions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const para = (content: object[], attrs?: Record<string, unknown>) =>
  ({
    type: 'doc',
    content: [{ type: 'docParagraph', ...(attrs ? { attrs } : {}), content }],
  }) as never

const text = (t: string, attrs?: Record<string, unknown>) => ({
  type: 'text',
  text: t,
  ...(attrs ? { marks: [{ type: 'docTextStyle', attrs }] } : {}),
})

/** inline declarations as `name:value` (jsdom serialises them with a space) */
function styleOf(doc: never): string {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: doc,
  })
  const style = (editor.view.dom.querySelector('p') as HTMLElement).getAttribute('style') ?? ''
  editor.destroy()
  return style.replace(/: /g, ':').replace(/"/g, "'")
}

describe('blockAttrs glyph shift inputs', () => {
  it('a declared Latin face carries its Chromium box and skew', () => {
    const style = styleOf(para([text('Body', { font: 'Calibri', fontAscii: 'Calibri' })]))
    expect(style).toContain('--doc-font-box:1.0000')
    expect(style).toContain('--doc-font-skew:0.5000')
    expect(style).not.toContain('--doc-lead-gap')
  })

  it('an empty paragraph sized by its mark face carries that face box', () => {
    const style = styleOf(para([], { emptyRunFont: 'Times New Roman', emptyRunSize: 24 }))
    expect(style).toContain('--doc-font-box:1.1074')
  })

  it('an empty paragraph on a CJK mark face opts out of the gap and resets the box', () => {
    const style = styleOf(para([], { emptyRunFont: 'SimSun', emptyRunSize: 21 }))
    expect(style).toContain('--doc-lead-gap:0px')
    expect(style).toContain('--doc-font-box:initial')
    expect(style).not.toContain('--doc-font-box:1')
  })

  it('a declared face outside the table resets the inherited box', () => {
    const style = styleOf(para([text('Body', { font: 'DM Sans', fontAscii: 'DM Sans' })]))
    expect(style).toContain('--doc-font-box:initial')
    expect(style).toContain('--doc-font-skew:initial')
  })

  it('inheriting paragraphs leave the box to the style cascade', () => {
    expect(styleOf(para([text('Body')]))).not.toContain('--doc-font-box')
  })

  it('CJK paragraphs opt out of the single-line leading gap', () => {
    const style = styleOf(
      para([text('\u672c\u7814\u7a76 UT', { font: 'SimSun', fontAscii: 'Times New Roman' })]),
    )
    expect(style).toContain('--doc-lead-gap:0px')
    expect(style).toContain('--doc-font-box:1.1074')
  })

  it('exact lines read the 0.8 h shift, direct atLeast lines drop an inherited exact cap', () => {
    const exact = styleOf(para([text('Body')], { lineRule: 'exact', lineRawTwips: 360 }))
    expect(exact).toContain('--doc-lh-cap:18.0pt')
    expect(exact).toContain('--doc-lead-top:var(--doc-lead-exact, 0px)')
    const atLeast = styleOf(para([text('Body')], { lineRule: 'atLeast', lineRawTwips: 360 }))
    expect(atLeast).toContain('--doc-lh-cap:initial')
    expect(atLeast).toContain('+ var(--doc-lead-gap, 0px)')
  })

  it('textbox paragraphs follow the same exact/atLeast cap rules', () => {
    const render = (attrs: Record<string, unknown>): string => {
      const editor = new Editor({
        element: document.createElement('div'),
        extensions: textboxSubExtensions,
        content: para([text('Box')], attrs),
      })
      const el = editor.view.dom.querySelector('.doc-textbox-para') as HTMLElement
      const style = (el.getAttribute('style') ?? '').replace(/: /g, ':')
      editor.destroy()
      return style
    }
    expect(render({ lineRule: 'atLeast', lineRawTwips: 360 })).toContain('--doc-lh-cap:initial')
    expect(render({ lineRule: 'exact', lineRawTwips: 360 })).toContain('--doc-lh-cap:18.0pt')
  })
})
