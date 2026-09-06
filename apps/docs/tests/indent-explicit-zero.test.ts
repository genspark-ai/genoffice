import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { ParsedDocFull, StyleDisplay, StyleInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

/** Word merges w:ind per attribute and an explicit 0 is a value that cancels the
 *  inherited indent: it must reach the DOM as `0pt`, never be dropped as "unset". */

function parsedWith(entries: Array<[string, StyleDisplay, boolean?]>): ParsedDocFull {
  const styles = new Map<string, StyleInfo>()
  for (const [styleId, display, isDefault] of entries) {
    styles.set(styleId, {
      styleId,
      name: styleId,
      type: 'paragraph',
      display,
      isDefault,
    } as StyleInfo)
  }
  return { styles, docDefaults: {}, blocks: [] } as unknown as ParsedDocFull
}

function renderPara(attrs: Record<string, unknown>): string {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', attrs, content: [{ type: 'text', text: 'body' }] }],
    },
  })
  const html = editor.getHTML()
  editor.destroy()
  return html
}

describe('explicit zero indents on paragraphs', () => {
  it('emits text-indent/margins of 0pt for direct zeros', () => {
    const html = renderPara({ indentLeft: 0, indentRight: 0, indentFirstLine: 0 })
    expect(html).toContain('margin-inline-start: 0pt')
    expect(html).toContain('margin-inline-end: 0pt')
    expect(html).toContain('text-indent: 0pt')
  })

  it('emits nothing for unset indents', () => {
    const html = renderPara({})
    expect(html).not.toContain('margin-inline-start')
    expect(html).not.toContain('text-indent')
  })
})

describe('docStyleCss explicit zero indents', () => {
  it("applies Normal's left/right indent to unstyled paragraphs and lets a child style cancel it", () => {
    const css = docStyleCss(
      parsedWith([
        ['Normal', { indentLeftTwips: 454, indentRightTwips: 120 }, true],
        ['HDR', { indentLeftTwips: 0, indentRightTwips: 0, indentFirstLineTwips: 0 }],
      ]),
    )
    expect(css).toContain(
      '.doc-page :is(p, h1, h2, h3, h4, h5, h6, .doc-protected-field):not([data-style]) { margin-inline-start:22.7pt;margin-inline-end:6.0pt }',
    )
    expect(css).toContain(
      '.doc-page [data-style="HDR"]:not(.doc-li, .doc-li-stray) { margin-inline-start:0.0pt }',
    )
    expect(css).toContain(
      '.doc-page [data-style="HDR"]:not(.doc-li, .doc-li-stray) { text-indent:0.0pt }',
    )
    expect(css).toMatch(/\[data-style="HDR"\] \{[^}]*margin-inline-end:0\.0pt/)
    expect(css).not.toContain('--style-li-hang')
  })

  it('emits no default indent rule when Normal has none', () => {
    const css = docStyleCss(parsedWith([['Normal', { sizeHalfPoints: 22 }, true]]))
    expect(css).not.toContain(':not([data-style]) { margin-inline-start')
  })
})
