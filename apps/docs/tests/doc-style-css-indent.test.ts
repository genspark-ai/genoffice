/**
 * Style-level first-line indent must not reach list items: .doc-li expresses
 * hanging via the ::before marker box (--li-hang chain), so a style text-indent
 * would double-shift the first line and inherit into the marker box, dragging
 * its glyphs out of the hanging area (prod-sas 037: "1." clipped to "l.").
 */
import { describe, expect, it } from 'vitest'
import type { ParsedDocFull, StyleDisplay, StyleInfo } from '@genoffice/docx-engine'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

function parsedWith(styleId: string, display: StyleDisplay): ParsedDocFull {
  const styles = new Map<string, StyleInfo>()
  styles.set(styleId, { styleId, name: styleId, type: 'paragraph', display } as StyleInfo)
  return { styles, docDefaults: {}, blocks: [] } as unknown as ParsedDocFull
}

describe('docStyleCss first-line indent', () => {
  it('emits style hanging as text-indent for non-list paragraphs and --style-li-hang for list items', () => {
    const css = docStyleCss(
      parsedWith('ListParagraph', { indentLeftTwips: 493, indentFirstLineTwips: -185 }),
    )
    expect(css).toContain(
      '.doc-page [data-style="ListParagraph"]:not(.doc-li) { text-indent:-9.3pt }',
    )
    expect(css).toContain('.doc-page .doc-li[data-style="ListParagraph"] { --style-li-hang:9.3pt }')
    // the shared style rule must not carry text-indent (it would hit .doc-li too)
    expect(css).not.toMatch(/\[data-style="ListParagraph"\] \{[^}]*text-indent/)
    // left indent keeps feeding the --li-left fallback chain
    expect(css).toContain(
      '.doc-page .doc-li[data-style="ListParagraph"] { --style-li-left:24.6pt }',
    )
  })

  it('keeps a positive first-line indent off list items without a hang fallback', () => {
    const css = docStyleCss(parsedWith('Body', { indentFirstLineTwips: 400 }))
    expect(css).toContain('.doc-page [data-style="Body"]:not(.doc-li) { text-indent:20.0pt }')
    expect(css).not.toContain('--style-li-hang')
  })
})
