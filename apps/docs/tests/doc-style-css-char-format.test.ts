import { describe, expect, it } from 'vitest'
import type {
  ParsedDocFull,
  StyleDisplay,
  StyleInfo,
  TableStyleDisplay,
} from '@genoffice/docx-engine'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

function parsedWith(info: Partial<StyleInfo> & { styleId: string }): ParsedDocFull {
  const styles = new Map<string, StyleInfo>()
  styles.set(info.styleId, { name: info.styleId, type: 'paragraph', ...info } as StyleInfo)
  return { styles, docDefaults: {}, blocks: [] } as unknown as ParsedDocFull
}

describe('docStyleCss character formatting from styles.xml', () => {
  it('strokes text with the style w14:textOutline, alpha as rgba', () => {
    const display: StyleDisplay = { textOutline: { color: 'FF0000', widthPt: 0.75, alpha: 0.4 } }
    const css = docStyleCss(parsedWith({ styleId: 'IntenseEmphasis', type: 'character', display }))
    expect(css).toContain(
      '[data-style="IntenseEmphasis"] { -webkit-text-stroke:0.75pt rgba(255,0,0,0.4) }',
    )
  })

  it('paints character-style shading and flips auto ink on a dark fill', () => {
    const css = docStyleCss(
      parsedWith({ styleId: 'ChangeType', type: 'character', display: { shading: '000000' } }),
    )
    expect(css).toMatch(
      /\[data-style="ChangeType"\] \{[^}]*background-color:#000000;--docs-paper-ink:#fff;color:var\(--docs-paper-ink\)/,
    )
    // the dark page remaps the fill light, so the auto ink goes paper-dark like run-level data-ink
    expect(css).toMatch(
      /\.page-dark \.doc-page \[data-style="ChangeType"\][^{]*\{ background:#[0-9a-f]{6};--docs-paper-ink:#1e1e1e \}/,
    )
  })

  it('keeps an explicit style colour over the auto ink and skips paragraph styles', () => {
    const explicit = docStyleCss(
      parsedWith({
        styleId: 'Tag',
        type: 'character',
        display: { shading: '000000', color: 'FFFF00' },
      }),
    )
    expect(explicit).not.toContain('color:#FFFFFF')
    const para = docStyleCss(
      parsedWith({ styleId: 'Boxed', type: 'paragraph', display: { shading: '000000' } }),
    )
    expect(para).not.toContain('background-color:#000000')
  })

  // TOC \h entries carry the (localized id) Hyperlink style;
  // Word's print layout keeps them in the TOC paragraph's ink, face and size intact
  it('hides the Hyperlink style colour and underline inside TOC entries only', () => {
    const styles = new Map<string, StyleInfo>()
    styles.set('a3', {
      styleId: 'a3',
      name: 'Hyperlink',
      type: 'character',
      display: { color: '0000FF', underline: true, fontAscii: 'Verdana' },
    } as StyleInfo)
    styles.set('20', { styleId: '20', name: 'toc 2', type: 'paragraph', display: {} } as StyleInfo)
    styles.set('Body', {
      styleId: 'Body',
      name: 'Body',
      type: 'paragraph',
      display: {},
    } as StyleInfo)
    const css = docStyleCss({ styles, docDefaults: {}, blocks: [] } as unknown as ParsedDocFull)
    expect(css).toMatch(/\[data-style="a3"\] \{[^}]*color:#0000FF;[^}]*text-decoration:underline/)
    expect(css).toContain(
      '.doc-page :is(.doc-toc-line,[data-style="20"]) [data-style="a3"] { color:inherit;text-decoration:none }',
    )
    expect(css).toContain(
      '.page-dark .doc-page :is(.doc-toc-line,[data-style="20"]) [data-style="a3"] { color:inherit }',
    )
    expect(css).toContain('.doc-link:has([data-style="a3"]) { color:inherit;text-decoration:none }')
    expect(css).not.toContain('[data-style="Body"]) [data-style="a3"]')
  })

  it('applies header-row caps and font to the first row and the leading tblHeader rows', () => {
    const tableDisplay: TableStyleDisplay = {
      firstRow: { caps: 'all', color: '4F81BD', fontAscii: 'Cambria', charSpacingTwips: 20 },
    }
    const css = docStyleCss(parsedWith({ styleId: 'Calendar2', type: 'table', tableDisplay }))
    const headerTr = 'tr[data-repeat-header="1"]:not(tr:not([data-repeat-header="1"]) ~ tr)'
    const rule = css.split('\n').find((line) => line.includes('tr:first-child td'))
    expect(rule).toContain(`table[data-tbl-style="Calendar2"] ${headerTr} td`)
    expect(rule).toContain('text-transform:uppercase')
    expect(rule).toContain('letter-spacing:1pt')
    expect(rule).toMatch(/font-family:[^;]*Cambria/)
  })
})
