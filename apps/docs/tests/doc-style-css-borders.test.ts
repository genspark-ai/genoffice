import { describe, expect, it } from 'vitest'
import type { ParsedDocFull, StyleDisplay, StyleInfo } from '@genoffice/docx-engine'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

function parsedWith(styleId: string, display: StyleDisplay): ParsedDocFull {
  const styles = new Map<string, StyleInfo>()
  styles.set('Normal', {
    styleId: 'Normal',
    name: 'Normal',
    type: 'paragraph',
    isDefault: true,
  } as StyleInfo)
  styles.set(styleId, {
    styleId,
    name: styleId,
    type: 'paragraph',
    basedOn: 'Normal',
    display,
  } as StyleInfo)
  return { styles, docDefaults: {}, blocks: [] } as unknown as ParsedDocFull
}

describe('docStyleCss paragraph borders', () => {
  it('draws style-level w:pBdr sides with the direct-border look', () => {
    const css = docStyleCss(
      parsedWith('HDR', { borderSides: { b: { color: '1F4E79', szPt: 6 }, t: {} } }),
    )
    expect(css).toContain(
      '[data-style="HDR"] { border-top:1px solid #444;border-bottom:8px solid #1F4E79;padding-top:0;padding-bottom:0 }',
    )
  })

  it('pads only the drawn sides, by w:space, as longhands (a direct side on another edge layers per side)', () => {
    const css = docStyleCss(
      parsedWith('Title', { borderSides: { b: { color: '4F81BD', szPt: 1, spacePt: 4 } } }),
    )
    expect(css).toContain(
      '[data-style="Title"] { border-bottom:1px solid #4F81BD;padding-bottom:4pt }',
    )
    const boxed = docStyleCss(
      parsedWith('Box', { borderSides: { t: { spacePt: 1 }, l: { spacePt: 4 }, b: {}, r: {} } }),
    )
    expect(boxed).toContain('padding-top:1pt;padding-right:4px;padding-bottom:0;padding-left:4px }')
  })

  it('emits nothing for a side reset to none', () => {
    const css = docStyleCss(parsedWith('Off', { borderSides: { b: null } }))
    expect(css).not.toContain('border-')
    expect(css).not.toContain('padding:')
  })
})
