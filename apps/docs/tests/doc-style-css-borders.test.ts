import { describe, expect, it } from 'vitest'
import type { ParsedDocFull, StyleDisplay, StyleInfo } from '@genoffice/docx-engine'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

function parsedWith(styleId: string, display: StyleDisplay): ParsedDocFull {
  const styles = new Map<string, StyleInfo>()
  styles.set(styleId, { styleId, name: styleId, type: 'paragraph', display } as StyleInfo)
  return { styles, docDefaults: {}, blocks: [] } as unknown as ParsedDocFull
}

describe('docStyleCss paragraph borders', () => {
  it('draws style-level w:pBdr sides with the direct-border look', () => {
    const css = docStyleCss(
      parsedWith('HDR', { borderSides: { b: { color: '1F4E79', szPt: 6 }, t: {} } }),
    )
    expect(css).toContain(
      '[data-style="HDR"] { border-top:1px solid #444;border-bottom:8px solid #1F4E79;padding:1px 4px }',
    )
  })

  it('emits nothing for a side reset to none', () => {
    const css = docStyleCss(parsedWith('Off', { borderSides: { b: null } }))
    expect(css).not.toContain('border-')
    expect(css).not.toContain('padding:1px 4px')
  })
})
