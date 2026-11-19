/**
 * w:autoHyphenation → CSS hyphens, gated by w:suppressAutoHyphens (github run
 * 20260825 samples 16/91: Word does not hyphenate when pPrDefault carries
 * w:suppressAutoHyphens even though settings.xml enables autoHyphenation).
 */
import { describe, expect, it } from 'vitest'
import type { DocDefaults, ParsedDocFull, StyleInfo } from '@genoffice/docx-engine'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

function parsedDoc(opts: {
  autoHyphenation?: boolean
  docDefaults?: DocDefaults
  styles?: StyleInfo[]
}): ParsedDocFull {
  const styles = new Map<string, StyleInfo>()
  for (const s of opts.styles ?? []) styles.set(s.styleId, s)
  return {
    styles,
    docDefaults: opts.docDefaults ?? {},
    blocks: [],
    ...(opts.autoHyphenation ? { autoHyphenation: true } : {}),
  } as unknown as ParsedDocFull
}

const AUTO = '.doc-page { hyphens:auto; -webkit-hyphens:auto; --doc-shy:"\\AD" }'

describe('autoHyphenation CSS', () => {
  it('emits document-wide hyphens:auto for plain autoHyphenation', () => {
    const css = docStyleCss(parsedDoc({ autoHyphenation: true }))
    expect(css).toContain(AUTO)
    // the strips and notes beside the editor root draw our break points too
    expect(css).toContain('.page-wrap, .pv-page { --doc-shy:"\\AD" }')
  })

  it('emits nothing without autoHyphenation', () => {
    expect(docStyleCss(parsedDoc({}))).not.toContain('hyphens')
  })

  it('pPrDefault suppressAutoHyphens turns the document-wide rule off', () => {
    const css = docStyleCss(
      parsedDoc({ autoHyphenation: true, docDefaults: { suppressAutoHyphens: true } }),
    )
    expect(css).not.toContain(AUTO)
  })

  it('default paragraph style suppressAutoHyphens turns the document-wide rule off', () => {
    const css = docStyleCss(
      parsedDoc({
        autoHyphenation: true,
        styles: [
          {
            styleId: 'Normal',
            name: 'Normal',
            type: 'paragraph',
            isDefault: true,
            display: { suppressAutoHyphens: true },
          } as StyleInfo,
        ],
      }),
    )
    expect(css).not.toContain(AUTO)
  })

  it("default style's explicit off overrides a suppressing pPrDefault", () => {
    const css = docStyleCss(
      parsedDoc({
        autoHyphenation: true,
        docDefaults: { suppressAutoHyphens: true },
        styles: [
          {
            styleId: 'Normal',
            name: 'Normal',
            type: 'paragraph',
            isDefault: true,
            display: { suppressAutoHyphens: false },
          } as StyleInfo,
        ],
      }),
    )
    expect(css).toContain(AUTO)
  })

  it('styled paragraphs override per style in both directions', () => {
    const css = docStyleCss(
      parsedDoc({
        autoHyphenation: true,
        styles: [
          {
            styleId: 'NoHyph',
            name: 'NoHyph',
            type: 'paragraph',
            isDefault: true,
            display: { suppressAutoHyphens: true },
          } as StyleInfo,
          {
            styleId: 'ReHyph',
            name: 'ReHyph',
            type: 'paragraph',
            isDefault: true,
            display: { suppressAutoHyphens: false },
          } as StyleInfo,
        ],
      }),
    )
    expect(css).toContain(
      '[data-style="NoHyph"] { hyphens:manual;-webkit-hyphens:manual;--doc-shy:none }',
    )
    expect(css).toContain(
      '[data-style="ReHyph"] { hyphens:auto;-webkit-hyphens:auto;--doc-shy:"\\AD" }',
    )
  })

  it('style values are inert without autoHyphenation', () => {
    const css = docStyleCss(
      parsedDoc({
        styles: [
          {
            styleId: 'NoHyph',
            name: 'NoHyph',
            type: 'paragraph',
            isDefault: true,
            display: { suppressAutoHyphens: true },
          } as StyleInfo,
        ],
      }),
    )
    expect(css).not.toContain('hyphens')
  })
})
