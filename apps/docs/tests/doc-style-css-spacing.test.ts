/**
 * Style-CSS spacing semantics measured against Word (probe 2026-08-23):
 * - before/afterAutospacing = a fixed 14pt regardless of font size, at direct,
 *   style and docDefaults level; it collapses to 0 between two list items.
 * - contextualSpacing suppresses same-style spacing whatever its source, but a
 *   direct w:contextualSpacing w:val="0" re-enables the paragraph's own spacing.
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

describe('docStyleCss spacing', () => {
  it('resolves style-level autospacing to Word auto 14pt and zeroes it between list items', () => {
    const css = docStyleCss(
      parsedWith('text-start', {
        spaceBeforeTwips: 100,
        spaceAfterTwips: 100,
        spaceBeforeAuto: true,
        spaceAfterAuto: true,
      }),
    )
    expect(css).toContain('[data-style="text-start"] { margin-top:14.0pt;margin-bottom:14.0pt }')
    expect(css).toContain('.doc-li[data-style="text-start"]:has(+ .doc-li) { margin-bottom:0 }')
    expect(css).toContain('.doc-li + .doc-li[data-style="text-start"] { margin-top:0 }')
    // no !important: a direct explicit w:before/w:after (inline margin, auto off)
    // must override the style-level list collapse
    expect(css).not.toContain(
      '.doc-li[data-style="text-start"]:has(+ .doc-li) { margin-bottom:0 !important }',
    )
    expect(css).not.toContain(
      '.doc-li + .doc-li[data-style="text-start"] { margin-top:0 !important }',
    )
  })

  it('suppresses style-level auto space-before on the document first block', () => {
    const css = docStyleCss(parsedWith('web', { spaceBeforeAuto: true }))
    expect(css).toContain(
      '.doc-page > [data-style="web"]:nth-child(1 of :not(.page-float-host)) { margin-top:0 }',
    )
  })

  it('suppresses default-level auto space-before on the document first block', () => {
    const parsed = parsedWith('S', {})
    ;(parsed as unknown as { docDefaults: object }).docDefaults = { spaceBeforeAuto: true }
    const css = docStyleCss(parsed)
    expect(css).toContain(
      ':not([data-style]):nth-child(1 of :not(.page-float-host)) { margin-top:0 }',
    )
  })

  it('keeps the literal twips when autospacing is off', () => {
    const css = docStyleCss(parsedWith('S', { spaceBeforeTwips: 100, spaceAfterTwips: 100 }))
    expect(css).toContain('[data-style="S"] { margin-top:5.0pt;margin-bottom:5.0pt }')
  })

  it('marks fixed-line styles with --doc-line-fixed for the pagination fit', () => {
    // style-level exact/atLeast lines have no doc-lh-fixed class; the marker
    // stops measureBlocks dividing an inherited document auto multiple out of
    // their full-height break-line box (Word probe 20260901: a style-level
    // exact break line demands its full box like a direct one)
    const exact = docStyleCss(parsedWith('BreakStyle', { lineRule: 'exact', lineRawTwips: 480 }))
    expect(exact).toMatch(/\[data-style="BreakStyle"\] \{[^}]*--doc-line-fixed:1/)
    const auto = docStyleCss(
      parsedWith('Body', { lineRule: 'auto', lineRawTwips: 480, spaceAfterTwips: 100 }),
    )
    expect(auto).not.toContain('--doc-line-fixed')
  })

  it('lets a direct ctxSp off (.ctx-sp-off) escape the style-level suppression', () => {
    const css = docStyleCss(parsedWith('ListBullet', { contextualSpacing: true }))
    expect(css).toContain(
      '[data-style="ListBullet"]:has(+ [data-style="ListBullet"]):not(.ctx-sp-off) { margin-bottom:0 !important }',
    )
    expect(css).toContain(
      '[data-style="ListBullet"] + [data-style="ListBullet"]:not(.ctx-sp-off) { margin-top:0 !important }',
    )
  })

  it('emits direct-ctxSp (.ctx-sp) same-style suppression for styles without style-level ctxSp', () => {
    const css = docStyleCss(parsedWith('SBul', {}))
    expect(css).toContain(
      '[data-style="SBul"].ctx-sp:has(+ [data-style="SBul"]) { margin-bottom:0 !important }',
    )
    expect(css).toContain(
      '[data-style="SBul"] + [data-style="SBul"].ctx-sp { margin-top:0 !important }',
    )
  })
})
