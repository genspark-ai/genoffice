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
  // one list definition: the .doc-li collapse rules are keyed per list (data-num)
  const numbering = new Map([['1', { numId: '1', abstractNumId: '0' }]])
  return { styles, docDefaults: {}, blocks: [], numbering } as unknown as ParsedDocFull
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
    expect(css).toContain(
      '.doc-li[data-num="1"][data-style="text-start"]:has(+ .doc-li[data-num="1"]) { margin-bottom:0 }',
    )
    expect(css).toContain(
      '.doc-li[data-num="1"] + .doc-li[data-num="1"][data-style="text-start"] { margin-top:0 }',
    )
    // no !important: a direct explicit w:before/w:after (inline margin, auto off)
    // must override the style-level list collapse
    expect(css).not.toMatch(
      /\[data-style="text-start"\]:has\(\+ [^{]*\{ margin-bottom:0 !important/,
    )
    expect(css).not.toMatch(
      /\+ \.doc-li[^{]*\[data-style="text-start"\] \{ margin-top:0 !important/,
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

  it('zeroes style-level auto spacing at the cell boundaries (Word probe 2026-09-03)', () => {
    const css = docStyleCss(parsedWith('web', { spaceBeforeAuto: true, spaceAfterAuto: true }))
    const cell = '.doc-table :is(td, th, .cell-clip, .cell-vert) > '
    expect(css).toContain(
      `${cell}[data-style="web"]:nth-child(1 of :not(.doc-cell-boxes)) { margin-top:0 }`,
    )
    expect(css).toContain(
      `${cell}[data-style="web"]:nth-last-child(1 of :not(.doc-cell-boxes)) { margin-bottom:0 }`,
    )
  })

  it('zeroes default-level auto spacing at the cell boundaries', () => {
    const parsed = parsedWith('S', {})
    ;(parsed as unknown as { docDefaults: object }).docDefaults = {
      spaceBeforeAuto: true,
      spaceAfterAuto: true,
    }
    const css = docStyleCss(parsed)
    const cell = '.doc-table :is(td, th, .cell-clip, .cell-vert) > '
    expect(css).toContain(
      `${cell}:is(p, .doc-li, h1, h2, h3, h4, h5, h6, .doc-protected-field, .doc-img-para):not([data-style]):nth-child(1 of :not(.doc-cell-boxes)) { margin-top:0 }`,
    )
    expect(css).toContain(
      `${cell}:is(p, .doc-li, h1, h2, h3, h4, h5, h6, .doc-protected-field, .doc-img-para):not([data-style]):nth-last-child(1 of :not(.doc-cell-boxes)) { margin-bottom:0 }`,
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
    // line="0" atLeast is a natural-height fixed line: its spans must not
    // re-snap to the grid either (Word probe 2026-09-11)
    const zero = docStyleCss(parsedWith('Loose', { lineRule: 'atLeast', lineRawTwips: 0 }))
    expect(zero).toMatch(/\[data-style="Loose"\] \{[^}]*--doc-line-fixed:1/)
    expect(zero).toContain('[data-style="Loose"]:not(.doc-lh-fixed) span { line-height:inherit }')
  })

  it('emits the fixed-rule glyph shift override per style, released by auto multiples', () => {
    const atLeast = docStyleCss(parsedWith('Tight', { lineRule: 'atLeast', lineRawTwips: 480 }))
    expect(atLeast).toMatch(
      /\[data-style="Tight"\] \{[^}]*--doc-lead-top:max\(0px, \(24\.0pt - var\(--doc-line-grid, calc\(var\(--doc-line-factor,1\.2\) \* 1em\)\)\) \/ 2\)/,
    )
    const exact = docStyleCss(parsedWith('Fixed', { lineRule: 'exact', lineRawTwips: 480 }))
    expect(exact).toMatch(/\[data-style="Fixed"\] \{[^}]*--doc-lead-top:0px/)
    const auto = docStyleCss(parsedWith('Body', { lineRule: 'auto', lineRawTwips: 360 }))
    expect(auto).toMatch(/\[data-style="Body"\] \{[^}]*--doc-lead-top:initial/)
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

describe('docStyleCss styles off the Normal chain', () => {
  const parsedOff = (display: StyleDisplay, docDefaults: object = {}): ParsedDocFull => {
    const styles = new Map<string, StyleInfo>()
    styles.set('Normal', {
      styleId: 'Normal',
      name: 'Normal',
      type: 'paragraph',
      isDefault: true,
      display: {
        spaceAfterTwips: 160,
        lineRule: 'auto',
        lineRawTwips: 259,
        lineSpacing: 259 / 240,
      },
    } as StyleInfo)
    styles.set('NoSpacing', {
      styleId: 'NoSpacing',
      name: 'No Spacing',
      type: 'paragraph',
      display,
    } as StyleInfo)
    return { styles, docDefaults, blocks: [] } as unknown as ParsedDocFull
  }

  it('a paragraph style without w:basedOn takes docDefaults spacing and line, not Normal', () => {
    const css = docStyleCss(parsedOff({ sizeHalfPoints: 22 }))
    const rule = /\[data-style="NoSpacing"\] \{ ([^}]*) \}/.exec(css)![1]
    expect(rule).toContain('margin-top:0.0pt')
    expect(rule).toContain('margin-bottom:0.0pt')
    expect(rule).toMatch(/line-height:/)
    expect(rule).not.toContain('1.079')
  })

  it('docDefaults spacing reaches it when declared', () => {
    const css = docStyleCss(
      parsedOff({ sizeHalfPoints: 22 }, { spaceAfterTwips: 200, spaceBeforeAuto: true }),
    )
    const rule = /\[data-style="NoSpacing"\] \{ ([^}]*) \}/.exec(css)![1]
    expect(rule).toContain('margin-top:14.0pt')
    expect(rule).toContain('margin-bottom:10.0pt')
  })

  it('own spacing still wins over docDefaults', () => {
    const css = docStyleCss(parsedOff({ spaceAfterTwips: 60 }, { spaceAfterTwips: 200 }))
    expect(/\[data-style="NoSpacing"\] \{ ([^}]*) \}/.exec(css)![1]).toContain(
      'margin-bottom:3.0pt',
    )
  })

  it('resets the glyph shift and snapping multiple to the line it lays out with', () => {
    const rule = (css: string): string => /\[data-style="NoSpacing"\] \{ ([^}]*) \}/.exec(css)![1]
    // single line: Normal's 1.08 multiple must not shift its glyphs
    const single = rule(docStyleCss(parsedOff({ sizeHalfPoints: 22 })))
    expect(single).toContain('--doc-lead-top:initial')
    expect(single).toContain('--doc-line-mult:1')
    // docDefaults multiple
    const ddAuto = rule(
      docStyleCss(parsedOff({}, { lineRule: 'auto', lineRawTwips: 360, lineSpacing: 1.5 })),
    )
    expect(ddAuto).toContain('--doc-lead-top:initial')
    expect(ddAuto).toContain('--doc-line-mult:1.5')
    // docDefaults atLeast bottom-aligns by its own floor
    const ddAtLeast = rule(docStyleCss(parsedOff({}, { lineRule: 'atLeast', lineRawTwips: 480 })))
    expect(ddAtLeast).toMatch(/--doc-lead-top:max\(0px, \(24\.0pt - /)
    expect(ddAtLeast).not.toContain('--doc-line-mult')
    // the style's own rule still wins over docDefaults
    const own = rule(
      docStyleCss(
        parsedOff(
          { lineRule: 'auto', lineRawTwips: 480, lineSpacing: 2 },
          { lineRule: 'atLeast', lineRawTwips: 480 },
        ),
      ),
    )
    expect(own).toContain('--doc-lead-top:initial')
    expect(own).toContain('--doc-line-mult:2')
  })
})
