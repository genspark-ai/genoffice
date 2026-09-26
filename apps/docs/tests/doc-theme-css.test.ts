/**
 * Applying a theme has to change the page, not only the saved file.
 * docThemeCss is generated from live theme state, so the render reflects a pick
 * immediately.
 */
import { describe, expect, it } from 'vitest'
import type { ParsedDocFull, StyleDisplay, StyleInfo } from '@genoffice/docx-engine'
import { docLineFactor, docStyleCss, docThemeCss } from '../src/renderer/doc-style-css'

describe('docThemeCss', () => {
  it('emits the body font from the theme minor face and leaves headings to their style chain', () => {
    const css = docThemeCss({ major: 'Trebuchet MS', minor: 'Georgia' }, null)
    expect(css).toContain('.doc-page, .pv-page {')
    expect(css).toContain('Georgia')
    expect(css).not.toContain('Trebuchet MS')
    expect(css).not.toContain('.doc-page h1')
  })

  it('a heading style without rFonts inherits the declared body face, one with a theme reference gets the resolved face', () => {
    ;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }
    const styles = new Map<string, StyleInfo>()
    styles.set('Normal', {
      styleId: 'Normal',
      name: 'Normal',
      type: 'paragraph',
      isDefault: true,
      display: {},
    } as StyleInfo)
    styles.set('Heading1', {
      styleId: 'Heading1',
      name: 'heading 1',
      type: 'paragraph',
      basedOn: 'Normal',
      headingLevel: 1,
      display: { font: 'Calibri', fontAscii: 'Calibri', sizeHalfPoints: 40 },
    } as StyleInfo)
    styles.set('Heading3', {
      styleId: 'Heading3',
      name: 'heading 3',
      type: 'paragraph',
      basedOn: 'Normal',
      headingLevel: 3,
      display: { sizeHalfPoints: 28 },
    } as StyleInfo)
    const parsed = {
      styles,
      docDefaults: { asciiFont: 'Arial' },
      themeFonts: { major: 'Calibri', minor: 'Cambria' },
      blocks: [],
    } as unknown as ParsedDocFull
    const css = docStyleCss(parsed) + docThemeCss(parsed.themeFonts, null, true)
    const rule = (id: string) =>
      [...css.matchAll(/\.doc-page \[data-style="(\w+)"\] \{ ([^}]*)\}/g)]
        .filter((m) => m[1] === id)
        .map((m) => m[2])
        .join(';')
    expect(rule('Heading1')).toContain('font-family:')
    expect(rule('Heading1')).toContain('Calibri')
    expect(rule('Heading3')).not.toContain('font-family')
    expect(css).not.toMatch(/\.doc-page h3[^{]*\{[^}]*font-family/)
  })

  it('exposes the ribbon accent without coloring headings that do not declare a color', () => {
    const css = docThemeCss(null, { accent1: '90C226' })
    expect(css).toContain('--theme-accent:#90C226')
    expect(css).not.toContain('color:#90C226')
    expect(css).not.toContain('.doc-page h1')
  })

  it('is empty without a theme, so document CSS stays authoritative', () => {
    expect(docThemeCss(null, null)).toBe('')
    expect(docThemeCss({ major: '', minor: '' }, {})).toBe('')
  })
})

describe('docLineFactor — CJK factor source', () => {
  const parsedWith = (
    normalDisplay: StyleDisplay | undefined,
    eastAsiaFont?: string,
  ): ParsedDocFull => {
    const styles = new Map<string, StyleInfo>()
    if (normalDisplay) {
      styles.set('Normal', {
        styleId: 'Normal',
        name: 'Normal',
        type: 'paragraph',
        isDefault: true,
        display: normalDisplay,
      } as StyleInfo)
    }
    return {
      styles,
      docDefaults: eastAsiaFont ? { eastAsiaFont } : {},
      blocks: [],
    } as unknown as ParsedDocFull
  }

  it("Normal's declared EA face wins over docDefaults", () => {
    const parsed = parsedWith({ font: 'Noto Sans KR', fontAscii: 'Calibri' }, 'SimSun')
    expect(docLineFactor(parsed, true)).toBe(1.3029)
  })

  it('a Latin-only Normal (font === fontAscii) does not override the docDefaults EA font', () => {
    const parsed = parsedWith({ font: 'Calibri', fontAscii: 'Calibri' }, 'SimSun')
    expect(docLineFactor(parsed, true)).toBe(1.3029)
  })

  it('a same-slot Japanese Normal (Meiryo in every slot) is an EA choice, not Latin-only', () => {
    const parsed = parsedWith({ font: 'メイリオ', fontAscii: 'メイリオ' }, undefined)
    expect(docLineFactor(parsed, true)).toBe(1.9429)
  })

  it('falls back to the SimSun-class factor without any declared EA font', () => {
    expect(docLineFactor(parsedWith(undefined), true)).toBe(1.3029)
  })

  it('an EA-only Normal leaves the Latin line factor to the docDefaults ascii face', () => {
    ;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }
    const parsed = parsedWith({ font: 'Aptos', eastAsiaFont: 'Aptos' }, 'SimSun')
    ;(parsed.docDefaults as { asciiFont?: string }).asciiFont = 'Times New Roman'
    expect(docLineFactor(parsed, false)).toBe(1.15)
    expect(docStyleCss(parsed)).toContain('--doc-line-factor-latin:1.15')
  })

  it('a same-slot Korean Normal (Malgun in both slots) still wins for the kr factor', () => {
    ;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }
    const parsed = parsedWith({ font: 'Malgun Gothic', fontAscii: 'Malgun Gothic' }, 'Batang')
    expect(docStyleCss(parsed)).toContain('--doc-line-factor-kr:1.7371')
    expect(docLineFactor(parsed, true)).toBe(1.7371)
  })

  it('kr factor skips a Latin-only Normal and reads the docDefaults EA font', () => {
    // node test env has no CSS.escape
    ;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }
    const parsed = parsedWith({ font: 'Calibri', fontAscii: 'Calibri' }, 'Malgun Gothic')
    expect(docStyleCss(parsed)).toContain('--doc-line-factor-kr:1.7371')
  })
})

describe('docStyleCss — paragraph spacing fallback', () => {
  const parsedWith = (docDefaults: Record<string, unknown>): ParsedDocFull => {
    ;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }
    return { styles: new Map(), docDefaults, blocks: [] } as unknown as ParsedDocFull
  }
  const blockRule = (css: string): string =>
    css.split('\n').find((l) => l.startsWith('.doc-page p,')) ?? ''

  it('undeclared spacing renders with zero margins like Word', () => {
    const rule = blockRule(docStyleCss(parsedWith({})))
    expect(rule).toContain('margin-top:0.0pt')
    expect(rule).toContain('margin-bottom:0.0pt')
  })

  it('declared docDefaults spacing still applies', () => {
    const rule = blockRule(docStyleCss(parsedWith({ spaceAfterTwips: 200, spaceBeforeTwips: 40 })))
    expect(rule).toContain('margin-top:2.0pt')
    expect(rule).toContain('margin-bottom:10.0pt')
  })
})

describe('docStyleCss — typed line grid', () => {
  const parsedWithSectPr = (sectPr: string): ParsedDocFull => {
    ;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }
    return {
      styles: new Map(),
      docDefaults: {},
      blocks: [{ docxIndex: 0, originalXml: `<w:p><w:pPr>${sectPr}</w:pPr></w:p>` }],
    } as unknown as ParsedDocFull
  }

  it('declares --doc-line-grid and --doc-line-max per element for uniform typed grids', () => {
    const css = docStyleCss(
      parsedWithSectPr('<w:sectPr><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>'),
    )
    expect(css).toContain(
      '--doc-line-grid:round(up, calc(var(--doc-line-factor,1.2) * 1em - var(--doc-grid-pitch,0.0001px) * 0.004), var(--doc-grid-pitch,0.0001px))',
    )
    // Word probe 2026-08-22: mult x pitch, floored at the snapped single
    expect(css).toContain(
      '--doc-line-max:max(calc(var(--doc-grid-pitch,0.0001px) * var(--doc-line-mult,1)), round(up',
    )
    expect(css).toContain('--doc-grid-single-mult:1')
    // a typed grid centres the multiple's extra leading like CSS (Word probe 2026-09-23)
    expect(css).toContain('--doc-lead-grid:1')
    // snapToGrid=0 paragraphs degrade to natural x mult on the paragraph AND its spans
    expect(css).toContain(
      '.doc-page :is(.doc-nosnap, .doc-grid-nosnap), .doc-page :is(.doc-nosnap, .doc-grid-nosnap) * { --doc-line-max:calc(var(--doc-line-factor,1.2) * 1em * var(--doc-line-mult,1)); --doc-lead-grid:0 }',
    )
  })

  it('grid-less docs get no --doc-line-grid declaration, so line heights stay unitless', () => {
    const css = docStyleCss(parsedWithSectPr('<w:sectPr></w:sectPr>'))
    expect(css).not.toContain('--doc-line-grid:')
    expect(css).not.toContain('--doc-lead-grid')
    expect(css).toContain('line-height:var(--doc-line-grid,var(--doc-line-factor,1.2))')
  })
})

describe('docStyleCss — style indent vs list geometry', () => {
  const parsedWithStyle = (display: StyleDisplay): ParsedDocFull => {
    ;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }
    const styles = new Map<string, StyleInfo>()
    styles.set('ListParagraph', {
      styleId: 'ListParagraph',
      name: 'List Paragraph',
      type: 'paragraph',
      display,
    } as StyleInfo)
    return { styles, docDefaults: {}, blocks: [] } as unknown as ParsedDocFull
  }

  // Word overrides indents per property (direct > numbering > style), never adds them:
  // the style indent must skip .doc-li margins and only feed the --li-left fallback
  it('style w:ind skips list items and becomes the --li-left fallback', () => {
    const css = docStyleCss(parsedWithStyle({ indentLeftTwips: 720 }))
    expect(css).toContain(
      '.doc-page [data-style="ListParagraph"]:not(.doc-li, .doc-li-stray) { margin-inline-start:36.0pt }',
    )
    expect(css).toContain(
      '.doc-page :is(.doc-li, .doc-li-stray)[data-style="ListParagraph"] { --style-li-left:36.0pt }',
    )
    expect(css).not.toContain('margin-left')
  })
})
