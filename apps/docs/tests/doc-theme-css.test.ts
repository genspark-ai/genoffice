/**
 * Applying a theme has to change the page, not only the saved file.
 * docThemeCss is generated from live theme state, so the render reflects a pick
 * immediately.
 */
import { describe, expect, it } from 'vitest'
import type { ParsedDocFull, StyleDisplay, StyleInfo } from '@genoffice/docx-engine'
import { docLineFactor, docStyleCss, docThemeCss } from '../src/renderer/doc-style-css'

describe('docThemeCss', () => {
  it('emits body and heading fonts from the theme font pair', () => {
    const css = docThemeCss({ major: 'Trebuchet MS', minor: 'Trebuchet MS' }, null)
    expect(css).toContain('.doc-page {')
    expect(css).toContain('Trebuchet MS')
    expect(css).toContain('.doc-page h1')
  })

  it('emits the accent color for headings and the ribbon accent variable', () => {
    const css = docThemeCss(null, { accent1: '90C226' })
    expect(css).toContain('--theme-accent:#90C226')
    expect(css).toContain('color:#90C226')
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
    expect(docLineFactor(parsed, true)).toBe(1.3)
  })

  it('a Latin-only Normal (font === fontAscii) does not override the docDefaults EA font', () => {
    const parsed = parsedWith({ font: 'Calibri', fontAscii: 'Calibri' }, 'SimSun')
    expect(docLineFactor(parsed, true)).toBe(1.3)
  })

  it('falls back to the SimSun-class factor without any declared EA font', () => {
    expect(docLineFactor(parsedWith(undefined), true)).toBe(1.3)
  })

  it('a same-slot Korean Normal (Malgun in both slots) still wins for the kr factor', () => {
    ;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }
    const parsed = parsedWith({ font: 'Malgun Gothic', fontAscii: 'Malgun Gothic' }, 'Batang')
    expect(docStyleCss(parsed)).toContain('--doc-line-factor-kr:1.73')
    expect(docLineFactor(parsed, true)).toBe(1.73)
  })

  it('kr factor skips a Latin-only Normal and reads the docDefaults EA font', () => {
    // node test env has no CSS.escape
    ;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }
    const parsed = parsedWith({ font: 'Calibri', fontAscii: 'Calibri' }, 'Malgun Gothic')
    expect(docStyleCss(parsed)).toContain('--doc-line-factor-kr:1.73')
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
      '.doc-page [data-style="ListParagraph"]:not(.doc-li) { margin-inline-start:36.0pt }',
    )
    expect(css).toContain(
      '.doc-page .doc-li[data-style="ListParagraph"] { --style-li-left:36.0pt }',
    )
    expect(css).not.toContain('margin-left')
  })
})
