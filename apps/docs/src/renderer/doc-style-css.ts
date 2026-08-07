import type { ParsedDocFull, StyleDisplay, ThemeColors, ThemeFonts } from '@genoffice/docx-engine'
import {
  cssDualFontFamily,
  cssFontFamily,
  cssLineHeight,
  lineHeightFactor,
  textHasCjk,
} from './line-metrics'

/**
 * CSS for the document theme (Design ▸ Themes / Fonts / Colors). Kept separate from
 * docStyleCss so the page reflects a theme pick immediately instead of only after
 * save + reopen in Word: App re-renders this from live state, while
 * docStyleCss is regenerated only on parse.
 */
export function docThemeCss(
  fonts: ThemeFonts | null | undefined,
  colors: ThemeColors | null | undefined,
  bodyFontDeclared = false,
): string {
  const rules: string[] = []
  if (fonts?.minor && !bodyFontDeclared) {
    // Body font from the theme's minor latin face — only when neither Normal nor
    // docDefaults names one (a declared body font supersedes the theme, and
    // docStyleCss already resolved theme references into it)
    rules.push(`.doc-page { font-family:${cssFontFamily(fonts.minor)} }`)
  }
  if (fonts?.major) {
    const headings = [1, 2, 3, 4, 5, 6].map((n) => `.doc-page h${n}`).join(', ')
    rules.push(`${headings} { font-family:${cssFontFamily(fonts.major)} }`)
  }
  if (colors?.accent1) {
    // Heading color + the accent variable the ribbon's style presets already read
    rules.push(`.doc-page { --theme-accent:#${colors.accent1} }`)
    const headings = [1, 2, 3, 4].map((n) => `.doc-page h${n}`).join(', ')
    rules.push(`${headings} { color:#${colors.accent1} }`)
  }
  return rules.join('\n')
}

/**
 * Per-document CSS generated from styles.xml, so paragraphs render with their
 * style's font size / color / spacing (display-only; the save
 * path never touches styles.xml).
 */
/** Body contains CJK text (drives the document-level line-height factor). */
export function docHasCjk(parsed: ParsedDocFull): boolean {
  return parsed.blocks.some((b) => !b.hidden && (b.runs ?? []).some((r) => textHasCjk(r.text)))
}

/**
 * Document-level line-height factor: bodies containing CJK use the Chinese font's
 * factor (Word takes the max of in-line fonts; the declared eastAsia default font
 * doesn't reflect actual content, and pure-English documents shouldn't get CJK
 * line height). Recomputed live while editing via App's liveDocCjk.
 */
export function docLineFactor(parsed: ParsedDocFull, hasCjk: boolean): number {
  const dd = parsed.docDefaults
  return hasCjk
    ? lineHeightFactor(dd?.eastAsiaFont ?? '宋体')
    : lineHeightFactor(docBodyFont(parsed) ?? 'Calibri')
}

/** the w:default="1" paragraph style's display (Word's baseline for un-styled paragraphs) */
export function defaultParaDisplay(parsed: ParsedDocFull): StyleDisplay | undefined {
  for (const info of parsed.styles.values()) {
    if (info.isDefault && info.type === 'paragraph' && info.display) return info.display
  }
  return undefined
}

/** Latin body font the document declares (Normal style or docDefaults, theme refs
 * resolved). Ascii slot first — StyleDisplay.font is eastAsia-first and would drag
 * the Latin line factor / theme override onto the CJK face. */
export function docBodyFont(parsed: ParsedDocFull): string | undefined {
  const normal = defaultParaDisplay(parsed)
  return normal?.fontAscii ?? normal?.font ?? parsed.docDefaults?.asciiFont
}

export function docStyleCss(parsed: ParsedDocFull): string {
  const rules: string[] = []
  const dd = parsed.docDefaults
  // Word applies the w:default="1" paragraph style (Normal) to every paragraph
  // without a w:pStyle, so its display merges into the document baseline here
  // ([data-style] rules only reach explicitly styled paragraphs).
  const normal = defaultParaDisplay(parsed)
  {
    const decls: string[] = []
    // Paragraph level also overrides this variable per paragraph's text (blockAttrs
    // at parse time + live decorations in LineFactorExtension).
    const factor = docLineFactor(parsed, docHasCjk(parsed))
    decls.push(`--doc-line-factor:${factor}`)
    // Latin factor for per-paragraph overrides (blockAttrs): pure-Western paragraphs
    // follow the body font's real single-line metric instead of a flat 1.2
    decls.push(`--doc-line-factor-latin:${lineHeightFactor(docBodyFont(parsed) ?? 'Calibri')}`)
    // dual-slot baseline: Latin families first, then the East Asian chain
    const baseAscii = normal?.fontAscii ?? dd?.asciiFont
    const baseEa = normal?.font ?? dd?.eastAsiaFont
    decls.push(
      `font-family:${
        baseAscii && baseEa && baseAscii !== baseEa
          ? cssDualFontFamily(baseAscii, baseEa)
          : cssFontFamily(baseEa ?? baseAscii ?? 'Calibri')
      }`,
    )
    const sizeHalf = normal?.sizeHalfPoints ?? dd?.sizeHalfPoints
    if (sizeHalf) decls.push(`font-size:${sizeHalf / 2}pt`)
    const color = normal?.color ?? dd?.color
    if (color) decls.push(`color:#${color}`)
    if (normal?.bold ?? dd?.bold) decls.push('font-weight:600')
    if (normal?.italic ?? dd?.italic) decls.push('font-style:italic')
    const lh =
      cssLineHeight(normal?.lineRule, normal?.lineRawTwips, normal?.lineSpacing) ??
      cssLineHeight(dd?.lineRule, dd?.lineRawTwips, dd?.lineSpacing)
    decls.push(`line-height:${lh ?? `calc(${factor} * 1)`}`)
    rules.push(`.doc-page { ${decls.join(';')} }`)
    // docDefaults paragraph spacing is Word's real fallback (the static stylesheet's
    // 8pt is a guess); declared per block so --doc-line-factor set inline on a
    // paragraph re-evaluates the line-height var (it wouldn't through inheritance)
    const blockSel =
      '.doc-page p, .doc-page .doc-li, .doc-page h1, .doc-page h2, .doc-page h3, .doc-page h4, .doc-page h5, .doc-page h6'
    const blockDecls = [
      `margin-top:${((normal?.spaceBeforeTwips ?? dd?.spaceBeforeTwips ?? 0) / 20).toFixed(1)}pt`,
      `margin-bottom:${((normal?.spaceAfterTwips ?? dd?.spaceAfterTwips ?? 160) / 20).toFixed(1)}pt`,
      `line-height:${lh ?? `calc(${factor} * 1)`}`,
    ]
    rules.push(`${blockSel} { ${blockDecls.join(';')} }`)
    // Normal's first-line indent applies to plain body paragraphs (not lists —
    // their geometry runs on --li-left/--li-hang)
    if ((normal?.indentFirstLineTwips ?? 0) > 0) {
      rules.push(
        `.doc-page p { text-indent:${((normal!.indentFirstLineTwips as number) / 20).toFixed(1)}pt }`,
      )
    }
  }
  // table styles: tables carrying data-tbl-style are colored by style (explicit cell shading
  // is inline style and naturally overrides these rules; parse gives exact display after save)
  for (const info of parsed.styles.values()) {
    const t = info.tableDisplay
    if (info.type !== 'table' || !t) continue
    const sel = `.doc-page table[data-tbl-style="${CSS.escape(info.styleId)}"]`
    if (t.fill) rules.push(`${sel} td, ${sel} th { background:#${t.fill} }`)
    // band1 = first data row after the header → even nth-child when a header row exists
    if (t.band1Fill) {
      rules.push(`${sel} tr:nth-child(even) td { background:#${t.band1Fill} }`)
    }
    if (t.band2Fill) {
      rules.push(`${sel} tr:nth-child(odd):not(:first-child) td { background:#${t.band2Fill} }`)
    }
    if (t.firstRow) {
      const decls: string[] = []
      if (t.firstRow.fill) decls.push(`background:#${t.firstRow.fill}`)
      if (t.firstRow.bold) decls.push('font-weight:600')
      if (t.firstRow.color) decls.push(`color:#${t.firstRow.color}`)
      if (decls.length > 0)
        rules.push(`${sel} tr:first-child td, ${sel} tr:first-child th { ${decls.join(';')} }`)
    }
    if (t.paraSpacing) {
      // Word precedence: paragraph style (Normal) > table style pPr > docDefaults —
      // emit only the table-style values Normal doesn't declare itself
      const ps = t.paraSpacing
      const decls: string[] = []
      if (ps.beforeTwips !== undefined && normal?.spaceBeforeTwips === undefined)
        decls.push(`margin-top:${(ps.beforeTwips / 20).toFixed(1)}pt`)
      if (ps.afterTwips !== undefined && normal?.spaceAfterTwips === undefined)
        decls.push(`margin-bottom:${(ps.afterTwips / 20).toFixed(1)}pt`)
      const psLh = cssLineHeight(ps.lineRule, ps.lineRawTwips, ps.lineSpacing)
      const normalLh = cssLineHeight(normal?.lineRule, normal?.lineRawTwips, normal?.lineSpacing)
      if (psLh && !normalLh) decls.push(`line-height:${psLh}`)
      if (decls.length > 0) {
        rules.push(
          `${sel} td p, ${sel} th p, ${sel} td .doc-li, ${sel} th .doc-li { ${decls.join(';')} }`,
        )
      }
    }
  }
  for (const info of parsed.styles.values()) {
    const d = info.display
    if (!d) continue
    const decls: string[] = []
    if (d.sizeHalfPoints) decls.push(`font-size:${d.sizeHalfPoints / 2}pt`)
    if (d.color) decls.push(`color:#${d.color}`)
    if (d.bold) decls.push('font-weight:600')
    if (d.italic) decls.push('font-style:italic')
    if (d.underline || d.strike) {
      decls.push(
        `text-decoration:${[d.underline && 'underline', d.strike && 'line-through'].filter(Boolean).join(' ')}`,
      )
    }
    if (d.font) {
      decls.push(
        `font-family:${
          d.fontAscii && d.fontAscii !== d.font
            ? cssDualFontFamily(d.fontAscii, d.font)
            : cssFontFamily(d.font)
        }`,
      )
    } else if (d.fontAscii) {
      decls.push(`font-family:${cssFontFamily(d.fontAscii)}`)
    }
    if (d.charSpacingTwips) decls.push(`letter-spacing:${d.charSpacingTwips / 20}pt`)
    const styleLh = cssLineHeight(d.lineRule, d.lineRawTwips, d.lineSpacing)
    if (styleLh) decls.push(`line-height:${styleLh}`)
    if (d.spaceBeforeTwips !== undefined)
      decls.push(`margin-top:${(d.spaceBeforeTwips / 20).toFixed(1)}pt`)
    if (d.spaceAfterTwips !== undefined)
      decls.push(`margin-bottom:${(d.spaceAfterTwips / 20).toFixed(1)}pt`)
    if (d.indentLeftTwips) decls.push(`margin-left:${(d.indentLeftTwips / 20).toFixed(1)}pt`)
    if (d.indentRightTwips) decls.push(`margin-right:${(d.indentRightTwips / 20).toFixed(1)}pt`)
    if (d.indentFirstLineTwips)
      decls.push(`text-indent:${(d.indentFirstLineTwips / 20).toFixed(1)}pt`)
    if (d.align) decls.push(`text-align:${d.align}`)
    // the static sheet guesses italic for h4-h6 (Word's built-in defaults);
    // a real style definition without w:i means upright
    if (info.headingLevel && info.headingLevel >= 4 && !d.italic) decls.push('font-style:normal')
    if (decls.length > 0) {
      rules.push(`.doc-page [data-style="${CSS.escape(info.styleId)}"] { ${decls.join(';')} }`)
    }
    // w:contextualSpacing: consecutive same-style paragraphs swallow the spacing
    // between them (ListParagraph/ListBullet carry this — Word lists are tight)
    if (d.contextualSpacing) {
      const s = `[data-style="${CSS.escape(info.styleId)}"]`
      rules.push(`.doc-page ${s}:has(+ ${s}) { margin-bottom:0 }`)
      rules.push(`.doc-page ${s} + ${s} { margin-top:0 }`)
    }
  }
  return rules.join('\n')
}
