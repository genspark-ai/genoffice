import type { ParsedDocFull, StyleDisplay, ThemeColors, ThemeFonts } from '@genoffice/docx-engine'
import {
  cssDualFontFamily,
  cssFontFamily,
  cssGridLineBase,
  cssGridSpacingPt,
  cssLineHeight,
  isCjkFontName,
  krLineFactor,
  lineHeightFactor,
  textHasCjk,
  isKoreanFontName,
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
  return hasCjk ? docCjkFactor(parsed) : lineHeightFactor(docBodyFont(parsed) ?? 'Calibri')
}

/** CJK line-height factor of the document's East Asian face (feeds --doc-line-factor-cjk:
 *  per-paragraph script overrides resolve CJK paragraphs through this var). */
export function docCjkFactor(parsed: ParsedDocFull): number {
  // Normal's EA face wins over docDefaults (a Normal declaring e.g. Noto KR must
  // not fall back to the SimSun factor). font === fontAscii means only a
  // Latin slot was declared (StyleDisplay.font is EA-first) — not an EA choice.
  const normal = defaultParaDisplay(parsed)
  const normalEa =
    normal?.font && (normal.font !== normal.fontAscii || isKoreanFontName(normal.font))
      ? normal.font
      : undefined
  if (normalEa && !normal?.eaSlotEmpty) return lineHeightFactor(normalEa)
  const dd = parsed.docDefaults
  if (dd?.eastAsiaFont && !dd.eaSlotEmpty) return lineHeightFactor(dd.eastAsiaFont)
  // An empty EA theme slot can still use a CJK-capable Latin theme face.
  // Japanese templates commonly put Yu Mincho or Yu Gothic in the Latin
  // slots, and LibreOffice lays undeclared CJK out with that face.
  // A Latin theme face (Calibri…) can't render CJK, so the lang backfill wins.
  const themeLatin = parsed.themeFonts?.minor
  if ((normal?.eaSlotEmpty || dd?.eaSlotEmpty) && themeLatin && isCjkFontName(themeLatin)) {
    return lineHeightFactor(themeLatin)
  }
  return lineHeightFactor(normalEa ?? dd?.eastAsiaFont ?? 'SimSun')
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
    // CJK paragraphs resolve their per-paragraph factor through this var
    // (paraLineFactorCss); value follows the document's East Asian face
    decls.push(`--doc-line-factor-cjk:${docCjkFactor(parsed)}`)
    // Latin factor for per-paragraph overrides (blockAttrs): pure-Western paragraphs
    // follow the body font's real single-line metric instead of a flat 1.2
    decls.push(`--doc-line-factor-latin:${lineHeightFactor(docBodyFont(parsed) ?? 'Calibri')}`)
    // Korean factor for hangul paragraphs (Batang-class 1.15 unless the EA face says otherwise)
    const normalEaKr =
      normal?.font && (normal.font !== normal.fontAscii || isKoreanFontName(normal.font))
        ? normal.font
        : undefined
    decls.push(`--doc-line-factor-kr:${krLineFactor(normalEaKr ?? dd?.eastAsiaFont)}`)
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
    // fallback references the var (not the resolved number) so per-paragraph
    // script factors and docGrid snapping re-evaluate on each block
    decls.push(`line-height:${lh ?? cssGridLineBase()}`)
    rules.push(`.doc-page { ${decls.join(';')} }`)
    // Word's fallback when neither Normal nor docDefaults declares w:spacing is 0
    // (the static stylesheet's 8pt guess inflated undeclared docs, table cells worst);
    // declared per block so --doc-line-factor set inline on a paragraph re-evaluates
    // the line-height var (it wouldn't through inheritance)
    const blockSel =
      '.doc-page p, .doc-page .doc-li, .doc-page h1, .doc-page h2, .doc-page h3, .doc-page h4, .doc-page h5, .doc-page h6, .doc-page .doc-protected-field'
    const blockDecls = [
      `margin-top:${cssGridSpacingPt((normal?.spaceBeforeTwips ?? dd?.spaceBeforeTwips ?? 0) / 20)}`,
      `margin-bottom:${cssGridSpacingPt((normal?.spaceAfterTwips ?? dd?.spaceAfterTwips ?? 0) / 20)}`,
      `line-height:${lh ?? cssGridLineBase()}`,
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
        decls.push(`margin-top:${cssGridSpacingPt(ps.beforeTwips / 20)}`)
      if (ps.afterTwips !== undefined && normal?.spaceAfterTwips === undefined)
        decls.push(`margin-bottom:${cssGridSpacingPt(ps.afterTwips / 20)}`)
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
      // style-declared EA face re-anchors the CJK line factor for its paragraphs
      // (runs without their own fonts resolve --doc-line-factor-cjk through this);
      // an empty-theme-slot backfill is not a document choice and stays silent
      if (!d.eaSlotEmpty && (d.font !== d.fontAscii || isCjkFontName(d.font))) {
        decls.push(`--doc-line-factor-cjk:${lineHeightFactor(d.font)}`)
      }
    } else if (d.fontAscii) {
      decls.push(`font-family:${cssFontFamily(d.fontAscii)}`)
    }
    if (d.charSpacingTwips) decls.push(`letter-spacing:${d.charSpacingTwips / 20}pt`)
    const styleLh = cssLineHeight(d.lineRule, d.lineRawTwips, d.lineSpacing)
    if (styleLh) decls.push(`line-height:${styleLh}`)
    if (d.spaceBeforeTwips !== undefined)
      decls.push(`margin-top:${cssGridSpacingPt(d.spaceBeforeTwips / 20)}`)
    if (d.spaceAfterTwips !== undefined)
      decls.push(`margin-bottom:${cssGridSpacingPt(d.spaceAfterTwips / 20)}`)
    if (d.indentRightTwips)
      decls.push(`margin-inline-end:${(d.indentRightTwips / 20).toFixed(1)}pt`)
    if (d.indentFirstLineTwips)
      decls.push(`text-indent:${(d.indentFirstLineTwips / 20).toFixed(1)}pt`)
    if (d.align) decls.push(`text-align:${d.align}`)
    // the static sheet guesses italic for h4-h6 (Word's built-in defaults);
    // a real style definition without w:i means upright
    if (info.headingLevel && info.headingLevel >= 4 && !d.italic) decls.push('font-style:normal')
    if (decls.length > 0) {
      rules.push(`.doc-page [data-style="${CSS.escape(info.styleId)}"] { ${decls.join(';')} }`)
    }
    // Word merges indents per property (direct ind > numbering level ind > style ind), never
    // adds them: list items run on --li-left geometry, so the style indent must not also apply
    // as a margin — it only feeds the --li-left fallback chain (styles.css)
    if (d.indentLeftTwips) {
      const s = `[data-style="${CSS.escape(info.styleId)}"]`
      const pt = (d.indentLeftTwips / 20).toFixed(1)
      rules.push(`.doc-page ${s}:not(.doc-li) { margin-inline-start:${pt}pt }`)
      rules.push(`.doc-page .doc-li${s} { --style-li-left:${pt}pt }`)
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
