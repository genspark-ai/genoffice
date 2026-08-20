/**
 * PostScript font name → family name heuristic (P5). PDFs frequently carry
 * only the PS name (`NotoSansSC-Regular`, `Arial-BoldMT`); written into
 * rFonts as-is, Word/LibreOffice cannot resolve it and substitute a different
 * face, shifting every metric. FPDFFont_GetFamilyName is preferred; this
 * mapping is the fallback when the descriptor has no family entry.
 */

/** `ABCDEF+` subset tag PDF producers prepend to embedded font names */
export const SUBSET_PREFIX = /^[A-Z]{6}\+/

/** style modifiers that live in the family name's hyphen/comma suffix,
 * including the abbreviations type foundries glue on (HelveticaNeueLTStd-BdIt) */
const STYLE_WORD =
  /^(regular|bold|italic|oblique|light|medium|semibold|demibold|extrabold|ultrabold|black|heavy|thin|extralight|ultralight|book|roman|normal|plain|w\d+|reg|bd|blk|lt|md|it|ital|obl|demi)$/i

/** foundry tags glued onto base names by legacy converters (ArialMT, TimesNewRomanPSMT) */
const FOUNDRY_TAG = /^(MT|PS|PSMT)$/

/**
 * Insert spaces at lower→Upper camelCase boundaries: NotoSansSC → Noto Sans SC.
 * Uppercase runs deliberately stay whole ("STSong", "CJKjp") — splitting them
 * mangles real family names more often than it helps.
 */
const splitCamel = (s: string): string => s.replace(/([a-z0-9])([A-Z])/g, '$1 $2')

/** brand words that are camelCase INSIDE the real family name — re-join after splitting */
const COMPOUND_FIXUPS: Array<[RegExp, string]> = [[/\bPing Fang\b/g, 'PingFang']]

const segmentTokens = (segment: string): string[] =>
  splitCamel(segment)
    .split(' ')
    .filter((t) => t.length > 0 && !FOUNDRY_TAG.test(t))

// ── explicit style declared by the PS name (P21 A) ──
// PDFium's FPDFText_GetFontWeight is derived from the descriptor's /StemV
// (not /FontWeight), and producers routinely write StemV values that read as
// bold (dvips stamps 154 on Times-Roman). A style token in the PS name is the
// author's explicit declaration and outranks that fiction.
const BOLD_STYLE = /^(bold|semibold|demibold|extrabold|ultrabold|black|heavy|demi|bd|blk)$/i
const REGULAR_STYLE =
  /^(regular|roman|normal|book|plain|light|thin|extralight|ultralight|medium|reg|lt|md)$/i
const ITALIC_STYLE = /^(italic|oblique|ital|it|obl)$/i

export type PsNameWeight = 'bold' | 'regular' | null

const suffixTokensOf = (psName: string): string[] => {
  const [, ...suffixes] = psName.replace(SUBSET_PREFIX, '').trim().split(/[-,]/)
  return suffixes.flatMap((seg) => splitCamel(seg).split(/[\s.]/)).filter((t) => t.length > 0)
}

/**
 * Weight explicitly declared by the PS name's style suffix ('Times-Roman' →
 * 'regular', 'Helvetica,Bold' → 'bold'), or by a style word glued onto the
 * base name ('ArialBold' — bold only: a trailing 'Roman'/'Book' inside an
 * unsuffixed base is family vocabulary, not a weight claim). Null when the
 * name declares nothing.
 */
export function weightFromPsName(psName: string): PsNameWeight {
  const tokens = suffixTokensOf(psName)
  if (tokens.some((t) => BOLD_STYLE.test(t))) return 'bold'
  const base = psName.replace(SUBSET_PREFIX, '').trim().split(/[-,]/)[0] ?? ''
  const baseTokens = splitCamel(base).split(' ')
  const lastBase = baseTokens[baseTokens.length - 1] ?? ''
  if (baseTokens.length > 1 && BOLD_STYLE.test(lastBase)) return 'bold'
  if (tokens.some((t) => REGULAR_STYLE.test(t))) return 'regular'
  return null
}

/** the PS name's style suffix declares an italic/oblique face */
export function italicFromPsName(psName: string): boolean {
  return suffixTokensOf(psName).some((t) => ITALIC_STYLE.test(t))
}

/**
 * Trailing pure-style words dropped from a SPACED family name ('Helvetica
 * Neue LTStd It' → 'Helvetica Neue LTStd') — the style belongs in w:b/w:i,
 * and the styled alias never resolves against installed families.
 */
export function stripTrailingStyleWords(family: string): string {
  const tokens = family.trim().split(/\s+/)
  while (tokens.length > 1 && STYLE_WORD.test(tokens[tokens.length - 1]!)) tokens.pop()
  return tokens.join(' ')
}

/**
 * Best-effort family name from a PostScript name. Names that already contain
 * spaces are family names and pass through untouched.
 */
export function familyFromPsName(psName: string): string {
  const name = psName.replace(SUBSET_PREFIX, '').trim()
  if (!name || name.includes(' ')) return name
  // suffix segments that are pure style modifiers get dropped
  // (NotoSansSC-Regular, Arial-BoldMT, Helvetica,Bold)
  const [base = '', ...suffixes] = name.split(/[-,]/)
  const kept = suffixes.filter((seg) => {
    const tokens = segmentTokens(seg)
    return tokens.length > 0 && !tokens.every((t) => STYLE_WORD.test(t))
  })
  let family = [base, ...kept].flatMap(segmentTokens).join(' ')
  for (const [pattern, replacement] of COMPOUND_FIXUPS) {
    family = family.replace(pattern, replacement)
  }
  return family
}
