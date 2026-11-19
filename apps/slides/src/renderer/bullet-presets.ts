import { formatAutoNum } from '@genoffice/pptx-render'

/**
 * PowerPoint's Bullets gallery (WPS ships the same seven). Each preset is the font + code
 * PowerPoint itself writes, so a deck bulleted here shows the same glyph over there;
 * `glyph` is the Unicode lookalike the ribbon tile draws with.
 */
export interface BulletPreset {
  char: string
  font: string
  glyph: string
}

export const BULLET_PRESETS: readonly BulletPreset[] = [
  { char: '•', font: 'Arial', glyph: '•' },
  { char: 'o', font: 'Courier New', glyph: '○' },
  { char: '§', font: 'Wingdings', glyph: '▪' },
  { char: 'q', font: 'Wingdings', glyph: '❑' },
  { char: 'v', font: 'Wingdings', glyph: '❖' },
  { char: 'Ø', font: 'Wingdings', glyph: '➢' },
  { char: 'ü', font: 'Wingdings', glyph: '✓' },
]

/** Symbols offered beyond the gallery: plain Unicode, written as buChar without a buFont */
export const EXTRA_BULLET_SYMBOLS = [
  '\u25c6',
  '\u25c7',
  '\u2605',
  '\u2606',
  '\u2726',
  '\u27a4',
  '\u2794',
  '\u2192',
  '\u2713',
  '\u2717',
  '\u2666',
  '\u25c9',
  '\u25ce',
  '\u25b8',
  '\u2023',
  '\u2013',
] as const

/** PowerPoint's Numbering gallery (seven) plus the two East Asian schemes WPS offers */
export interface NumberPreset {
  numType: string
  /** First three numbers as the ribbon tile shows them */
  sample: string[]
}
export const NUMBER_PRESETS: readonly NumberPreset[] = [
  'arabicPeriod',
  'arabicParenR',
  'romanUcPeriod',
  'alphaUcPeriod',
  'alphaLcParenR',
  'alphaLcPeriod',
  'romanLcPeriod',
  'circleNumDbPlain',
  'ea1ChsPeriod',
].map((numType) => ({ numType, sample: [1, 2, 3].map((n) => formatAutoNum(n, numType)) }))

/** Hanging indents offered next to the gallery (EMU); Normal is what PowerPoint writes for a plain text box */
export const BULLET_HANG_PRESETS = [
  ['ribbonBulletHangNarrow', 171450],
  ['ribbonBulletHangNormal', 285750],
  ['ribbonBulletHangWide', 342900],
] as const

const SYMBOL_BULLET_RE = /^(wingdings|webdings)/i

/**
 * Text of the bullet run the layout emits for a char + buFont (symbol-font codes normalize
 * into the font's F0xx private range) — what the gallery highlight compares against.
 */
export function bulletRunText(char: string, font?: string): string {
  if (!font || !SYMBOL_BULLET_RE.test(font)) return char
  const cp = char.codePointAt(0) ?? 0
  return cp >= 0x20 && cp <= 0xff ? String.fromCodePoint(0xf000 + cp) : char
}
