/**
 * Symbol-encoded legacy fonts (Symbol / Wingdings / Webdings) carry glyphs at
 * arbitrary byte positions, stored either as that byte or as U+F0xx (private
 * use). Without the font installed they render as the raw letter or tofu, so
 * map the common glyphs to their Unicode equivalents for display.
 */

const SYMBOL: Record<number, string> = {
  // Greek (Adobe Symbol layout: Latin letter position -> Greek letter)
  0x41: 'Α',
  0x42: 'Β',
  0x43: 'Χ',
  0x44: 'Δ',
  0x45: 'Ε',
  0x46: 'Φ',
  0x47: 'Γ',
  0x48: 'Η',
  0x49: 'Ι',
  0x4b: 'Κ',
  0x4c: 'Λ',
  0x4d: 'Μ',
  0x4e: 'Ν',
  0x4f: 'Ο',
  0x50: 'Π',
  0x51: 'Θ',
  0x52: 'Ρ',
  0x53: 'Σ',
  0x54: 'Τ',
  0x55: 'Υ',
  0x57: 'Ω',
  0x58: 'Ξ',
  0x59: 'Ψ',
  0x5a: 'Ζ',
  0x61: 'α',
  0x62: 'β',
  0x63: 'χ',
  0x64: 'δ',
  0x65: 'ε',
  0x66: 'φ',
  0x67: 'γ',
  0x68: 'η',
  0x69: 'ι',
  0x6b: 'κ',
  0x6c: 'λ',
  0x6d: 'μ',
  0x6e: 'ν',
  0x6f: 'ο',
  0x70: 'π',
  0x71: 'θ',
  0x72: 'ρ',
  0x73: 'σ',
  0x74: 'τ',
  0x75: 'υ',
  0x77: 'ω',
  0x78: 'ξ',
  0x79: 'ψ',
  0x7a: 'ζ',
  // math / arrows
  0x22: '∀',
  0x24: '∃',
  0xa3: '≤',
  0xa5: '∞',
  0xab: '↔',
  0xac: '←',
  0xad: '↑',
  0xae: '→',
  0xaf: '↓',
  0xb0: '°',
  0xb1: '±',
  0xb3: '≥',
  0xb4: '×',
  0xb6: '∂',
  0xb7: '•',
  0xb8: '÷',
  0xb9: '≠',
  0xba: '≡',
  0xbb: '≈',
  0xbc: '…',
  0xc4: '⊗',
  0xc5: '⊕',
  0xc6: '∅',
  0xc7: '∩',
  0xc8: '∪',
  0xc9: '⊃',
  0xcc: '⊂',
  0xce: '∈',
  0xd1: '∇',
  0xd5: '∏',
  0xd6: '√',
  0xd7: '⋅',
  0xd8: '¬',
  0xd9: '∧',
  0xda: '∨',
  0xe5: '∑',
  0xf2: '∫',
}

const WINGDINGS: Record<number, string> = {
  0x21: '✏',
  0x22: '✂',
  0x28: '☎',
  0x2a: '✉',
  0x45: '☜',
  0x46: '☞',
  0x47: '☝',
  0x48: '☟',
  0x4a: '☺',
  0x4b: '😐',
  0x4c: '☹',
  0x6c: '●',
  0x6d: '❍',
  0x6e: '■',
  0x6f: '❑',
  0x70: '❒',
  0x75: '◆',
  0x76: '❖',
  0x78: '⌧',
  0xa7: '▪',
  0xa8: '□',
  0xab: '★',
  0xd8: '➢',
  0xfb: '✗',
  0xfc: '✓',
  0xfd: '☒',
  0xfe: '☑',
}

// 0x95-0xA6: circle/ring/square/box bullet series (sizes collapsed to the nearest Unicode shape)
const WINGDINGS_2: Record<number, string> = {
  0x95: '•',
  0x96: '●',
  0x97: '●',
  0x98: '●',
  0x99: '◦',
  0x9a: '○',
  0x9b: '○',
  0x9c: '○',
  0x9d: '◉',
  0x9e: '◉',
  0x9f: '▪',
  0xa0: '■',
  0xa1: '■',
  0xa2: '■',
  0xa3: '□',
  0xa4: '□',
  0xa5: '□',
  0xa6: '□',
}

const SYMBOL_FONT_MAPS: Record<string, Record<number, string>> = {
  symbol: SYMBOL,
  wingdings: WINGDINGS,
  'wingdings 2': WINGDINGS_2,
  // recognized as symbol-encoded so unmapped glyphs get sane fallbacks, but no table yet
  'wingdings 3': {},
  webdings: {},
}

export function isSymbolFont(font: string | null | undefined): boolean {
  return !!font && font.trim().toLowerCase() in SYMBOL_FONT_MAPS
}

/** glyph code (raw byte or its U+F0xx private-use form) → Unicode equivalent, null when unknown */
export function decodeSymbolChar(font: string, code: number): string | null {
  const map = SYMBOL_FONT_MAPS[font.trim().toLowerCase()]
  if (!map) return null
  const low = code >= 0xf000 && code <= 0xf0ff ? code - 0xf000 : code
  return map[low] ?? null
}

/** raw-byte glyph codes → their U+F0xx private-use form (how symbol-font cmaps index glyphs) */
export function toSymbolPua(text: string): string {
  let out = ''
  for (const ch of text) {
    const code = ch.codePointAt(0) as number
    out += code > 0x20 && code <= 0xff ? String.fromCodePoint(0xf000 + code) : ch
  }
  return out
}

/** whole string decode; null unless every non-whitespace char maps */
export function decodeSymbolText(font: string, text: string): string | null {
  if (!isSymbolFont(font)) return null
  let out = ''
  for (const ch of text) {
    const code = ch.codePointAt(0) as number
    if (code <= 0x20 || code === 0xa0) {
      out += ch
      continue
    }
    const mapped = decodeSymbolChar(font, code)
    if (mapped === null) return null
    out += mapped
  }
  return out
}
