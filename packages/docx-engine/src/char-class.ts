/**
 * The one character classification every Word-layout rule reads. Rules never
 * spell code-point ranges themselves: a class here is the vocabulary (Word's
 * single-byte / double-byte split, East Asian script membership, the gaps
 * autospace pads), and each predicate names the set a probe measured.
 */
export type CharClass =
  | 'han'
  | 'kana'
  | 'hangul'
  | 'bopomofo'
  | 'cjkPunct'
  | 'cjkSymbol'
  | 'fullwidth'
  | 'halfwidthEA'
  | 'ideographicSpace'
  | 'latinLetter'
  | 'digit'
  | 'asciiPunct'
  | 'space'
  | 'nbsp'
  | 'tab'
  | 'other'

export function charClass(cp: number): CharClass {
  if (cp < 0x80) {
    if (cp === 0x20) return 'space'
    if (cp === 0x09) return 'tab'
    if (cp >= 0x30 && cp <= 0x39) return 'digit'
    if ((cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a)) return 'latinLetter'
    if (cp > 0x20 && cp < 0x7f) return 'asciiPunct'
    return 'other'
  }
  if (cp === 0xa0) return 'nbsp'
  if (cp < 0x1100) {
    // Latin-1 Supplement / Extended letters, minus the multiply and divide signs
    if (cp >= 0xc0 && cp <= 0x24f && cp !== 0xd7 && cp !== 0xf7) return 'latinLetter'
    return 'other'
  }
  if (cp <= 0x11ff) return 'hangul'
  if (cp < 0x2e80) return 'other'
  if (cp <= 0x2fff) return 'han' // Kangxi / CJK radicals, ideographic description
  if (cp === 0x3000) return 'ideographicSpace'
  if (cp <= 0x303f) {
    if (cp === 0x3005 || cp === 0x3007 || (cp >= 0x3021 && cp <= 0x3029)) return 'han'
    return 'cjkPunct'
  }
  if (cp <= 0x30ff) return cp === 0x30fb ? 'cjkPunct' : 'kana'
  if (cp <= 0x312f) return 'bopomofo'
  if (cp <= 0x318f) return 'hangul'
  if (cp <= 0x319f) return 'cjkSymbol' // kanbun
  if (cp <= 0x31bf) return 'bopomofo'
  if (cp <= 0x31ef) return 'cjkSymbol' // CJK strokes
  if (cp <= 0x31ff) return 'kana'
  if (cp <= 0x33ff) return 'cjkSymbol' // enclosed letters and months, compatibility
  if (cp <= 0x4dbf) return 'han'
  if (cp <= 0x4dff) return 'cjkSymbol' // Yijing hexagrams
  if (cp <= 0x9fff) return 'han'
  if (cp >= 0xa960 && cp <= 0xa97f) return 'hangul'
  if (cp >= 0xac00 && cp <= 0xd7ff) return 'hangul'
  if (cp >= 0xf900 && cp <= 0xfaff) return 'han'
  if (cp >= 0xff00 && cp <= 0xffef) {
    if (cp <= 0xff60) return 'fullwidth'
    if (cp <= 0xffdc) return 'halfwidthEA'
    if (cp >= 0xffe0 && cp <= 0xffe6) return 'fullwidth'
    if (cp >= 0xffe8 && cp <= 0xffee) return 'halfwidthEA'
    return 'other'
  }
  if (cp >= 0x20000 && cp <= 0x3ffff) return 'han'
  return 'other'
}

/** hangul syllables, jamo and compatibility jamo (Korean line and wrap rules) */
export function isHangulChar(cp: number): boolean {
  return charClass(cp) === 'hangul'
}

/** every East Asian character: the scripts, their punctuation, symbols and width forms */
export function isEastAsianChar(cp: number): boolean {
  switch (charClass(cp)) {
    case 'han':
    case 'kana':
    case 'hangul':
    case 'bopomofo':
    case 'cjkPunct':
    case 'cjkSymbol':
    case 'fullwidth':
    case 'halfwidthEA':
    case 'ideographicSpace':
      return true
    default:
      return false
  }
}

/** Word's double-byte character: an East Asian character of full-width advance
 *  (halfwidth forms are single-byte for the character grid and balance rules) */
export function isDoubleByteChar(cp: number): boolean {
  return isEastAsianChar(cp) && charClass(cp) !== 'halfwidthEA'
}

/** Han / kana / hangul letters: the East Asian side of an autoSpaceDE/DN gap
 *  (punctuation, symbols and width forms get none) */
export function isCjkLetterChar(cp: number): boolean {
  switch (charClass(cp)) {
    case 'han':
    case 'kana':
    case 'hangul':
    case 'bopomofo':
      return true
    default:
      return false
  }
}

/** Latin letters and digits: the other side of an autospace gap */
export function isLatinAlnumChar(cp: number): boolean {
  const c = charClass(cp)
  return c === 'latinLetter' || c === 'digit'
}

export function textHasEastAsian(text: string): boolean {
  for (const ch of text) if (isEastAsianChar(ch.codePointAt(0)!)) return true
  return false
}

export function textHasHangulChars(text: string): boolean {
  for (const ch of text) if (isHangulChar(ch.codePointAt(0)!)) return true
  return false
}

/** w:lang tag of an East Asian language (zh / ja / ko); `korean:false` for the
 *  ja/zh-only rules (punctuation compression, hanging: Word probe
 *  2026-09-05 ko-KR neither hangs nor compresses) */
export function isEastAsianLangTag(tag: string | null | undefined, korean = true): boolean {
  if (!tag) return false
  return (korean ? /^(?:ja|zh|ko)(?:[-_]|$)/i : /^(?:ja|zh)(?:[-_]|$)/i).test(tag)
}
