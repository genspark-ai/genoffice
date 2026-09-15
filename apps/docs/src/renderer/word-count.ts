/**
 * Word-parity text statistics.
 *
 * Word's CJK rule: Words = Asian characters (counted one by one, punctuation
 * included) + non-Asian words (whitespace/punct-delimited). The dialog also
 * reports the two addends separately — Chinese users care about the
 * Asian-character figure.
 */

// Han (incl. radicals/compat/ext-B+), kana, hangul, bopomofo, CJK symbols
// and punctuation (U+3001 up: the ideographic space stays whitespace),
// fullwidth forms. The surrogate pair covers CJK Extensions B-H
// (U+20000-U+3134F, high surrogates D840-D8BF); it stops at D8BF so emoji
// (D83C-D83E high surrogates) and the planes 4-16 private-use area and
// variation selectors supplement (D8C0-DBFF) are never counted as Asian chars.
const ASIAN_RE =
  /[ᄀ-ᇿ⺀-⿟、-〿぀-ヿ㄀-ㄯ㄰-㆏㇀-ㇿ㐀-䶿一-鿿가-힯豈-﫿！-｠￠-￦]|[\uD840-\uD8BF][\uDC00-\uDFFF]/g

// Latin Extended Additional (Vietnamese, …) plus the space-delimited
// non-Latin scripts Word counts as words: Greek, Cyrillic, Hebrew, Arabic,
// Armenian, Devanagari, Thai, Lao, Tibetan, Myanmar, Georgian, Khmer.
// Deliberately excludes the ASIAN_RE ranges above (counted char-by-char).
/* eslint-disable no-misleading-character-class -- Thai/Lao etc ranges contain combining marks but are intentional word characters */
const NON_ASIAN_WORD_RE =
  /[A-Za-z0-9À-ɏͰ-ϿЀ-ӿ֐-׿؀-ۿḀ-ỿ\u0E00-\u0EFF\u1000-\u109F\u10A0-\u10FF\u1780-\u17FF\u0530-\u058F\u0900-\u097F\u0F00-\u0FFF]+(?:['-][A-Za-z0-9À-ɏͰ-ϿЀ-ӿ֐-׿؀-ۿḀ-ỿ\u0E00-\u0EFF\u1000-\u109F\u10A0-\u10FF\u1780-\u17FF\u0530-\u058F\u0900-\u097F\u0F00-\u0FFF]+)*/g
/* eslint-enable no-misleading-character-class */

export function asianCharCount(text: string): number {
  return (text.match(ASIAN_RE) ?? []).length
}

export function nonAsianWordCount(text: string): number {
  return (text.match(NON_ASIAN_WORD_RE) ?? []).length
}

/** Word's Words figure: asian chars + non-asian words */
export function countWords(text: string): number {
  return asianCharCount(text) + nonAsianWordCount(text)
}
