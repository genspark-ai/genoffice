import { symbolGlyph, symbolPuaChar } from '@genoffice/docx-engine'
import { symbolFontCovers } from '../font-check'

/** the symbol fonts Word's Symbol dialog lists, whose characters save as w:sym */
export const SYMBOL_FONTS = ['Symbol', 'Webdings', 'Wingdings', 'Wingdings 2', 'Wingdings 3']

const FIRST_CODE = 0x21
const LAST_CODE = 0xff

/** a symbol font's character as w:sym names it: the F0xx form of its one-byte code */
export const fontSymbolChar = (code: number): string =>
  (0xf000 + (code & 0xff)).toString(16).toUpperCase()

/**
 * What draws a symbol font's character here: the font's own glyph when it is
 * installed, else the Unicode stand-in. Null when neither can draw it.
 */
export function fontSymbolDisplay(
  font: string,
  char: string,
): { text: string; ownFont: boolean } | null {
  const pua = symbolPuaChar(char)
  if (pua && symbolFontCovers(font, pua)) return { text: pua, ownFont: true }
  const glyph = symbolGlyph(font, char)
  return glyph && glyph !== pua ? { text: glyph, ownFont: false } : null
}

/** the characters of a symbol font that can be drawn on this machine, as w:sym chars */
export function fontSymbolChars(font: string): string[] {
  const chars: string[] = []
  for (let code = FIRST_CODE; code <= LAST_CODE; code++) {
    const char = fontSymbolChar(code)
    if (code !== 0x7f && fontSymbolDisplay(font, char)) chars.push(char)
  }
  return chars
}

/** editor content for one symbol-font character: its display text under the mark that saves it as w:sym */
export function fontSymbolContent(font: string, char: string) {
  const shown = fontSymbolDisplay(font, char)
  if (!shown) return null
  return { type: 'text', text: shown.text, marks: [{ type: 'docSym', attrs: { font, char } }] }
}

const SEPARATOR = '\t'

export const fontSymbolPayload = (font: string, char: string): string =>
  `${font}${SEPARATOR}${char}`

export function parseFontSymbolPayload(payload: string): { font: string; char: string } | null {
  const [font, char] = payload.split(SEPARATOR)
  return font && /^[0-9A-F]{4}$/i.test(char ?? '') ? { font, char } : null
}
