// Script and face-name tests the emitters used to make next to their output:
// what counts as East Asian text or an East Asian face for a layout decision.
import { isEastAsianFontName } from '../font-list'
import { isCjkFontName, isKoreanFontName, textHasCjk } from '../line-metrics'

/** the body holds CJK text (drives the document-level line-height factor) */
export function blocksHaveCjkText(
  blocks: ReadonlyArray<{ hidden?: boolean; runs?: ReadonlyArray<{ text: string }> }>,
): boolean {
  return blocks.some((b) => !b.hidden && (b.runs ?? []).some((r) => textHasCjk(r.text)))
}

/** a CJK face that is not Korean: hangul under it takes Word's fallback face height */
export const takesHangulFallbackHeight = (font: string): boolean =>
  isCjkFontName(font) && !isKoreanFontName(font)

/** the font of the spaces leading a picture is East Asian (they are full-width cells) */
export const isEastAsianLeadingFont = (font: string): boolean => isCjkFontName(font)

/** the East Asian and the Latin family a CSS family list stands for; the Simplified
 *  Noto CJK tail only counts as the East Asian one when it leads */
export function eastAsianAndLatinOf(families: string[]): {
  ea: string | undefined
  latin: string | undefined
} {
  const ea = families.find(
    (f, i) => isEastAsianFontName(f) && (i === 0 || !/^noto (sans|serif) cjk sc$/i.test(f)),
  )
  const latin = families.find((f) => !isEastAsianFontName(f))
  return { ea, latin }
}

/** inner whitespace or CJK text: a wrap opportunity inside a tab segment */
export const SEG_BREAKABLE_RE = /\s|[\u2E80-\u9FFF\uAC00-\uD7AF\uF900-\uFAFF\uFF00-\uFFEF]/

/** halfwidth kana / hangul / symbols: East Asian face, single-byte for the balanced
 *  character grid (Word probe 2026-10-22) */
export const HALFWIDTH_FORM_RE = /[\uff61-\uffdc\uffe8-\uffee]/u
