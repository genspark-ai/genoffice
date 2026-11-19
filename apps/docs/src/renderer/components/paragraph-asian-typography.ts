export const TEXT_ALIGNMENT_VALUES = ['top', 'center', 'baseline', 'bottom', 'auto'] as const
export type TextAlignment = (typeof TEXT_ALIGNMENT_VALUES)[number]

export interface AsianTypographyState {
  textAlignment: TextAlignment
  kinsoku: boolean
  topLinePunct: boolean
  latinWrap: boolean
  hangingPunct: boolean
  autoSpaceLatin: boolean
  autoSpaceDigits: boolean
}

/** the effective values the page lays the paragraph out with (Word's defaults: all on) */
export function asianTypographyFromAttrs(attrs: Record<string, unknown>): AsianTypographyState {
  const mode = attrs.autoSpace
  const declared = TEXT_ALIGNMENT_VALUES.find((v) => v === attrs.textAlignment)
  return {
    textAlignment: declared ?? 'auto',
    kinsoku: attrs.kinsoku !== false,
    topLinePunct: attrs.topLinePunct === true,
    latinWrap: attrs.wordWrap === false,
    hangingPunct: attrs.overflowPunct !== false,
    autoSpaceLatin: mode == null || mode === true || mode === 'latin',
    autoSpaceDigits: mode == null || mode === true || mode === 'digits',
  }
}

/** only a changed checkbox becomes direct pPr; the untouched ones keep following the style */
export function asianTypographyAttrs(
  state: AsianTypographyState,
  initial: AsianTypographyState,
): Record<string, unknown> {
  const out: Record<string, unknown> = {}
  // Auto is what a paragraph without the element has: it takes the element away
  if (state.textAlignment !== initial.textAlignment)
    out.textAlignment = state.textAlignment === 'auto' ? null : state.textAlignment
  if (state.kinsoku !== initial.kinsoku) out.kinsoku = state.kinsoku
  if (state.topLinePunct !== initial.topLinePunct) out.topLinePunct = state.topLinePunct
  if (state.latinWrap !== initial.latinWrap) {
    out.wordWrap = !state.latinWrap
    out.wordWrapDirect = !state.latinWrap
  }
  if (state.hangingPunct !== initial.hangingPunct) {
    out.overflowPunct = state.hangingPunct
    out.overflowPunctDirect = state.hangingPunct
  }
  const latin = state.autoSpaceLatin !== initial.autoSpaceLatin
  const digits = state.autoSpaceDigits !== initial.autoSpaceDigits
  if (latin) out.autoSpaceDE = state.autoSpaceLatin
  if (digits) out.autoSpaceDN = state.autoSpaceDigits
  if (latin || digits) {
    out.autoSpace =
      state.autoSpaceLatin && state.autoSpaceDigits
        ? true
        : state.autoSpaceLatin
          ? 'latin'
          : state.autoSpaceDigits
            ? 'digits'
            : false
  }
  return out
}
