export type LineSpacingChoice = 'single' | 'oneHalf' | 'double' | 'multiple' | 'atLeast' | 'exact'

export interface LineSpacingState {
  choice: LineSpacingChoice
  /** multiple of single spacing, 2 decimals (the "At" box under Multiple) */
  multiple: number
  /** pt value under At least / Exactly */
  pt: number
}

const FIXED_MULTIPLES: Partial<Record<number, LineSpacingChoice>> = {
  1: 'single',
  1.5: 'oneHalf',
  2: 'double',
}

/** Word's default "At" value when Multiple is picked from a fixed choice */
export const DEFAULT_MULTIPLE = 3

export function lineSpacingFromAttrs(attrs: Record<string, unknown>): LineSpacingState {
  const rawTwips = Number(attrs.lineRawTwips) || 0
  const pt = rawTwips ? Math.round(rawTwips / 2) / 10 : 12
  if (attrs.lineRule === 'exact' || attrs.lineRule === 'atLeast')
    return { choice: attrs.lineRule, multiple: 1, pt }
  const raw =
    Number(attrs.lineSpacing) || (attrs.lineRule === 'auto' && rawTwips ? rawTwips / 240 : 1)
  const multiple = Math.round(raw * 100) / 100
  return { choice: FIXED_MULTIPLES[multiple] ?? 'multiple', multiple, pt }
}

export function pickLineSpacing(
  prev: LineSpacingState,
  choice: LineSpacingChoice,
): LineSpacingState {
  if (choice === 'multiple' && FIXED_MULTIPLES[prev.multiple])
    return { ...prev, choice, multiple: DEFAULT_MULTIPLE }
  return { ...prev, choice }
}

export function lineSpacingAttrs(state: LineSpacingState): {
  lineSpacing: number | null
  lineRule: 'exact' | 'atLeast' | null
  lineRawTwips: number | null
} {
  const { choice } = state
  if (choice === 'exact' || choice === 'atLeast')
    return {
      lineSpacing: null,
      lineRule: choice,
      lineRawTwips: Math.max(20, Math.round(state.pt * 20)),
    }
  const multiple =
    choice === 'single' ? 1 : choice === 'oneHalf' ? 1.5 : choice === 'double' ? 2 : state.multiple
  return { lineSpacing: multiple === 1 ? null : multiple, lineRule: null, lineRawTwips: null }
}
