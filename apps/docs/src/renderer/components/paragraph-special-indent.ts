/** Word's Paragraph ▸ Special dropdown over the single w:ind firstLine value
 * (positive = first line, negative = hanging, as the editor stores it). */
export type SpecialIndent = 'none' | 'firstLine' | 'hanging'

/** unit of the "By" field: an absolute length, or Word's character unit for
 * East Asian text (w:firstLineChars, hundredths of a character, issue #1892) */
export type SpecialByUnit = 'cm' | 'chars'

export interface SpecialState {
  special: SpecialIndent
  /** twips when unit is cm (or undefined); hundredths of a character when unit is 'chars' */
  by: number
  /** undefined = the absolute length unit; only 'chars' is stored explicitly so
   *  the twips-only shape stays backward compatible */
  unit?: SpecialByUnit
}

/** Word's default "By" when switching from (none): 0.5" */
export const DEFAULT_SPECIAL_BY = 720

/** Word's default character-unit "By" for East Asian text: two characters */
export const DEFAULT_SPECIAL_BY_CHARS = 200

export function specialFromFirstLine(firstLine: unknown, firstLineChars?: unknown): SpecialState {
  const chars = Number(firstLineChars)
  if (Number.isFinite(chars) && chars > 0) return { special: 'firstLine', by: chars, unit: 'chars' }
  const v = Number(firstLine) || 0
  if (v > 0) return { special: 'firstLine', by: v }
  if (v < 0) return { special: 'hanging', by: -v }
  return { special: 'none', by: 0 }
}

export function firstLineFromSpecial({ special, by }: SpecialState): number | null {
  if (special === 'none' || !(by > 0)) return null
  return special === 'hanging' ? -by : by
}

/**
 * The paragraph attrs the dialog writes for the Special state. The character
 * unit rides along its resolved twips twin: the twin feeds the ruler and the
 * w:firstLine attribute Word falls back to, while w:firstLineChars stays
 * authoritative and rescales with the first run's font size.
 */
export function specialIndentAttrs(
  { special, by, unit }: SpecialState,
  charUnitTwips: number,
): { indentFirstLine: number | null; indentFirstLineChars: number | null } {
  if (special === 'none' || !(by > 0)) return { indentFirstLine: null, indentFirstLineChars: null }
  if (special === 'hanging') return { indentFirstLine: -by, indentFirstLineChars: null }
  if (unit === 'chars')
    return { indentFirstLineChars: by, indentFirstLine: Math.round((by / 100) * charUnitTwips) }
  return { indentFirstLine: by, indentFirstLineChars: null }
}

/** Picking a kind keeps the By value; leaving (none) seeds Word's 0.5" (or the
 *  character-unit default when the field is already in character units). The
 *  character unit is a first-line-indent unit only (w:hangingChars is not
 *  modelled): a hanging pick drops to the absolute default. */
export function pickSpecial(prev: SpecialState, special: SpecialIndent): SpecialState {
  if (special === 'none') return { special, by: 0 }
  if (special === 'hanging') {
    if (prev.unit === 'chars') return { special, by: DEFAULT_SPECIAL_BY }
    return { special, by: prev.by > 0 ? prev.by : DEFAULT_SPECIAL_BY }
  }
  const seeded = prev.unit === 'chars' ? DEFAULT_SPECIAL_BY_CHARS : DEFAULT_SPECIAL_BY
  return {
    special,
    by: prev.by > 0 ? prev.by : seeded,
    ...(prev.unit === 'chars' ? { unit: 'chars' as const } : {}),
  }
}

/** Switching the By unit seeds the unit's default (2 characters / 0.5") */
export function pickByUnit(prev: SpecialState, unit: SpecialByUnit): SpecialState {
  if (unit === 'chars')
    return { special: prev.special, by: DEFAULT_SPECIAL_BY_CHARS, unit: 'chars' }
  return { special: prev.special, by: DEFAULT_SPECIAL_BY }
}
