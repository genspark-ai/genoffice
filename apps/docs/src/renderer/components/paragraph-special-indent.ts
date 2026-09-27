/** Word's Paragraph ▸ Special dropdown over the single w:ind firstLine value
 * (positive = first line, negative = hanging, as the editor stores it). */
export type SpecialIndent = 'none' | 'firstLine' | 'hanging'

export interface SpecialState {
  special: SpecialIndent
  /** twips, always ≥ 0 (the "By" field) */
  by: number
}

/** Word's default "By" when switching from (none): 0.5" */
export const DEFAULT_SPECIAL_BY = 720

export function specialFromFirstLine(firstLine: unknown): SpecialState {
  const v = Number(firstLine) || 0
  if (v > 0) return { special: 'firstLine', by: v }
  if (v < 0) return { special: 'hanging', by: -v }
  return { special: 'none', by: 0 }
}

export function firstLineFromSpecial({ special, by }: SpecialState): number | null {
  if (special === 'none' || !(by > 0)) return null
  return special === 'hanging' ? -by : by
}

/** Picking a kind keeps the By value; leaving (none) seeds Word's 0.5". */
export function pickSpecial(prev: SpecialState, special: SpecialIndent): SpecialState {
  if (special === 'none') return { special, by: 0 }
  return { special, by: prev.by > 0 ? prev.by : DEFAULT_SPECIAL_BY }
}
