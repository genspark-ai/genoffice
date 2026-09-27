/** Persisted AutoCorrect switches (Tools → AutoCorrect Options…), read on every keystroke. */
export const AUTOCORRECT_KEY = 'aidocs.autocorrect'

export const AUTOCORRECT_RULES = [
  'smartQuotes',
  'dashes',
  'autoLists',
  'symbols',
  'ordinals',
  'capitalize',
] as const

export type AutocorrectRule = (typeof AUTOCORRECT_RULES)[number]

// Word's defaults, except sentence capitalization: it misfires on CJK-mixed
// and technical text often enough that we leave it opt-in
export const AUTOCORRECT_DEFAULTS: Record<AutocorrectRule, boolean> = {
  smartQuotes: true,
  dashes: true,
  autoLists: true,
  symbols: true,
  ordinals: true,
  capitalize: false,
}

function readStored(): Partial<Record<AutocorrectRule, boolean>> {
  try {
    const raw = globalThis.localStorage?.getItem(AUTOCORRECT_KEY)
    if (!raw) return {}
    const parsed: unknown = JSON.parse(raw)
    return parsed && typeof parsed === 'object' ? (parsed as Record<AutocorrectRule, boolean>) : {}
  } catch {
    return {}
  }
}

export function autocorrectPrefs(): Record<AutocorrectRule, boolean> {
  const stored = readStored()
  const out = { ...AUTOCORRECT_DEFAULTS }
  for (const rule of AUTOCORRECT_RULES) {
    if (typeof stored[rule] === 'boolean') out[rule] = stored[rule]
  }
  return out
}

export function autocorrectEnabled(rule: AutocorrectRule): boolean {
  return autocorrectPrefs()[rule]
}

export function setAutocorrectPref(rule: AutocorrectRule, on: boolean): void {
  const next = { ...readStored(), [rule]: on }
  try {
    globalThis.localStorage?.setItem(AUTOCORRECT_KEY, JSON.stringify(next))
  } catch {
    /* storage unavailable: the switch just doesn't persist */
  }
}

export function resetAutocorrectPrefs(): void {
  try {
    globalThis.localStorage?.removeItem(AUTOCORRECT_KEY)
  } catch {
    /* ignore */
  }
}
