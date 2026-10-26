/**
 * Sheets spelling preferences and dictionaries (Review → Proofing).
 *
 * Docs runs its spellcheck on the Chromium native checker (squiggle + native
 * right-click suggestions on its contenteditable body). Sheets' Univer cell
 * editor is canvas-rendered — its only DOM editable is a hidden 1px proxy
 * whose content Univer clears on every input event — so native squiggles and
 * native suggestions are architecturally unavailable there, and Electron
 * exposes no batch "check this text" API against the native checker. The F7 /
 * Review ▸ Spelling dialog therefore checks cell texts through the AI backend
 * (the same route docs' F7 proofread takes), and this module keeps the two
 * word lists that drive its white-listing, in docs' preference shape:
 *
 * - the on/off preference: localStorage, default on (docs: aidocs.spellcheck)
 * - "Add to Dictionary": persisted across sessions (docs writes the OS custom
 *   dictionary; the AI checker has no OS dictionary, so the list lives here)
 * - "Ignore All": session-only, like docs' spell-ignored map that is dropped
 *   when the document closes
 */
export const SPELLCHECK_KEY = 'aisheets.spellcheck'
export const SPELL_DICTIONARY_KEY = 'aisheets.spell-dictionary'

export type StringStorage = Pick<Storage, 'getItem' | 'setItem'>

export function spellcheckEnabled(storage: StringStorage = globalThis.localStorage): boolean {
  return storage.getItem(SPELLCHECK_KEY) !== '0'
}

export function storeSpellcheckEnabled(
  enabled: boolean,
  storage: StringStorage = globalThis.localStorage,
): void {
  storage.setItem(SPELLCHECK_KEY, enabled ? '1' : '0')
}

/** Upper cap so a decade of "Add to Dictionary" clicks cannot grow the key unbounded. */
const DICTIONARY_CAP = 500

export function loadLearnedWords(storage: StringStorage = globalThis.localStorage): string[] {
  try {
    const raw = storage.getItem(SPELL_DICTIONARY_KEY)
    if (!raw) return []
    const parsed: unknown = JSON.parse(raw)
    if (!Array.isArray(parsed)) return []
    return parsed.filter((word): word is string => typeof word === 'string' && word.length > 0)
  } catch {
    return []
  }
}

export function saveLearnedWords(
  words: readonly string[],
  storage: StringStorage = globalThis.localStorage,
): void {
  // most recent wins when the cap is hit; the tail holds the oldest entries
  const trimmed = [...new Set(words)].slice(-DICTIONARY_CAP)
  try {
    storage.setItem(SPELL_DICTIONARY_KEY, JSON.stringify(trimmed))
  } catch {
    // storage full or unavailable: this session keeps the words in memory
  }
}
