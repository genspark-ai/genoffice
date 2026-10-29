/**
 * Sheets spelling preferences and dictionaries (Review → Proofing).
 *
 * The F7 / Review ▸ Spelling dialog checks cell texts against the frame's own
 * spell engine (webFrame.isWordMisspelled / getWordSuggestions — Chromium
 * Hunspell / the OS dictionary, the same checker docs' squiggles run on), so
 * no cell text leaves the machine. This module keeps the two word lists that
 * drive its white-listing, in docs' preference shape:
 *
 * - the on/off preference: localStorage, default on (docs: aidocs.spellcheck)
 * - "Add to Dictionary": persisted across sessions here (the dialog's
 *   white-list; a future wiring could feed the OS custom dictionary instead)
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
