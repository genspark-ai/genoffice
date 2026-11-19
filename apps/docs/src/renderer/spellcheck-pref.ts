/** Persisted "check spelling as you type" toggle (Review → Spelling).
 * Drives the native Chromium checker's red squiggle on the document body and
 * the header/footer edit surfaces; on by default, matching Word. */
export const SPELLCHECK_KEY = 'aidocs.spellcheck'

export function spellcheckEnabled(): boolean {
  return localStorage.getItem(SPELLCHECK_KEY) !== '0'
}
