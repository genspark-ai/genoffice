/** Tools > Spelling and Grammar > Hide Grammar Errors, and the dialog's Check grammar box. */
export const HIDE_GRAMMAR_KEY = 'aidocs.hideGrammar'
export const CHECK_GRAMMAR_KEY = 'aidocs.checkGrammar'

/** false where there is no storage: the editor also runs headless, in the CLI */
export const grammarHidden = (): boolean =>
  typeof localStorage !== 'undefined' && localStorage.getItem(HIDE_GRAMMAR_KEY) === '1'
export const setGrammarHiddenPref = (hidden: boolean): void =>
  localStorage.setItem(HIDE_GRAMMAR_KEY, hidden ? '1' : '0')

/** off until ticked: a grammar check sends the document's text to the AI provider */
export const checkGrammarEnabled = (): boolean => localStorage.getItem(CHECK_GRAMMAR_KEY) === '1'
export const setCheckGrammarPref = (on: boolean): void =>
  localStorage.setItem(CHECK_GRAMMAR_KEY, on ? '1' : '0')
