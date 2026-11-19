/** Tools > Translate: the language text is translated into. */
import { beforeEach, describe, expect, it } from 'vitest'
import { setTranslationLanguage, translationLanguage } from '../src/renderer/translate-lang'

describe('translation language', () => {
  beforeEach(() => localStorage.clear())

  it('starts as the UI language, or English when that is not a language we translate into', () => {
    expect(translationLanguage('ja')).toBe('ja')
    expect(translationLanguage('xx')).toBe('en')
  })

  it('keeps the language that was set, whatever the UI language', () => {
    setTranslationLanguage('fr')
    expect(translationLanguage('ja')).toBe('fr')
  })

  it('ignores a stored value that is not a language', () => {
    localStorage.setItem('aidocs.translateLang', 'klingon')
    expect(translationLanguage('de')).toBe('de')
  })
})
