import { TRANSLATE_LANGS, type TranslateLang } from './components/translate-langs'

const KEY = 'aidocs.translateLang'

const known = (code: unknown): TranslateLang | null =>
  TRANSLATE_LANGS.find((l) => l.code === code)?.code ?? null

/** the language Tools > Translate turns text into: the one set, else the UI language, else English */
export function translationLanguage(uiLang: string): TranslateLang {
  let stored: string | null = null
  try {
    stored = localStorage.getItem(KEY)
  } catch {
    // storage disabled: fall through to the UI language
  }
  return known(stored) ?? known(uiLang) ?? 'en'
}

export function setTranslationLanguage(code: TranslateLang): void {
  try {
    localStorage.setItem(KEY, code)
  } catch {
    // the choice lasts for nothing then, and translating still works
  }
}
