import type { StringKey } from '../i18n/locale'

/**
 * Languages the AI backend can translate into, in the order both dropdowns
 * list them.
 *
 * The right-click menu and the Review → Translate dropdown used to carry their
 * own copies of this list, which is how the two drifted: adding a language meant
 * editing two arrays, and a row added to one never reached the other. They now
 * share this one.
 *
 * The code → i18n mapping is written out rather than derived, because the
 * existing keys are spelled out in full (`appLangSimplifiedChinese`, not
 * `appLangZh`) and a rule that "looks right" would silently miss all seven that
 * ship today. `zh-TW` is also not a bare identifier, so a join-then-capitalise
 * approach cannot express it.
 */
export const TRANSLATE_LANGS = [
  { code: 'en', key: 'English' },
  { code: 'zh', key: 'SimplifiedChinese' },
  { code: 'ja', key: 'Japanese' },
  { code: 'ko', key: 'Korean' },
  { code: 'fr', key: 'French' },
  { code: 'de', key: 'German' },
  { code: 'es', key: 'Spanish' },
  { code: 'th', key: 'Thai' },
  { code: 'id', key: 'Indonesian' },
  { code: 'ru', key: 'Russian' },
  { code: 'ar', key: 'Arabic' },
  { code: 'pt', key: 'Portuguese' },
  { code: 'it', key: 'Italian' },
  { code: 'pl', key: 'Polish' },
  { code: 'cs', key: 'Czech' },
  { code: 'nl', key: 'Dutch' },
  { code: 'ms', key: 'Malay' },
  { code: 'he', key: 'Hebrew' },
  { code: 'hi', key: 'Hindi' },
  { code: 'zh-TW', key: 'TraditionalChinese' },
  { code: 'vi', key: 'Vietnamese' },
] as const

export type TranslateLang = (typeof TRANSLATE_LANGS)[number]['code']

/** The right-click menu's i18n key for a language. */
export function appLangKey(code: TranslateLang): StringKey {
  const entry = TRANSLATE_LANGS.find((l) => l.code === code)
  return `appLang${entry?.key ?? ''}` as StringKey
}

/** The ribbon dropdown's i18n key for a language. */
export function ribbonLangKey(code: TranslateLang): StringKey {
  const entry = TRANSLATE_LANGS.find((l) => l.code === code)
  return `ribbonLang${entry?.key ?? ''}` as StringKey
}
