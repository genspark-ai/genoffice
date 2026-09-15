import type { LanguageCode, LanguageOption } from './types'

/**
 * Languages the AI translation UI offers.
 *
 * Labels are in the option's own language so the picker reads naturally for
 * every user; the English label travels into the prompt as the canonical name
 * the LLM should target. Keep this list in lock-step with the language select
 * in `packages/ui/src/TranslateDialog.tsx`.
 */
export const LANGUAGES: readonly LanguageOption[] = [
  { value: 'auto', label: '自动检测', englishLabel: 'Auto-detect' },
  { value: 'zh-CN', label: '简体中文', englishLabel: 'Simplified Chinese' },
  { value: 'zh-TW', label: '繁體中文', englishLabel: 'Traditional Chinese' },
  { value: 'en-US', label: 'English', englishLabel: 'English' },
  { value: 'ja-JP', label: '日本語', englishLabel: 'Japanese' },
  { value: 'ko-KR', label: '한국어', englishLabel: 'Korean' },
  { value: 'fr-FR', label: 'Français', englishLabel: 'French' },
  { value: 'de-DE', label: 'Deutsch', englishLabel: 'German' },
  { value: 'es-ES', label: 'Español', englishLabel: 'Spanish' },
  { value: 'it-IT', label: 'Italiano', englishLabel: 'Italian' },
  { value: 'pt-PT', label: 'Português', englishLabel: 'Portuguese' },
  { value: 'ru-RU', label: 'Русский', englishLabel: 'Russian' },
  { value: 'ar-SA', label: 'العربية', englishLabel: 'Arabic' },
  { value: 'hi-IN', label: 'हिन्दी', englishLabel: 'Hindi' },
  { value: 'th-TH', label: 'ไทย', englishLabel: 'Thai' },
] as const

const BY_VALUE = new Map<LanguageCode, LanguageOption>(LANGUAGES.map((opt) => [opt.value, opt]))

function familyOf(value: string): string {
  const idx = value.indexOf('-')
  return (idx >= 0 ? value.slice(0, idx) : value).toLowerCase()
}

/** Look up a language by BCP-47 code; falls back to an English-label guess. */
export function getLanguage(value: string | undefined | null): LanguageOption | null {
  if (!value) return null
  const direct = BY_VALUE.get(value as LanguageCode)
  if (direct) return direct
  // tolerate stripped region codes (e.g. 'en' from a foreign UI)
  const family = familyOf(value)
  for (const opt of LANGUAGES) {
    if (opt.value === 'auto') continue
    if (familyOf(opt.value) === family) return opt
  }
  return null
}

/** Canonical name for prompts: prefers the English label, falls back to the code. */
export function englishLabelFor(value: string | undefined | null): string {
  const lang = getLanguage(value)
  if (lang) return lang.englishLabel
  if (!value || value === 'auto') return 'Auto-detect'
  // last resort — pass the code through so the model still gets a target
  return value
}
