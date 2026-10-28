import type { Lang } from '@genoffice/i18n'
import { en } from './ai/en'
import type { zh } from './ai/zh'

/** User-visible strings for the ai/ panel and tool feedback (LLM prompts are not here) */
export type AiDict = typeof zh

/** en stays eager so t() has a fallback before any lazy locale chunk lands */
export const aiEn: AiDict = en

const loaders: Record<Lang, () => Promise<AiDict>> = {
  en: () => Promise.resolve(aiEn),
  zh: async () => (await import('./ai/zh')).zh,
  ja: async () => (await import('./ai/ja')).ja,
  ko: async () => (await import('./ai/ko')).ko,
  fr: async () => (await import('./ai/fr')).fr,
  de: async () => (await import('./ai/de')).de,
  es: async () => (await import('./ai/es')).es,
  th: async () => (await import('./ai/th')).th,
  id: async () => (await import('./ai/id')).id,
  ru: async () => (await import('./ai/ru')).ru,
  ar: async () => (await import('./ai/ar')).ar,
  pt: async () => (await import('./ai/pt')).pt,
  it: async () => (await import('./ai/it')).it,
  pl: async () => (await import('./ai/pl')).pl,
  cs: async () => (await import('./ai/cs')).cs,
  nl: async () => (await import('./ai/nl')).nl,
  ms: async () => (await import('./ai/ms')).ms,
  he: async () => (await import('./ai/he')).he,
  hi: async () => (await import('./ai/hi')).hi,
  'zh-TW': async () => (await import('./ai/zh-TW')).zhTW,
  vi: async () => (await import('./ai/vi')).vi,
}

export function loadAiStrings(lang: Lang): Promise<AiDict> {
  return loaders[lang]()
}
