import type { Lang } from '@genoffice/i18n'
import { en } from './app/en'
import type { zh } from './app/zh'

/** Strings for App.tsx / SettingsModal / non-Ribbon components (keys use the app prefix to mark the area) */
export type AppDict = typeof zh

/** en stays eager so t() has a fallback before any lazy locale chunk lands */
export const appEn: AppDict = en

const loaders: Record<Lang, () => Promise<AppDict>> = {
  en: () => Promise.resolve(appEn),
  zh: async () => (await import('./app/zh')).zh,
  ja: async () => (await import('./app/ja')).ja,
  ko: async () => (await import('./app/ko')).ko,
  fr: async () => (await import('./app/fr')).fr,
  de: async () => (await import('./app/de')).de,
  es: async () => (await import('./app/es')).es,
  th: async () => (await import('./app/th')).th,
  id: async () => (await import('./app/id')).id,
  ru: async () => (await import('./app/ru')).ru,
  ar: async () => (await import('./app/ar')).ar,
  pt: async () => (await import('./app/pt')).pt,
  it: async () => (await import('./app/it')).it,
  pl: async () => (await import('./app/pl')).pl,
  cs: async () => (await import('./app/cs')).cs,
  nl: async () => (await import('./app/nl')).nl,
  ms: async () => (await import('./app/ms')).ms,
  he: async () => (await import('./app/he')).he,
  hi: async () => (await import('./app/hi')).hi,
  'zh-TW': async () => (await import('./app/zh-TW')).zhTW,
  vi: async () => (await import('./app/vi')).vi,
}

export function loadAppStrings(lang: Lang): Promise<AppDict> {
  return loaders[lang]()
}
