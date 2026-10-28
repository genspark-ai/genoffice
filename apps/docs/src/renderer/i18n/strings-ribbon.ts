import type { Lang } from '@genoffice/i18n'
import { en } from './ribbon/en'
import type { zh } from './ribbon/zh'

/** Strings for Ribbon / ribbon-tabs / icons */
export type RibbonDict = typeof zh

/** en stays eager so t() has a fallback before any lazy locale chunk lands */
export const ribbonEn: RibbonDict = en

const loaders: Record<Lang, () => Promise<RibbonDict>> = {
  en: () => Promise.resolve(ribbonEn),
  zh: async () => (await import('./ribbon/zh')).zh,
  ja: async () => (await import('./ribbon/ja')).ja,
  ko: async () => (await import('./ribbon/ko')).ko,
  fr: async () => (await import('./ribbon/fr')).fr,
  de: async () => (await import('./ribbon/de')).de,
  es: async () => (await import('./ribbon/es')).es,
  th: async () => (await import('./ribbon/th')).th,
  id: async () => (await import('./ribbon/id')).id,
  ru: async () => (await import('./ribbon/ru')).ru,
  ar: async () => (await import('./ribbon/ar')).ar,
  pt: async () => (await import('./ribbon/pt')).pt,
  it: async () => (await import('./ribbon/it')).it,
  pl: async () => (await import('./ribbon/pl')).pl,
  cs: async () => (await import('./ribbon/cs')).cs,
  nl: async () => (await import('./ribbon/nl')).nl,
  ms: async () => (await import('./ribbon/ms')).ms,
  he: async () => (await import('./ribbon/he')).he,
  hi: async () => (await import('./ribbon/hi')).hi,
  'zh-TW': async () => (await import('./ribbon/zh-TW')).zhTW,
  vi: async () => (await import('./ribbon/vi')).vi,
}

export function loadRibbonStrings(lang: Lang): Promise<RibbonDict> {
  return loaders[lang]()
}
