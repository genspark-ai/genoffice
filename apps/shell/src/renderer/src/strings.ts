import { format, platformShortcuts, type Lang, type Params } from '@genoffice/i18n'
import type { Dict, StringKey } from './strings/zh'

export type { Dict, StringKey }

/** one dynamic import per language: the entry chunk ships no dictionary */
const LOADERS: Record<Lang, () => Promise<Dict>> = {
  zh: () => import('./strings/zh').then((m) => m.zh),
  en: () => import('./strings/en').then((m) => m.en),
  vi: () => import('./strings/vi').then((m) => m.vi),
  ja: () => import('./strings/ja').then((m) => m.ja),
  ko: () => import('./strings/ko').then((m) => m.ko),
  fr: () => import('./strings/fr').then((m) => m.fr),
  de: () => import('./strings/de').then((m) => m.de),
  es: () => import('./strings/es').then((m) => m.es),
  th: () => import('./strings/th').then((m) => m.th),
  id: () => import('./strings/id').then((m) => m.id),
  ru: () => import('./strings/ru').then((m) => m.ru),
  ar: () => import('./strings/ar').then((m) => m.ar),
  pt: () => import('./strings/pt').then((m) => m.pt),
  it: () => import('./strings/it').then((m) => m.it),
  pl: () => import('./strings/pl').then((m) => m.pl),
  cs: () => import('./strings/cs').then((m) => m.cs),
  nl: () => import('./strings/nl').then((m) => m.nl),
  ms: () => import('./strings/ms').then((m) => m.ms),
  he: () => import('./strings/he').then((m) => m.he),
  hi: () => import('./strings/hi').then((m) => m.hi),
  'zh-TW': () => import('./strings/zh-TW').then((m) => m.zhTW),
}

const loaded = new Map<Lang, Dict>()
const pending = new Map<Lang, Promise<void>>()

/** Fetch one language's dictionary (idempotent); a language whose chunk cannot
 *  be loaded falls back to en. */
export function loadStrings(lang: Lang): Promise<void> {
  if (loaded.has(lang)) return Promise.resolve()
  let request = pending.get(lang)
  if (!request) {
    request = LOADERS[lang]()
      .then((dict) => {
        loaded.set(lang, dict)
      })
      .catch(async (err: unknown) => {
        console.warn(`dictionary for ${lang} unavailable`, err)
        if (lang !== 'en' && !loaded.has('en')) await loadStrings('en').catch(() => undefined)
      })
    pending.set(lang, request)
  }
  return request
}

/** Load every dictionary; tests read the whole table, the app only one. */
export async function loadAllStrings(): Promise<Record<Lang, Dict>> {
  await Promise.all(Object.keys(LOADERS).map((lang) => loadStrings(lang as Lang)))
  return Object.fromEntries(loaded) as Record<Lang, Dict>
}

/** Synchronous translate over the loaded dictionaries; the fallback chain
 *  (requested language, then en, then zh, then the key) matches createI18n(). */
export function translate(lang: Lang, key: StringKey, params?: Params): string {
  const text = loaded.get(lang)?.[key] ?? loaded.get('en')?.[key] ?? loaded.get('zh')?.[key] ?? key
  return format(platformShortcuts(text), params)
}
