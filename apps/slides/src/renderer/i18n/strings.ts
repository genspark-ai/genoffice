import { createI18n, type Lang, type LangDicts } from '@genoffice/i18n'
import { zh as appZh } from './app/zh'
import { en as appEn } from './app/en'
import { zh as ribbonZh } from './ribbon/zh'
import { en as ribbonEn } from './ribbon/en'
import { zh as paneZh } from './panes/zh'
import { en as paneEn } from './panes/en'
import { zh as aiZh } from './ai/zh'
import { en as aiEn } from './ai/en'

const zhDict = { ...appZh, ...ribbonZh, ...paneZh, ...aiZh }
const enDict = { ...appEn, ...ribbonEn, ...paneEn, ...aiEn }

/** the merged UI dictionary; zh defines the key set every locale chunk must match */
export type Dict = typeof zhDict
export type StringKey = keyof Dict

// zh (default UI language) and en (the translator's fallback) ship in the entry
// chunk; the other 19 locales load on demand via loadLocale().
const dicts = { zh: zhDict, en: enDict } as LangDicts<Dict>

/** translate with the createI18n fallback chain: lang → en → zh → key */
export const translate = createI18n(dicts)

type LazyLang = Exclude<Lang, 'zh' | 'en'>

// one dynamic import per locale: rollup emits each locale's shards as their own chunk
const LOADERS: { [L in LazyLang]: () => Promise<Dict> } = {
  ja: () => import('./locales/ja').then((m) => m.ja),
  ko: () => import('./locales/ko').then((m) => m.ko),
  fr: () => import('./locales/fr').then((m) => m.fr),
  de: () => import('./locales/de').then((m) => m.de),
  es: () => import('./locales/es').then((m) => m.es),
  th: () => import('./locales/th').then((m) => m.th),
  id: () => import('./locales/id').then((m) => m.id),
  ru: () => import('./locales/ru').then((m) => m.ru),
  ar: () => import('./locales/ar').then((m) => m.ar),
  pt: () => import('./locales/pt').then((m) => m.pt),
  it: () => import('./locales/it').then((m) => m.it),
  pl: () => import('./locales/pl').then((m) => m.pl),
  cs: () => import('./locales/cs').then((m) => m.cs),
  nl: () => import('./locales/nl').then((m) => m.nl),
  ms: () => import('./locales/ms').then((m) => m.ms),
  he: () => import('./locales/he').then((m) => m.he),
  hi: () => import('./locales/hi').then((m) => m.hi),
  'zh-TW': () => import('./locales/zh-TW').then((m) => m.zhTW),
  vi: () => import('./locales/vi').then((m) => m.vi),
}

const loaded = new Set<Lang>(['zh', 'en'])
const pending = new Map<Lang, Promise<void>>()

/**
 * Fetch a locale's dictionary chunk into the registry. Memoized per language and
 * never rejects: on failure the translator stays on its en/zh fallback.
 */
export function loadLocale(lang: Lang): Promise<void> {
  if (loaded.has(lang)) return Promise.resolve()
  const inFlight = pending.get(lang)
  if (inFlight) return inFlight
  const load = async (): Promise<void> => {
    const loader = lang === 'zh' || lang === 'en' ? undefined : LOADERS[lang]
    if (!loader) return
    dicts[lang] = await loader()
    loaded.add(lang)
  }
  const attempt = load().catch((err) => {
    pending.delete(lang) // freed so a later call can retry
    console.warn(`[i18n] locale ${lang} chunk failed to load; using en/zh fallback`, err)
  })
  pending.set(lang, attempt)
  return attempt
}
