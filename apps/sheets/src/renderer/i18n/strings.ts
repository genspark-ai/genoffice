import type { Lang, LangDicts } from '@genoffice/i18n'

/** One language's merged dictionary: each domain's zh shard defines its key set. */
type AppStrings = (typeof import('./app/zh'))['zh']
type DialogStrings = (typeof import('./dialogs/zh'))['zh']
type AiStrings = (typeof import('./ai/zh'))['zh']
export type Strings = AppStrings & DialogStrings & AiStrings

/**
 * Resident dictionaries for the synchronous translators. The entry bundle
 * ships none of them; `loadStrings` fills a slot in when that language's
 * chunks arrive. Every slot exists so lookups stay fully typed.
 */
export const dicts: LangDicts<Strings> = {
  zh: {} as Strings,
  en: {} as Strings,
  ja: {} as Strings,
  ko: {} as Strings,
  fr: {} as Strings,
  de: {} as Strings,
  es: {} as Strings,
  th: {} as Strings,
  id: {} as Strings,
  ru: {} as Strings,
  ar: {} as Strings,
  pt: {} as Strings,
  it: {} as Strings,
  pl: {} as Strings,
  cs: {} as Strings,
  nl: {} as Strings,
  ms: {} as Strings,
  he: {} as Strings,
  hi: {} as Strings,
  vi: {} as Strings,
  'zh-TW': {} as Strings,
}

/** Per-language chunk loaders; the spread order keeps ai winning key conflicts, as before. */
const LOADERS: Record<Lang, () => Promise<Strings>> = {
  zh: async () => ({
    ...(await import('./app/zh')).zh,
    ...(await import('./dialogs/zh')).zh,
    ...(await import('./ai/zh')).zh,
  }),
  en: async () => ({
    ...(await import('./app/en')).en,
    ...(await import('./dialogs/en')).en,
    ...(await import('./ai/en')).en,
  }),
  ja: async () => ({
    ...(await import('./app/ja')).ja,
    ...(await import('./dialogs/ja')).ja,
    ...(await import('./ai/ja')).ja,
  }),
  ko: async () => ({
    ...(await import('./app/ko')).ko,
    ...(await import('./dialogs/ko')).ko,
    ...(await import('./ai/ko')).ko,
  }),
  fr: async () => ({
    ...(await import('./app/fr')).fr,
    ...(await import('./dialogs/fr')).fr,
    ...(await import('./ai/fr')).fr,
  }),
  de: async () => ({
    ...(await import('./app/de')).de,
    ...(await import('./dialogs/de')).de,
    ...(await import('./ai/de')).de,
  }),
  es: async () => ({
    ...(await import('./app/es')).es,
    ...(await import('./dialogs/es')).es,
    ...(await import('./ai/es')).es,
  }),
  th: async () => ({
    ...(await import('./app/th')).th,
    ...(await import('./dialogs/th')).th,
    ...(await import('./ai/th')).th,
  }),
  id: async () => ({
    ...(await import('./app/id')).id,
    ...(await import('./dialogs/id')).id,
    ...(await import('./ai/id')).id,
  }),
  ru: async () => ({
    ...(await import('./app/ru')).ru,
    ...(await import('./dialogs/ru')).ru,
    ...(await import('./ai/ru')).ru,
  }),
  ar: async () => ({
    ...(await import('./app/ar')).ar,
    ...(await import('./dialogs/ar')).ar,
    ...(await import('./ai/ar')).ar,
  }),
  pt: async () => ({
    ...(await import('./app/pt')).pt,
    ...(await import('./dialogs/pt')).pt,
    ...(await import('./ai/pt')).pt,
  }),
  it: async () => ({
    ...(await import('./app/it')).it,
    ...(await import('./dialogs/it')).it,
    ...(await import('./ai/it')).it,
  }),
  pl: async () => ({
    ...(await import('./app/pl')).pl,
    ...(await import('./dialogs/pl')).pl,
    ...(await import('./ai/pl')).pl,
  }),
  cs: async () => ({
    ...(await import('./app/cs')).cs,
    ...(await import('./dialogs/cs')).cs,
    ...(await import('./ai/cs')).cs,
  }),
  nl: async () => ({
    ...(await import('./app/nl')).nl,
    ...(await import('./dialogs/nl')).nl,
    ...(await import('./ai/nl')).nl,
  }),
  ms: async () => ({
    ...(await import('./app/ms')).ms,
    ...(await import('./dialogs/ms')).ms,
    ...(await import('./ai/ms')).ms,
  }),
  he: async () => ({
    ...(await import('./app/he')).he,
    ...(await import('./dialogs/he')).he,
    ...(await import('./ai/he')).he,
  }),
  hi: async () => ({
    ...(await import('./app/hi')).hi,
    ...(await import('./dialogs/hi')).hi,
    ...(await import('./ai/hi')).hi,
  }),
  vi: async () => ({
    ...(await import('./app/vi')).vi,
    ...(await import('./dialogs/vi')).vi,
    ...(await import('./ai/vi')).vi,
  }),
  'zh-TW': async () => ({
    ...(await import('./app/zh-TW')).zhTW,
    ...(await import('./dialogs/zh-TW')).zhTW,
    ...(await import('./ai/zh-TW')).zhTW,
  }),
}

const loaded = new Map<Lang, Strings>()
const pending = new Map<Lang, Promise<Strings>>()

/**
 * Fetch one language's dictionary and register it for the synchronous
 * translators; concurrent calls share one load. A failed load serves the
 * English dictionary, which bootstrap keeps resident as the fallback.
 */
export function loadStrings(lang: Lang): Promise<Strings> {
  const cached = loaded.get(lang)
  if (cached) return Promise.resolve(cached)
  const inFlight = pending.get(lang)
  if (inFlight) return inFlight
  const load = LOADERS[lang]().then(
    (dict) => {
      loaded.set(lang, dict)
      Object.assign(dicts[lang], dict)
      return dict
    },
    () =>
      lang === 'en'
        ? Promise.reject<Strings>(new Error('en strings chunk failed to load'))
        : loadStrings('en').then((fallback) => {
            loaded.set(lang, fallback)
            Object.assign(dicts[lang], fallback)
            return fallback
          }),
  )
  pending.set(lang, load)
  void load.catch(() => pending.delete(lang))
  return load
}
