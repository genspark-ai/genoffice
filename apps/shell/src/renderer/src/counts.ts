import type { Lang, Params } from '@genoffice/i18n'
import type { RecentPage } from '../../shared/home-api'
import type { StringKey } from './locale'

/** Sidebar counts use the same filtered total as the visible list. */
export function visiblePageCount(page: Pick<RecentPage, 'total'>): number {
  return page.total
}

type ExtraCategory = Exclude<Intl.LDMLPluralRule, 'one' | 'other'>

// forms beyond one/other, only for locales whose CLDR categories need
// distinct wording; the dictionaries keep the shared one/other pair
const FILE_COUNT_FORMS: Partial<Record<Lang, Partial<Record<ExtraCategory, string>>>> = {
  cs: { few: '{n} soubory' },
  ar: { zero: 'لا توجد ملفات', two: 'ملفان', few: '{n} ملفات', many: '{n} ملفًا' },
}

const rulesCache = new Map<Lang, Intl.PluralRules>()

export function fileCountLabel(
  count: number,
  lang: Lang,
  t: (key: StringKey, params?: Params) => string,
): string {
  let rules = rulesCache.get(lang)
  if (!rules) {
    rules = new Intl.PluralRules(lang)
    rulesCache.set(lang, rules)
  }
  const category = rules.select(count)
  const extra =
    category === 'one' || category === 'other' ? undefined : FILE_COUNT_FORMS[lang]?.[category]
  if (extra) return extra.replace('{n}', String(count))
  return t(category === 'one' ? 'fileCountOne' : 'fileCount', { n: count })
}
