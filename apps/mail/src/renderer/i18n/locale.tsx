import { createContext, useContext, useEffect, useState } from 'react'
import type { ReactNode } from 'react'
import { createI18n, htmlLang, type Lang, type Params } from '@genoffice/i18n'
import { strings } from './strings'

const translate = createI18n(strings)

export type StringKey = keyof typeof strings.zh
export type TFunc = (key: StringKey, params?: Params) => string

/** Module-level current language: for code outliving render closures, kept in sync with the Provider */
let moduleLang: Lang = 'zh'

export function t(key: StringKey, params?: Params): string {
  return translate(moduleLang, key, params)
}

const LocaleContext = createContext<Lang>('zh')

export function LocaleProvider({ initial, children }: { initial: Lang; children: ReactNode }) {
  const [lang, setLang] = useState<Lang>(initial)
  moduleLang = lang
  useEffect(
    () =>
      window.mailApi.onLanguageChanged((next) => {
        document.documentElement.lang = htmlLang(next)
        setLang(next)
      }),
    [],
  )
  return <LocaleContext.Provider value={lang}>{children}</LocaleContext.Provider>
}

export function useI18n(): { lang: Lang; t: TFunc } {
  const lang = useContext(LocaleContext)
  // t is not referentially stable by design; never put it in a hook dep array.
  return { lang, t: (key, params) => translate(lang, key, params) }
}
