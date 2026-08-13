import { createContext, useContext, useMemo, useState, type ReactNode } from 'react'
import { createI18n, type Lang } from '@genoffice/i18n'
import { strings, type StringKeys } from './strings'

const tFunc = createI18n(strings)

export type TFunc = (key: StringKeys, params?: Record<string, string | number>) => string

interface I18n {
  lang: Lang
  t: TFunc
}

const LocaleContext = createContext<I18n>({ lang: 'en', t: (k) => strings.en[k] })

export function useI18n(): I18n {
  return useContext(LocaleContext)
}

export function t(key: StringKeys, params?: Record<string, string | number>): string {
  return tFunc(moduleLang, key, params)
}

let moduleLang: Lang = 'en'

export function LocaleProvider({ children }: { children: ReactNode }) {
  const [lang, setLang] = useState<Lang>(() => {
    try {
      return (navigator.language.split('-')[0] as Lang) || 'en'
    } catch {
      return 'en'
    }
  })

  moduleLang = lang

  useEffect(() => {
    const unsub = window.markdownApi?.onLanguageChanged?.((newLang: string) => {
      const l = newLang.split('-')[0] as Lang
      if (l) {
        setLang(l)
        moduleLang = l
      }
    })
    return unsub
  }, [])

  const value = useMemo(
    () => ({ lang, t: (k: StringKeys, p?: Record<string, string | number>) => tFunc(lang, k, p) }),
    [lang],
  )

  return <LocaleContext.Provider value={value}>{children}</LocaleContext.Provider>
}

import { useEffect } from 'react'
