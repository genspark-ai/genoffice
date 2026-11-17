import { zh as l0, type LocalVlmKey } from './local-vlm/zh'
import { copy as l1 } from './local-vlm/zh-TW'
import { copy as l2 } from './local-vlm/en'
import { copy as l3 } from './local-vlm/vi'
import { copy as l4 } from './local-vlm/ar'
import { copy as l5 } from './local-vlm/cs'
import { copy as l6 } from './local-vlm/de'
import { copy as l7 } from './local-vlm/es'
import { copy as l8 } from './local-vlm/fr'
import { copy as l9 } from './local-vlm/he'
import { copy as l10 } from './local-vlm/hi'
import { copy as l11 } from './local-vlm/id'
import { copy as l12 } from './local-vlm/it'
import { copy as l13 } from './local-vlm/ja'
import { copy as l14 } from './local-vlm/ko'
import { copy as l15 } from './local-vlm/ms'
import { copy as l16 } from './local-vlm/nl'
import { copy as l17 } from './local-vlm/pl'
import { copy as l18 } from './local-vlm/pt'
import { copy as l19 } from './local-vlm/ru'
import { copy as l20 } from './local-vlm/th'

const copies: Record<string, Record<LocalVlmKey, string>> = {
  zh: l0,
  'zh-TW': l1,
  en: l2,
  vi: l3,
  ar: l4,
  cs: l5,
  de: l6,
  es: l7,
  fr: l8,
  he: l9,
  hi: l10,
  id: l11,
  it: l12,
  ja: l13,
  ko: l14,
  ms: l15,
  nl: l16,
  pl: l17,
  pt: l18,
  ru: l19,
  th: l20,
}

export function localVlmString(
  lang: string,
  key: LocalVlmKey,
  params?: Record<string, string | number>,
): string {
  const dict = copies[lang] ?? l2
  let text = dict[key] ?? l2[key]
  if (params) {
    for (const [k, v] of Object.entries(params)) {
      text = text.replaceAll(`{${k}}`, String(v))
    }
  }
  return text
}
