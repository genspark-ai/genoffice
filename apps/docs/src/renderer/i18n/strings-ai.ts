import { defineStrings } from '@genoffice/i18n'
import { zh } from './ai/zh'
import { en } from './ai/en'
import { ja } from './ai/ja'
import { ko } from './ai/ko'
import { fr } from './ai/fr'
import { de } from './ai/de'
import { es } from './ai/es'
import { th } from './ai/th'
import { id } from './ai/id'
import { ru } from './ai/ru'
import { ar } from './ai/ar'
import { pt } from './ai/pt'
import { it } from './ai/it'
import { pl } from './ai/pl'
import { nl } from './ai/nl'
import { ms } from './ai/ms'
import { he } from './ai/he'
import { hi } from './ai/hi'
import { zhTW } from './ai/zh-TW'

/** User-visible strings for the ai/ panel and tool feedback (LLM prompts are not here) */
export const aiStrings = defineStrings({
  zh,
  en,
  ja,
  ko,
  fr,
  de,
  es,
  th,
  id,
  ru,
  ar,
  pt,
  it,
  pl,
  nl,
  ms,
  he,
  hi,
  'zh-TW': zhTW,
})
