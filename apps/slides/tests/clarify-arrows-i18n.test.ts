import { describe, expect, it } from 'vitest'
import { zh } from '../src/renderer/i18n/ai/zh'
import { en } from '../src/renderer/i18n/ai/en'
import { ja } from '../src/renderer/i18n/ai/ja'
import { ko } from '../src/renderer/i18n/ai/ko'
import { fr } from '../src/renderer/i18n/ai/fr'
import { de } from '../src/renderer/i18n/ai/de'
import { es } from '../src/renderer/i18n/ai/es'
import { th } from '../src/renderer/i18n/ai/th'
import { id } from '../src/renderer/i18n/ai/id'
import { ru } from '../src/renderer/i18n/ai/ru'
import { ar } from '../src/renderer/i18n/ai/ar'
import { pt } from '../src/renderer/i18n/ai/pt'
import { it as itStrings } from '../src/renderer/i18n/ai/it'
import { pl } from '../src/renderer/i18n/ai/pl'
import { cs } from '../src/renderer/i18n/ai/cs'
import { nl } from '../src/renderer/i18n/ai/nl'
import { ms } from '../src/renderer/i18n/ai/ms'
import { he } from '../src/renderer/i18n/ai/he'
import { hi } from '../src/renderer/i18n/ai/hi'
import { zhTW } from '../src/renderer/i18n/ai/zh-TW'
import { vi } from '../src/renderer/i18n/ai/vi'

const shards = {
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
  it: itStrings,
  pl,
  cs,
  nl,
  ms,
  he,
  hi,
  zhTW,
  vi,
}

describe('clarify arrow labels', () => {
  it('every locale labels the prev/next question arrows', () => {
    for (const [locale, strings] of Object.entries(shards)) {
      expect(strings.aiClarifyPrev, `${locale}.aiClarifyPrev`).toBeTruthy()
      expect(strings.aiClarifyNext, `${locale}.aiClarifyNext`).toBeTruthy()
    }
  })
})
