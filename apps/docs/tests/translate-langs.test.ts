import { describe, expect, it } from 'vitest'
import {
  TRANSLATE_LANGS,
  appLangKey,
  ribbonLangKey,
  type TranslateLang,
} from '../src/renderer/components/translate-langs'

/** The seven languages the app shipped with, in their original order. */
const SHIPPED: Array<[TranslateLang, string]> = [
  ['en', 'English'],
  ['zh', 'SimplifiedChinese'],
  ['ja', 'Japanese'],
  ['ko', 'Korean'],
  ['fr', 'French'],
  ['de', 'German'],
  ['es', 'Spanish'],
]

describe('translate language keys', () => {
  it('reproduces the seven keys that already ship, in order', () => {
    // These strings appear in 21 locale files. A refactor that changed one
    // would leave the menu showing a raw key, or nothing at all.
    expect(TRANSLATE_LANGS.slice(0, 7).map((l) => appLangKey(l.code))).toEqual([
      'appLangEnglish',
      'appLangSimplifiedChinese',
      'appLangJapanese',
      'appLangKorean',
      'appLangFrench',
      'appLangGerman',
      'appLangSpanish',
    ])
    expect(TRANSLATE_LANGS.slice(0, 7).map((l) => ribbonLangKey(l.code))).toEqual(
      SHIPPED.map(([, k]) => `ribbonLang${k}`),
    )
  })

  it('covers every language the UI itself ships', () => {
    // 21 locales; the translate menu must be able to offer every one of them
    expect(TRANSLATE_LANGS).toHaveLength(21)
  })

  it('names zh-TW explicitly, which no join-then-capitalise rule could express', () => {
    expect(appLangKey('zh-TW')).toBe('appLangTraditionalChinese')
    expect(ribbonLangKey('zh-TW')).toBe('ribbonLangTraditionalChinese')
    expect(appLangKey('zh')).toBe('appLangSimplifiedChinese')
  })

  it('gives every language a distinct key on both surfaces', () => {
    const app = TRANSLATE_LANGS.map((l) => appLangKey(l.code))
    const ribbon = TRANSLATE_LANGS.map((l) => ribbonLangKey(l.code))
    expect(new Set(app).size).toBe(TRANSLATE_LANGS.length)
    expect(new Set(ribbon).size).toBe(TRANSLATE_LANGS.length)
  })

  it('spells keys out in full rather than abbreviating the code', () => {
    // a "smart" rule would produce appLangEn / appLangZh-tw and break every
    // one of these against the strings that already exist
    expect(appLangKey('en')).not.toBe('appLangEn')
    expect(appLangKey('zh-TW')).not.toContain('-')
    for (const { key } of TRANSLATE_LANGS) expect(key.length).toBeGreaterThan(2)
  })

  it('keeps both surfaces in step — the same language maps to parallel keys', () => {
    for (const { code } of TRANSLATE_LANGS) {
      expect(appLangKey(code).replace('appLang', '')).toBe(
        ribbonLangKey(code).replace('ribbonLang', ''),
      )
    }
  })
})
