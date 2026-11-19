import { describe, expect, it } from 'vitest'
import { strings } from '../src/renderer/i18n/strings'

/**
 * zh defines the key set, every sibling locale must cover it. Each locale
 * block here spreads `fillFormStringsFor(lang)` on top of its own keys, so
 * the comparison runs on the effective key set the object exposes. The
 * LangDicts constraint at the `createI18n(strings)` call already turns a
 * missing key into a compile error, but a type error names neither the
 * locale nor the key — these runtime checks read the shipped table
 * directly, in the style of apps/docs/tests/translate-i18n-completeness
 * .test.ts, and fail with a message that says exactly which keys moved.
 */

const LOCALES = Object.keys(strings) as Array<keyof typeof strings>

describe('pdf strings — key parity', () => {
  it('every locale carries exactly the zh key set', () => {
    const reference = Object.keys(strings.zh).sort()
    expect(reference.length, 'zh defines the key set').toBeGreaterThan(0)
    for (const lang of LOCALES) {
      const keys = Object.keys(strings[lang]).sort()
      const missing = reference.filter((key) => !keys.includes(key))
      const extra = keys.filter((key) => !reference.includes(key))
      expect(
        missing,
        `${lang} is missing ${missing.length} key(s): ${missing.slice(0, 8).join(', ')}`,
      ).toEqual([])
      expect(
        extra,
        `${lang} carries ${extra.length} key(s) zh does not have: ${extra.slice(0, 8).join(', ')}`,
      ).toEqual([])
    }
  })

  it('no locale ships an empty label', () => {
    for (const lang of LOCALES) {
      for (const [key, value] of Object.entries(strings[lang])) {
        expect(value.length, `${lang}.${key} is an empty string`).toBeGreaterThan(0)
      }
    }
  })
})
