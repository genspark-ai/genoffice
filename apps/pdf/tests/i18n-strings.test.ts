import { describe, expect, it } from 'vitest'
import { LANGS } from '@genoffice/i18n'
import { strings } from '../src/renderer/i18n/strings'

const dicts = strings as Record<string, Record<string, string>>
const zhKeys = Object.keys(dicts.zh!).sort()

describe('i18n string tables', () => {
  it('provides a dictionary for every supported language and nothing else', () => {
    expect(Object.keys(dicts).sort()).toEqual([...LANGS].sort())
  })

  it.each([...LANGS])('locale %s has exactly the zh key set', (lang) => {
    expect(Object.keys(dicts[lang]!).sort()).toEqual(zhKeys)
  })

  it.each([...LANGS])('locale %s has no empty values', (lang) => {
    for (const [key, value] of Object.entries(dicts[lang]!)) {
      expect(typeof value, `${lang}.${key}`).toBe('string')
      expect(value.trim().length, `${lang}.${key} is empty`).toBeGreaterThan(0)
    }
  })

  it.each([...LANGS])('locale %s keeps the same placeholders as zh', (lang) => {
    for (const key of zhKeys) {
      const placeholders = (s: string) => (s.match(/\{\w+\}/g) ?? []).sort()
      expect(placeholders(dicts[lang]![key]!), `${lang}.${key}`).toEqual(
        placeholders(dicts.zh![key]!),
      )
    }
  })
})
