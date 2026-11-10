import { describe, expect, it } from 'vitest'
import { strings } from '../src/renderer/src/strings'

/**
 * Home-screen locale tables (src/renderer/src/strings.ts): zh defines the key
 * set; every other locale must cover exactly the same keys with real content.
 */

const locales = Object.keys(strings) as Array<keyof typeof strings>
const referenceKeys = Object.keys(strings.zh).sort()

/** placeholders like {n}, {name}, {v} embedded in a template */
function placeholdersOf(template: string): string[] {
  return (template.match(/\{[a-zA-Z0-9]+\}/g) ?? []).sort()
}

describe('home-screen locale tables', () => {
  it('includes the expected UI languages', () => {
    expect(locales).toContain('zh')
    expect(locales).toContain('en')
    expect(locales).toContain('zh-TW')
    expect(locales.length).toBeGreaterThanOrEqual(19)
  })

  it.each(locales)('locale %s has exactly the zh key set', (locale) => {
    expect(Object.keys(strings[locale]).sort()).toEqual(referenceKeys)
  })

  it.each(locales)('locale %s has no empty or whitespace-only values', (locale) => {
    const empty = Object.entries(strings[locale])
      .filter(([, value]) => typeof value !== 'string' || value.trim().length === 0)
      .map(([key]) => key)
    expect(empty).toEqual([])
  })

  it.each(locales)('locale %s keeps the zh placeholder set for each key', (locale) => {
    const table = strings[locale] as Record<string, string>
    const mismatched = referenceKeys.filter((key) => {
      const localePlaceholders = placeholdersOf(table[key])
      const zhPlaceholders = placeholdersOf((strings.zh as Record<string, string>)[key])
      // Singular-count keys ("...One") may drop the numeral entirely in
      // languages that express "one" grammatically (e.g. ar/he), so a
      // missing placeholder is fine there — an extra one is not.
      if (key.endsWith('One')) {
        return localePlaceholders.some((p) => !zhPlaceholders.includes(p))
      }
      return localePlaceholders.join(',') !== zhPlaceholders.join(',')
    })
    expect(mismatched).toEqual([])
  })

  it('has no duplicate values that suggest an untranslated copy-paste between zh and en', () => {
    // sanity check that en is actually translated, not a zh copy
    const zh = strings.zh as Record<string, string>
    const en = strings.en as Record<string, string>
    const identical = referenceKeys.filter((key) => zh[key] === en[key])
    // a few shared strings (brand names, "PDF", "OK"-style tokens) are fine
    expect(identical.length).toBeLessThan(referenceKeys.length / 4)
  })

  /**
   * The app is still called AI Markdown. A rename of the app's name was started
   * in this branch and then dropped: the naming is the maintainers' call, not
   * this PR's, and README and the docs still say Markdown. Pinned so a stray
   * locale cannot quietly reintroduce a second name — the quick-create card,
   * the File > New menu and a new tab's title all have to agree.
   */
  it.each(locales)('locale %s calls the app AI Markdown on the quick-create card', (locale) => {
    expect(strings[locale].newMarkdown).toBe('AI Markdown')
  })
})
