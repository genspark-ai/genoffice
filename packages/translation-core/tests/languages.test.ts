import { describe, expect, it } from 'vitest'

import { LANGUAGES, englishLabelFor, getLanguage } from '../src/languages'

describe('getLanguage', () => {
  it('returns the canonical entry for known codes', () => {
    expect(getLanguage('zh-CN')?.englishLabel).toBe('Simplified Chinese')
    expect(getLanguage('en-US')?.label).toBe('English')
    expect(getLanguage('ja-JP')?.englishLabel).toBe('Japanese')
  })

  it('falls back to the family when the region is unknown', () => {
    expect(getLanguage('en-GB')?.value).toBe('en-US')
    expect(getLanguage('pt-BR')?.value).toBe('pt-PT')
  })

  it('returns null for empty / nullish / unknown input', () => {
    expect(getLanguage(undefined)).toBeNull()
    expect(getLanguage(null)).toBeNull()
    expect(getLanguage('')).toBeNull()
    expect(getLanguage('xx-YY')).toBeNull()
  })
})

describe('englishLabelFor', () => {
  it('uses the English label for prompts', () => {
    expect(englishLabelFor('zh-CN')).toBe('Simplified Chinese')
    expect(englishLabelFor('auto')).toBe('Auto-detect')
  })

  it('falls back to the raw code when the language is unknown', () => {
    expect(englishLabelFor('xx-YY')).toBe('xx-YY')
  })

  it('handles nullish input', () => {
    expect(englishLabelFor(undefined)).toBe('Auto-detect')
    expect(englishLabelFor(null)).toBe('Auto-detect')
  })
})

describe('LANGUAGES', () => {
  it('has a unique value for every entry', () => {
    const seen = new Set<string>()
    for (const l of LANGUAGES) {
      expect(seen.has(l.value)).toBe(false)
      seen.add(l.value)
    }
  })

  it('contains auto + a balanced set of locales', () => {
    expect(LANGUAGES.find((l) => l.value === 'auto')).toBeTruthy()
    expect(LANGUAGES.length).toBeGreaterThanOrEqual(10)
  })
})
