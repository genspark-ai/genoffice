import { describe, expect, it } from 'vitest'

import {
  DEFAULT_SHORT_DATE,
  getSystemShortDate,
  setSystemShortDate,
  shortDateNumFmtId,
  shortDatePatternForSystemLocale,
} from '../src/shared/short-date'

describe('shortDatePatternForSystemLocale', () => {
  it.each([
    ['en-US', 'm/d/yyyy'],
    ['zh-CN', 'yyyy/m/d'],
    // The region decides, not the language: an English UI on a CN-region
    // machine (macOS AppleLocale en_CN) must still produce yyyy/m/d.
    ['en-CN', 'yyyy/m/d'],
    ['en-GB', 'dd/mm/yyyy'],
    ['de-DE', 'd.m.yyyy'],
    ['ja-JP', 'yyyy/m/d'],
  ])('%s → %s', (tag, pattern) => {
    expect(shortDatePatternForSystemLocale(tag)).toBe(pattern)
  })

  it('keeps only date tokens and plain separators', () => {
    // ar-SA emits RTL marks as literals and defaults to the Islamic
    // calendar; both must be neutralized.
    const pattern = shortDatePatternForSystemLocale('ar-SA')
    expect(pattern).toMatch(/^[ymd./\- ,]+$/)
    expect(pattern).toContain('yyyy')
  })

  it('falls back on invalid tags', () => {
    expect(shortDatePatternForSystemLocale('')).toBe(DEFAULT_SHORT_DATE)
  })

  it('defaults the region to US when the tag has none', () => {
    expect(shortDatePatternForSystemLocale('en')).toBe('m/d/yyyy')
  })
})

describe('shortDateNumFmtId', () => {
  it('maps the system short date (and its datetime twin) back to 14/22', () => {
    const previous = getSystemShortDate()
    try {
      setSystemShortDate('yyyy/m/d')
      expect(shortDateNumFmtId('yyyy/m/d')).toBe(14)
      expect(shortDateNumFmtId('yyyy/m/d h:mm')).toBe(22)
      expect(shortDateNumFmtId('yyyy/mm/dd')).toBeUndefined()
    } finally {
      setSystemShortDate(previous)
    }
  })
})
