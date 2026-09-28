import { describe, expect, it } from 'vitest'
import { bidiLangOf, eastAsiaLangOf } from '../src/generate/document-meta'

describe('docsettings lang helpers (#1381)', () => {
  it('treats a null or non-string lang as absent instead of throwing', () => {
    expect(eastAsiaLangOf(null)).toBe(eastAsiaLangOf(''))
    expect(eastAsiaLangOf(42 as unknown as string)).toBe(eastAsiaLangOf(''))
    expect(bidiLangOf(null)).toBe(bidiLangOf(''))
    expect(bidiLangOf(42 as unknown as string)).toBe(bidiLangOf(''))
  })

  it('still resolves real lang tags', () => {
    expect(eastAsiaLangOf('ja')).toBe('ja-JP')
    expect(eastAsiaLangOf('zh-TW')).toBe('zh-TW')
    expect(bidiLangOf('he-IL')).toBe('he-IL')
  })
})
