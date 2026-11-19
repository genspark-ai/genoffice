import { describe, expect, it } from 'vitest'
import { bidiLangOf, eastAsiaLangOf } from '../src/generate/document-meta'

/** one IR node with a run, the shape eastAsiaLangOf's script counter walks */
const ir = (text: string) => [{ runs: [{ text }] }]

describe('eastAsiaLangOf', () => {
  it('falls back to script counting when docsettings carries lang: null', () => {
    expect(eastAsiaLangOf(null, ir('にほんごのにゅうしょをかく'))).toBe('ja-JP')
    expect(eastAsiaLangOf(null, ir('안녕하세요 여러분 오늘 날씨가 좋습니다'))).toBe('ko-KR')
    expect(eastAsiaLangOf(null, [])).toBe('zh-CN')
  })

  it('honors a real lang attribute', () => {
    expect(eastAsiaLangOf('zh-Hant', [])).toBe('zh-TW')
    expect(eastAsiaLangOf('en-US', [])).toBe('zh-CN')
  })
})

describe('bidiLangOf', () => {
  it('falls back to the rtl default when docsettings carries lang: null', () => {
    expect(bidiLangOf(null)).toBe('ar-SA')
    expect(bidiLangOf('he-IL')).toBe('he-IL')
    expect(bidiLangOf('ur')).toBe('ur-PK')
  })
})
