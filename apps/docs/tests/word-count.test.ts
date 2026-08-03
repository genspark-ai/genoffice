import { describe, expect, it } from 'vitest'
import { asianCharCount, countWords, nonAsianWordCount } from '../src/renderer/word-count'

describe('word count CJK rule', () => {
  it('counts asian chars one by one, punctuation included', () => {
    expect(asianCharCount('深圳市人民政府')).toBe(7)
    expect(asianCharCount('残疾预防，从我做起。')).toBe(10)
    expect(asianCharCount('（市残联、市卫生健康委）')).toBe(12)
    expect(asianCharCount('ひらがなカタカナ')).toBe(8)
    expect(asianCharCount('한국어')).toBe(3)
  })

  it('counts non-asian words by word, not by character', () => {
    expect(nonAsianWordCount('hello world')).toBe(2)
    expect(nonAsianWordCount("it's a state-of-the-art plan")).toBe(4)
    expect(nonAsianWordCount('2026年8月2日')).toBe(3)
    expect(nonAsianWordCount('深圳市')).toBe(0)
  })

  it('Words = asian chars + non-asian words (Word parity)', () => {
    const text = '按照 WHO 的要求，2025 年底前完成。'
    // 12 asian chars (incl. the fullwidth comma and period); non-asian words "WHO" + "2025" = 2
    expect(asianCharCount(text)).toBe(12)
    expect(nonAsianWordCount(text)).toBe(2)
    expect(countWords(text)).toBe(14)
  })

  it('half-width digits/latin embedded in CJK are words, full-width punctuation is a char', () => {
    expect(countWords('第1章：GenOffice 使用指南')).toBe(9)
    // 7 asian chars (incl. the fullwidth colon) + "1" and "GenOffice" as 2 words
  })
})
