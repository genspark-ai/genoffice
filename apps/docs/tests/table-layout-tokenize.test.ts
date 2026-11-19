import { describe, expect, it } from 'vitest'
import { maxWordWidthPx, middleDotGlue } from '../src/renderer/table-layout/tokenize'

describe('maxWordWidthPx', () => {
  const w = (text: string) => maxWordWidthPx([{ text }])

  it('breaks after "-" like Word, never at "/" (Word probe: "mmmm/mmmm" one word)', () => {
    expect(w('aa-bbbb')).toBe(w('bbbb'))
    expect(w('aa/bbbb')).toBe(w('aa') + w('/') + w('bbbb'))
    // the token may span runs; the break opportunity carries across the boundary
    expect(maxWordWidthPx([{ text: 'aa-' }, { text: 'bbbb' }])).toBe(w('bbbb'))
  })

  it('breaks after "-" before a digit like Word, keeps fractions whole', () => {
    // Word probe: "10-|20", "DDR5-|6400", "UU-|2M4-|20-|5V"; a part number
    // whose every hyphen precedes a digit is not one unbreakable word
    expect(w('10-12')).toBe(w('10-'))
    expect(w('SPLC-6ASFC-26-3DH')).toBe(w('6ASFC-'))
    // 5 chars x 0.52em x 16px, unbroken
    expect(w('24/75')).toBeCloseTo(5 * 0.52 * 16)
  })

  it('breaks after an en/em dash like Word, never at U+2212 (Word probe 2026-11-08)', () => {
    // "mmmm\u2013mmmm" and "mmmm\u2014mmmm" wrapped after the dash in a 600-twip
    // column; "mmmm\u2212mmmm" forced the column to the whole string
    expect(w('aa\u2013bbbb')).toBe(w('bbbb'))
    expect(w('aa\u2014bbbb')).toBe(w('bbbb'))
    expect(w('3.9/R7\u2013R9 \u00b7 QE-22/23')).toBe(w('R9 \u00b7 QE-'))
    expect(w('aa\u2212bbbb')).toBe(w('aa') + w('\u2212') + w('bbbb'))
  })

  it('keeps a space beside a middle dot inside the word like Word', () => {
    // Word probe: a tblW-auto column grew to the whole "a \u00b7 b \u00b7 c" chain
    // (181.9 pt measured, 181.7 drawn) and never broke at its spaces
    expect(w('aa \u00b7 bb \u00b7 cc')).toBeGreaterThan(w('aa \u00b7 bb'))
    expect(w('aa \u00b7 bb')).toBeGreaterThan(w('aa') + w('bb'))
    expect(w('aa \u00b7 bb cc')).toBe(w('aa \u00b7 bb'))
    // the chain spans runs
    expect(maxWordWidthPx([{ text: 'aa \u00b7' }, { text: ' bb' }])).toBe(w('aa \u00b7 bb'))
    // a space before a colon stays inside the word, the one after it breaks
    expect(w('aa : bb')).toBe(w('aa :'))
    expect(w('aa : bb')).toBeGreaterThan(w('aa'))
  })

  it('the middle-dot glue follows the run East Asian language (Word probe 2026-11-08)', () => {
    const lang = (text: string, eastAsiaLang: string) => maxWordWidthPx([{ text, eastAsiaLang }])
    // zh-CN = Word's default: neither side breaks
    expect(lang('aa \u00b7 bb', 'zh-CN')).toBe(w('aa \u00b7 bb'))
    // zh-TW keeps the dot off a line start only: "aa \u00b7" | "bb"
    expect(lang('aa \u00b7 bbbb', 'zh-TW')).toBe(w('bbbb'))
    expect(lang('aaaa \u00b7 bb', 'zh-TW')).toBe(w('aaaa \u00b7'))
    expect(lang('aaaa : bb', 'zh-TW')).toBe(w('aaaa :'))
    // any other language breaks at the spaces: the dot and the colon stand alone
    for (const tag of ['pt-BR', 'en-US', 'ja-JP', 'ko-KR']) {
      expect(lang('aa \u00b7 bbbb', tag)).toBe(w('bbbb'))
      expect(lang('aaaa \u00b7 bb', tag)).toBe(w('aaaa'))
      expect(lang('aaaa : bb', tag)).toBe(w('aaaa'))
    }
    expect(middleDotGlue(undefined)).toBe('both')
    expect(middleDotGlue('zh-Hant-TW')).toBe('before')
  })
})
