import { describe, expect, it } from 'vitest'
import {
  isCombiningMark,
  isEastAsianScript,
  isNoSpaceScript,
  isRtlScript,
  scriptOf,
} from '../src/script'

const cp = (s: string) => s.codePointAt(0)!

describe('scriptOf', () => {
  it('classifies the P1 scripts', () => {
    expect(scriptOf(cp('a'))).toBe('latin')
    expect(scriptOf(cp('Ω') - 0)).toBe('common') // Greek is out of P1 scope → common
    expect(scriptOf(cp('é'))).toBe('latin')
    expect(scriptOf(cp('中'))).toBe('cjk')
    expect(scriptOf(cp('𠀀'))).toBe('cjk') // ext B (astral)
    expect(scriptOf(cp('。'))).toBe('cjk') // fullwidth punctuation
    expect(scriptOf(0xff21)).toBe('cjk') // fullwidth latin Ａ
    expect(scriptOf(cp('あ'))).toBe('kana')
    expect(scriptOf(cp('ア'))).toBe('kana')
    expect(scriptOf(cp('ｱ'))).toBe('kana') // halfwidth katakana
    expect(scriptOf(cp('한'))).toBe('hangul')
    expect(scriptOf(cp('ㄱ'))).toBe('hangul')
    expect(scriptOf(cp('ก'))).toBe('thai')
    expect(scriptOf(cp('م'))).toBe('arabic')
    expect(scriptOf(cp('ﻣ'))).toBe('arabic') // presentation form
    expect(scriptOf(cp('א'))).toBe('hebrew')
    expect(scriptOf(cp('1'))).toBe('common')
    expect(scriptOf(cp('.'))).toBe('common')
    expect(scriptOf(cp(' '))).toBe('common')
  })

  it('radical / compatibility code points are cjk (P7: no machine spaces, eastAsia slot)', () => {
    expect(scriptOf(0x2f00)).toBe('cjk') // Kangxi radical ⼀
    expect(scriptOf(0x2f2f)).toBe('cjk') // Kangxi radical ⼯
    expect(scriptOf(0x2fdf)).toBe('cjk') // Kangxi block end
    expect(scriptOf(0x2e80)).toBe('cjk') // CJK Radicals Supplement start
    expect(scriptOf(0x2eff)).toBe('cjk')
    expect(scriptOf(0xf900)).toBe('cjk') // CJK Compatibility Ideographs
    expect(scriptOf(0xfaff)).toBe('cjk')
    expect(scriptOf(0xfe30)).toBe('cjk') // CJK Compatibility Forms
    expect(scriptOf(0xfe4f)).toBe('cjk')
  })
})

describe('script predicates', () => {
  it('no-space scripts are cjk/kana/thai — NOT hangul', () => {
    expect(isNoSpaceScript('cjk')).toBe(true)
    expect(isNoSpaceScript('kana')).toBe(true)
    expect(isNoSpaceScript('thai')).toBe(true)
    expect(isNoSpaceScript('hangul')).toBe(false)
    expect(isNoSpaceScript('latin')).toBe(false)
  })

  it('rtl scripts', () => {
    expect(isRtlScript('arabic')).toBe(true)
    expect(isRtlScript('hebrew')).toBe(true)
    expect(isRtlScript('latin')).toBe(false)
  })

  it('east-asian font-slot scripts include hangul', () => {
    expect(isEastAsianScript('cjk')).toBe(true)
    expect(isEastAsianScript('kana')).toBe(true)
    expect(isEastAsianScript('hangul')).toBe(true)
    expect(isEastAsianScript('thai')).toBe(false)
    expect(isEastAsianScript('latin')).toBe(false)
  })

  it('combining marks: latin diacritics, thai vowels/tones, zero-width chars', () => {
    expect(isCombiningMark(0x0301)).toBe(true) // combining acute
    expect(isCombiningMark(0x0e34)).toBe(true) // thai sara i
    expect(isCombiningMark(0x0e48)).toBe(true) // thai mai ek
    expect(isCombiningMark(0x200b)).toBe(true) // zero-width space
    expect(isCombiningMark(cp('a'))).toBe(false)
    expect(isCombiningMark(cp('中'))).toBe(false)
  })
})
