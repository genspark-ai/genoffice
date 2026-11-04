import { describe, expect, it } from 'vitest'
import {
  charClass,
  isCjkLetterChar,
  isDoubleByteChar,
  isEastAsianChar,
  isEastAsianLangTag,
  isHangulChar,
  textHasEastAsian,
} from '../src/char-class'

const cp = (s: string) => s.codePointAt(0)!

describe('charClass', () => {
  it('classifies the scripts, punctuation and width forms', () => {
    expect(charClass(cp('\u6f22'))).toBe('han')
    expect(charClass(cp('\u3005'))).toBe('han')
    expect(charClass(0x20000)).toBe('han')
    expect(charClass(cp('\u3042'))).toBe('kana')
    expect(charClass(cp('\u30fc'))).toBe('kana')
    expect(charClass(cp('\u30fb'))).toBe('cjkPunct')
    expect(charClass(cp('\u3002'))).toBe('cjkPunct')
    expect(charClass(cp('\u3000'))).toBe('ideographicSpace')
    expect(charClass(cp('\ud55c'))).toBe('hangul')
    expect(charClass(0x1100)).toBe('hangul')
    expect(charClass(cp('\u3131'))).toBe('hangul')
    expect(charClass(cp('\uff08'))).toBe('fullwidth')
    expect(charClass(cp('\uff61'))).toBe('halfwidthEA')
    expect(charClass(cp('\uff76'))).toBe('halfwidthEA')
    expect(charClass(cp('\u33a1'))).toBe('cjkSymbol')
    expect(charClass(cp('A'))).toBe('latinLetter')
    expect(charClass(cp('\u00e9'))).toBe('latinLetter')
    expect(charClass(cp('7'))).toBe('digit')
    expect(charClass(cp(','))).toBe('asciiPunct')
    expect(charClass(0x20)).toBe('space')
    expect(charClass(0xa0)).toBe('nbsp')
    expect(charClass(cp('\u0631'))).toBe('other')
  })

  it('derives the rule sets', () => {
    expect(isEastAsianChar(cp('\uff61'))).toBe(true)
    expect(isDoubleByteChar(cp('\uff61'))).toBe(false)
    expect(isDoubleByteChar(cp('\u6f22'))).toBe(true)
    expect(isDoubleByteChar(0x1100)).toBe(true)
    expect(isCjkLetterChar(cp('\u3002'))).toBe(false)
    expect(isCjkLetterChar(cp('\u30fb'))).toBe(false)
    expect(isCjkLetterChar(cp('\uff08'))).toBe(false)
    expect(isCjkLetterChar(cp('\ud55c'))).toBe(true)
    expect(isHangulChar(cp('\u6f22'))).toBe(false)
    expect(textHasEastAsian('abc 123')).toBe(false)
    expect(textHasEastAsian('abc \u30fc')).toBe(true)
  })

  it('reads East Asian language tags, Korean optional', () => {
    expect(isEastAsianLangTag('zh-CN')).toBe(true)
    expect(isEastAsianLangTag('ja')).toBe(true)
    expect(isEastAsianLangTag('ko-KR')).toBe(true)
    expect(isEastAsianLangTag('ko-KR', false)).toBe(false)
    expect(isEastAsianLangTag('en-US')).toBe(false)
    expect(isEastAsianLangTag(undefined)).toBe(false)
    expect(isEastAsianLangTag('zho')).toBe(false)
  })
})
