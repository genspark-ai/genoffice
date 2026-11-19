import { describe, expect, it } from 'vitest'
import {
  fontSizeLabel,
  fontSizeList,
  fontSizeOptions,
  namedFontSizes,
  parseFontSize,
  stepFontSize,
} from '../src/renderer/font-sizes'

describe('font size list per UI language', () => {
  it('lists Word en-US sizes for Latin UIs and adds 10.5 for Japanese', () => {
    expect(fontSizeList('en')).toEqual([
      8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72,
    ])
    expect(fontSizeList('fr')).toEqual(fontSizeList('en'))
    expect(fontSizeList('ja')).toContain(10.5)
    expect(namedFontSizes('en')).toEqual([])
    expect(namedFontSizes('ja')).toEqual([])
  })
  it('puts the CJK named scale ahead of the half-point list for zh and zh-TW', () => {
    expect(fontSizeList('zh')).toEqual([
      5, 5.5, 6.5, 7.5, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72,
    ])
    const named = namedFontSizes('zh')
    expect(named.map((n) => n.pt)).toEqual([
      42, 36, 26, 24, 22, 18, 16, 15, 14, 12, 10.5, 9, 7.5, 6.5, 5.5, 5,
    ])
    expect(named[0].name).toBe('\u521d\u53f7')
    expect(named[9].name).toBe('\u5c0f\u56db')
    expect(namedFontSizes('zh-TW')[0].name).toBe('\u521d\u865f')
    expect(namedFontSizes('zh-TW')[1].name).toBe('\u5c0f\u521d')
    expect(fontSizeOptions('zh').slice(0, 16)).toEqual(named)
    expect(fontSizeOptions('zh')[16]).toEqual({ name: '5', pt: 5 })
  })
  it('labels a size by its CJK name only in a CJK UI and parses both forms', () => {
    expect(fontSizeLabel(12, 'zh')).toBe('\u5c0f\u56db')
    expect(fontSizeLabel(12, 'zh-TW')).toBe('\u5c0f\u56db')
    expect(fontSizeLabel(12, 'en')).toBe('12')
    expect(fontSizeLabel(13, 'zh')).toBe('13')
    expect(parseFontSize('\u5c0f\u56db', 'zh')).toBe(12)
    expect(parseFontSize('\u4e94\u865f', 'zh-TW')).toBe(10.5)
    expect(parseFontSize('\u5c0f\u56db', 'en')).toBeNull()
    expect(parseFontSize(' 13.3 ', 'en')).toBe(13.5)
    expect(parseFontSize('12pt', 'en')).toBe(12)
    expect(parseFontSize('\uff11\uff12', 'zh')).toBe(12)
    expect(parseFontSize('0', 'en')).toBeNull()
    expect(parseFontSize('abc', 'en')).toBeNull()
    expect(parseFontSize('5000', 'en')).toBe(1638)
  })
})

describe('Increase / Decrease Font Size stepping', () => {
  it('walks the list inside it', () => {
    expect(stepFontSize(11, 1, 'en')).toBe(12)
    expect(stepFontSize(11.5, 1, 'en')).toBe(12)
    expect(stepFontSize(12, -1, 'en')).toBe(11)
    expect(stepFontSize(11.5, -1, 'en')).toBe(11)
    expect(stepFontSize(10, 1, 'ja')).toBe(10.5)
    expect(stepFontSize(10, 1, 'en')).toBe(11)
    expect(stepFontSize(28, 1, 'zh')).toBe(36)
  })
  it('goes by tens above the list up to 1638', () => {
    expect(stepFontSize(72, 1, 'en')).toBe(80)
    expect(stepFontSize(75, 1, 'en')).toBe(80)
    expect(stepFontSize(80, 1, 'en')).toBe(90)
    expect(stepFontSize(1630, 1, 'en')).toBe(1638)
    expect(stepFontSize(1638, 1, 'en')).toBe(1638)
    expect(stepFontSize(1638, -1, 'en')).toBe(1630)
    expect(stepFontSize(95, -1, 'en')).toBe(90)
    expect(stepFontSize(90, -1, 'en')).toBe(80)
    expect(stepFontSize(80, -1, 'en')).toBe(72)
    expect(stepFontSize(75, -1, 'zh')).toBe(72)
  })
  it('moves one point at a time below the list down to 1', () => {
    expect(stepFontSize(8, -1, 'en')).toBe(7)
    expect(stepFontSize(7.5, -1, 'en')).toBe(7)
    expect(stepFontSize(1, -1, 'en')).toBe(1)
    expect(stepFontSize(6, 1, 'en')).toBe(7)
    expect(stepFontSize(7.5, 1, 'en')).toBe(8)
    expect(stepFontSize(5, -1, 'zh')).toBe(4)
    expect(stepFontSize(4, 1, 'zh')).toBe(5)
  })
})
