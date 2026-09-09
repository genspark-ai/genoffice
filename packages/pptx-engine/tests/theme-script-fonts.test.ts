import { describe, expect, it } from 'vitest'
import { eaScriptOfLang, parseTheme, resolveFontRef } from '../src/theme'

const THEME = (ea: string) => `<?xml version="1.0"?><a:theme xmlns:a="a"><a:themeElements>
<a:clrScheme name="x"><a:dk1><a:srgbClr val="000000"/></a:dk1></a:clrScheme>
<a:fontScheme name="x">
  <a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="굴림"/></a:majorFont>
  <a:minorFont><a:latin typeface="Arial"/><a:ea typeface="${ea}"/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hans" typeface="\u9ed1\u4f53"/></a:minorFont>
</a:fontScheme><a:fmtScheme name="x"/></a:themeElements></a:theme>`

describe('theme per-script fonts', () => {
  it('parses <a:font script=…> entries per major/minor font', () => {
    const t = parseTheme(THEME(''))
    expect(t.minorEaFont).toBeUndefined()
    expect(t.minorScriptFonts).toEqual({ Jpan: 'ＭＳ Ｐゴシック', Hans: '\u9ed1\u4f53' })
    expect(t.majorScriptFonts).toEqual({ Jpan: 'ＭＳ Ｐゴシック', Hang: '굴림' })
  })

  it('+mn-ea with an empty ea slot takes the script font for the text language, else Latin', () => {
    const t = parseTheme(THEME(''))
    expect(resolveFontRef('+mn-ea', t, 'ja')).toBe('ＭＳ Ｐゴシック')
    expect(resolveFontRef('+mn-ea', t, 'sc')).toBe('\u9ed1\u4f53')
    expect(resolveFontRef('+mn-ea', t, 'ko')).toBe('Arial')
    expect(resolveFontRef('+mn-ea', t)).toBe('Arial')
    expect(resolveFontRef('+mj-ea', t, 'ko')).toBe('굴림')
  })

  it("'han' (ideographs without a language) resolves only when the theme names a single Han script", () => {
    const t = parseTheme(THEME(''))
    // minor: Jpan + Hans → ambiguous → Latin; major: Jpan only → Jpan
    expect(resolveFontRef('+mn-ea', t, 'han')).toBe('Arial')
    expect(resolveFontRef('+mj-ea', t, 'han')).toBe('ＭＳ Ｐゴシック')
  })

  it('an explicit ea typeface still wins over the script entry', () => {
    const t = parseTheme(THEME('Meiryo'))
    expect(resolveFontRef('+mn-ea', t, 'ja')).toBe('Meiryo')
  })

  it('maps lang tags to East Asian scripts', () => {
    expect(eaScriptOfLang('ja-JP')).toBe('ja')
    expect(eaScriptOfLang('zh-TW')).toBe('tc')
    expect(eaScriptOfLang('zh-CN')).toBe('sc')
    expect(eaScriptOfLang('en-US')).toBeUndefined()
  })
})
