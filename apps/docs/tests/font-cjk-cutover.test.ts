import { beforeEach, describe, expect, it } from 'vitest'
import {
  cssFontFamily,
  lineHeightFactor,
  setDocFontTable,
  setEmbeddedLineMetrics,
  simsunGapLineFactor,
  standInFontFamily,
} from '../src/renderer/line-metrics'
import { resolveDeclaredFace } from '../src/renderer/font-resolve'

/**
 * East Asian faces through the resolver. Word for Mac probe 2026-10-22
 * (40 single-spaced lines per face, pitch from pymupdf baselines, +-0.0006):
 * every face Word has paces at hhea (ascent + descent) x 1.3.
 */
describe('East Asian line factors follow hhea (ascent + descent) x 1.3', () => {
  beforeEach(() => {
    setDocFontTable([])
    setEmbeddedLineMetrics([])
  })

  it.each([
    ['Malgun Gothic', 1.7291, 1.7292],
    ['Yu Gothic', 1.4327, 1.433],
    ['Yu Mincho', 1.4327, 1.433],
    ['Meiryo', 1.95, 1.9505],
    ['Meiryo UI', 1.6511, 1.6505],
    ['Microsoft YaHei', 1.7157, 1.7151],
    ['Microsoft YaHei UI', 1.651, 1.6505],
    ['DengXian', 1.3546, 1.3545],
    ['Microsoft JhengHei', 1.7291, 1.7292],
    ['SimSun', 1.3, 1.297],
    ['MS Gothic', 1.3, 1.297],
    ['Batang', 1.3, 1.3009],
    ['STZhongsong', 1.7225, 1.7231],
    ['STSong', 1.3, 1.3003],
    ['Songti SC', 1.82, 1.8197],
    ['Yu Gothic UI', 1.7291, 1.7292],
    ['UD \u30c7\u30b8\u30bf\u30eb \u6559\u79d1\u66f8\u4f53 N-R', 1.5038, 1.5034],
  ])('%s -> %d (Word measured %d)', (name, factor, measured) => {
    expect(lineHeightFactor(name)).toBe(factor)
    expect(Math.abs(factor - measured)).toBeLessThan(0.0035)
  })

  it('missing names Word for Mac lays out in SimSun (probe 2026-10-22)', () => {
    // Hanyi Shusong, PingFang SC by its localized name, a fictional Song name
    for (const name of [
      '\u6c49\u4eea\u4e66\u5b8b\u4e8cKW',
      '\u82f9\u65b9-\u7b80',
      '\u67b6\u7a7a\u5b8b\u4f53',
    ]) {
      expect(resolveDeclaredFace(name)?.family.toLowerCase()).toBe('simsun')
      expect(lineHeightFactor(name)).toBe(1.3)
    }
  })

  it('the fictional Song name lifts kana-free text to SimSun only; gap glyphs go YaHei', () => {
    // word-truths L186 records 17.04-17.28 @10pt for "a i u abc" under the name:
    // text-driven, the name alone meters SimSun (probe 2026-10-22, Han text)
    expect(simsunGapLineFactor('\u67b6\u7a7a\u5b8b\u4f53')).toBe(1.7157)
  })

  it('Yu Gothic UI is its own face (word-truths L186: 17.28 @10pt), not Yu Gothic', () => {
    expect(resolveDeclaredFace('Yu Gothic UI')?.family).toBe('Yu Gothic UI')
    expect(lineHeightFactor('Yu Gothic UI') * 10).toBeCloseTo(17.28, 1)
  })

  it('UD Digital Kyokasho N-R is a cloud face (word-truths L186: 15.0 @10pt)', () => {
    expect(
      lineHeightFactor('UD \u30c7\u30b8\u30bf\u30eb \u6559\u79d1\u66f8\u4f53 N-R') * 10,
    ).toBeCloseTo(15.0, 1)
  })

  it('missing Arabic names substitute the installed Times New Roman (word-truths L449: 13.80 @12pt)', () => {
    expect(lineHeightFactor('Arabic Transparent') * 12).toBeCloseTo(13.8, 1)
    expect(lineHeightFactor('Noto Naskh Arabic')).toBe(1.1499)
  })

  it("Word's own Yu copy wins over the installed macOS one; a Hiragino JP name goes Yu Gothic", () => {
    expect(resolveDeclaredFace('Yu Gothic')?.source).toBe('dfonts')
    expect(lineHeightFactor('\u30d2\u30e9\u30ae\u30ce\u89d2\u30b4 Pro W3')).toBe(1.4327)
  })

  it('an embedded East Asian face is metered by its own hhea x 1.3', () => {
    setEmbeddedLineMetrics([
      { family: 'Malgun Gothic', ascent: 1.0781, descent: 0.252, lineGap: 0 },
    ])
    expect(lineHeightFactor('Malgun Gothic')).toBe(1.7291)
  })

  it('table exceptions stay: STXihei, Nanum assets, Noto JP Latin slot, Simplified Arabic', () => {
    expect(lineHeightFactor('STXihei')).toBe(1.79)
    expect(lineHeightFactor('NanumGothic')).toBe(1.495)
    expect(lineHeightFactor('Noto Sans CJK JP')).toBe(1.172)
    expect(lineHeightFactor('Simplified Arabic')).toBe(1.66)
  })
})

describe('CSS chains keyed by the target face', () => {
  it('a directly laid-out face takes its target chain; substitutions stay on the name branches', () => {
    expect(standInFontFamily('Calibri')).toBe(cssFontFamily('Calibri'))
    expect(standInFontFamily('\uff2d\uff33 \u660e\u671d')).toContain("'MS Mincho GO'")
    expect(standInFontFamily('Yu Gothic Light')).toContain("'Yu Gothic Light GO'")
    expect(standInFontFamily('Calibri Bold')).toBeNull()
    expect(standInFontFamily('Noto Sans CJK SC')).toBeNull()
  })
})
