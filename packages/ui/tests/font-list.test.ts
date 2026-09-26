// The shared candidate list and how the pickers split it (#864, PR A).
//
// The case that matters most is the one the maintainers called out: Chromium reports
// family names in the system's language, so a non-Chinese Windows enumerates `SimSun`
// while our candidate is `宋体`. Matching spellings literally would hide a font that
// is really installed.
import { describe, expect, it } from 'vitest'
import { BUILTIN_FONT_FAMILIES, fontFamiliesFor, partitionFontFamilies } from '../src/font-list'

// Nanum Myeongjo (the other Korean serif) is not a built-in candidate: it ships in the
// downloadable catalog, so the pickers surface it as a catalog row with a download marker
// (apps/slides/src/main/font-catalog.ts, covered by apps/slides/tests/font-store.test.ts).

describe('partitionFontFamilies', () => {
  it('keeps a localized candidate the machine reports under its English name', () => {
    // non-Chinese Windows: SimSun is enumerated, the localized name is not
    const { builtin } = partitionFontFamilies(
      ['宋体', '仿宋_GB2312', 'PingFang SC'],
      ['SimSun', 'FangSong_GB2312', 'Arial'],
    )
    expect(builtin).toContain('宋体')
    expect(builtin).toContain('仿宋_GB2312')
    // …while a family the machine genuinely lacks still drops out
    expect(builtin).not.toContain('PingFang SC')
  })

  it('matches the localized spelling too (Chinese Windows)', () => {
    const { builtin } = partitionFontFamilies(['宋体', 'SimSun'], ['宋体'])
    expect(builtin).toEqual(['宋体'])
  })

  it('keeps every candidate when the enumeration is unavailable or denied', () => {
    const { builtin, system } = partitionFontFamilies(['Calibri', '宋体', 'Hiragino Sans'], [])
    expect(builtin).toEqual(['Calibri', '宋体', 'Hiragino Sans'])
    expect(system).toEqual([])
  })

  it('keeps catalog families the enumeration cannot see through knownAvailable', () => {
    const { builtin, system } = partitionFontFamilies(
      ['Noto Serif SC'],
      ['Arial'],
      ['Noto Serif SC'],
    )
    expect(builtin).toContain('Noto Serif SC')
    expect(system).not.toContain('Noto Serif SC')
  })

  it('lists the machine families that are not candidates under system, in order', () => {
    const { system } = partitionFontFamilies(['Arial'], ['Arial', 'Comic Sans MS', 'Menlo'])
    expect(system).toEqual(['Comic Sans MS', 'Menlo'])
  })
})

describe('candidate lists', () => {
  it('carry Windows, macOS and downloadable names for the same script', () => {
    const sc = fontFamiliesFor('zh')
    expect(sc).toContain('宋体') // Windows
    expect(sc).toContain('PingFang SC') // macOS
    expect(sc).toContain('Songti SC') // macOS serif
    expect(sc).toContain('Noto Serif SC') // downloadable
  })

  it('offer a serif next to the sans for every CJK script', () => {
    for (const [lang, sans, serif] of [
      ['zh', 'Noto Sans SC', 'Noto Serif SC'],
      ['zh-TW', 'Noto Sans TC', 'Noto Serif TC'],
      ['ja', 'Noto Sans JP', 'Noto Serif JP'],
      ['ko', 'Noto Sans KR', 'Noto Serif KR'],
    ] as const) {
      const list = fontFamiliesFor(lang)
      expect(list).toContain(sans)
      expect(list).toContain(serif)
    }
  })

  it('has no duplicate candidates in the merged list', () => {
    expect(new Set(BUILTIN_FONT_FAMILIES).size).toBe(BUILTIN_FONT_FAMILIES.length)
  })
})
