import { describe, expect, it } from 'vitest'

import {
  CELL_FONT_ALIASES,
  rewriteScopedFamilies,
  withSansSerifFallback,
} from '../src/renderer/cell-font-fallback'

const EMOJI = '"Cell Text Dingbats", "Apple Color Emoji", "Segoe UI Emoji", "Noto Color Emoji"'

describe('withSansSerifFallback', () => {
  it('appends sans-serif to a bare family string', () => {
    expect(withSansSerifFallback('11pt "ＭＳ Ｐゴシック"')).toBe(
      `11pt "ＭＳ Ｐゴシック", sans-serif, ${EMOJI}`,
    )
    expect(withSansSerifFallback('bold 14.6667px "Aptos Narrow"')).toBe(
      `bold 14.6667px "Aptos Narrow", sans-serif, ${EMOJI}`,
    )
    expect(withSansSerifFallback('italic bold 11pt Carlito')).toBe(
      `italic bold 11pt Carlito, sans-serif, ${EMOJI}`,
    )
  })

  it('appends after a multi-family list without a generic', () => {
    expect(withSansSerifFallback('12px "Yu Gothic", Meiryo')).toBe(
      `12px "Yu Gothic", Meiryo, sans-serif, ${EMOJI}`,
    )
  })

  it('appends only the emoji faces to strings already ending in a generic', () => {
    expect(withSansSerifFallback('12px Arial, sans-serif')).toBe(`12px Arial, sans-serif, ${EMOJI}`)
    expect(withSansSerifFallback('12px serif')).toBe(`12px serif, ${EMOJI}`)
    expect(withSansSerifFallback('10px monospace')).toBe(`10px monospace, ${EMOJI}`)
    const univerPlane =
      '11pt Arial, "Helvetica Neue", Helvetica, Arial, "PingFang SC", "Hiragino Sans GB", "Heiti SC", "Microsoft YaHei", "WenQuanYi Micro Hei", sans-serif '
    expect(withSansSerifFallback(univerPlane)).toBe(`${univerPlane.trimEnd()}, ${EMOJI}`)
  })

  it('is idempotent: never re-appends to an already-expanded chain', () => {
    const once = withSansSerifFallback('12px Arial, sans-serif')
    expect(withSansSerifFallback(once)).toBe(once)
    const serifOnce = withSansSerifFallback('11pt "MS Mincho"')
    expect(withSansSerifFallback(serifOnce)).toBe(serifOnce)
  })

  it('appends serif (not sans-serif) for alias-known serif family names', () => {
    expect(withSansSerifFallback('11pt "MS Mincho"')).toBe(`11pt "MS Mincho", serif, ${EMOJI}`)
    expect(withSansSerifFallback('bold 12px SimSun')).toBe(`bold 12px SimSun, serif, ${EMOJI}`)
    expect(withSansSerifFallback('12px Batang')).toBe(`12px Batang, serif, ${EMOJI}`)
    expect(withSansSerifFallback('12px "ＭＳ Ｐ明朝"')).toBe(`12px "ＭＳ Ｐ明朝", serif, ${EMOJI}`)
    expect(withSansSerifFallback('12px "Times New Roman"')).toBe(
      `12px "Times New Roman", serif, ${EMOJI}`,
    )
    expect(withSansSerifFallback('12px "PT Serif"')).toBe(`12px "PT Serif", serif, ${EMOJI}`)
    // Office-for-Mac DFont faces that Excel renders as a serif
    expect(withSansSerifFallback('12pt "Baskerville Old Face"')).toBe(
      `12pt "Baskerville Old Face", serif, ${EMOJI}`,
    )
    expect(withSansSerifFallback('12pt "Gill Sans MT"')).toBe(
      `12pt "Gill Sans MT", sans-serif, ${EMOJI}`,
    )
  })

  it('keeps sans-serif for sans families whose names contain "serif"', () => {
    expect(withSansSerifFallback('12px "Microsoft Sans Serif"')).toBe(
      `12px "Microsoft Sans Serif", sans-serif, ${EMOJI}`,
    )
  })

  // Excel substitutes its sans default for names it cannot resolve — even
  // myeongjo/mincho-keyworded ones. Hancom composite chains and the single
  // names Univer's per-glyph fallback re-probes must both come out sans
  // (Excel reference).
  it('keeps sans-serif for unrecognized names, keyworded or composite', () => {
    expect(withSansSerifFallback('12px 휴먼명조')).toBe(`12px 휴먼명조, sans-serif, ${EMOJI}`)
    expect(withSansSerifFallback('bold 20pt 휴먼명조, 한컴돋움')).toBe(
      `bold 20pt 휴먼명조, 한컴돋움, sans-serif, ${EMOJI}`,
    )
    expect(withSansSerifFallback('11pt HY그래픽B, 한컴돋움')).toBe(
      `11pt HY그래픽B, 한컴돋움, sans-serif, ${EMOJI}`,
    )
  })

  it('lets the first alias-known member of a list carry its intent', () => {
    expect(withSansSerifFallback('12px 바탕, 한컴돋움')).toBe(
      `12px 바탕, 한컴돋움, serif, ${EMOJI}`,
    )
    expect(withSansSerifFallback('12px "ＭＳ 明朝", Unknown')).toBe(
      `12px "ＭＳ 明朝", Unknown, serif, ${EMOJI}`,
    )
    expect(withSansSerifFallback('12px Unknown, "Yu Gothic"')).toBe(
      `12px Unknown, "Yu Gothic", sans-serif, ${EMOJI}`,
    )
  })

  it('passes empty values through', () => {
    expect(withSansSerifFallback('')).toBe('')
    expect(withSansSerifFallback('   ')).toBe('   ')
  })
})

describe('CELL_FONT_ALIASES', () => {
  it('has unique families and non-empty chains', () => {
    const families = CELL_FONT_ALIASES.map((a) => a.family)
    expect(new Set(families).size).toBe(families.length)
    for (const alias of CELL_FONT_ALIASES) {
      expect(alias.regular.length).toBeGreaterThan(0)
      if (alias.bold) expect(alias.bold.length).toBeGreaterThan(0)
    }
  })

  it('lists the genuine bold face first where the family has a real bold', () => {
    const expectations: Record<string, string> = {
      Cambria: 'Cambria Bold',
      Garamond: 'Garamond Bold',
    }
    for (const [family, face] of Object.entries(expectations)) {
      const alias = CELL_FONT_ALIASES.find((a) => a.family === family)
      expect(alias?.bold?.[0], family).toBe(face)
    }
    // Gated JP families keep the genuine bold in the rename map used when
    // the real font exists; the substitute chain is Hiragino-only.
    const gated: Record<string, string> = {
      Meiryo: 'Meiryo Bold',
      メイリオ: 'Meiryo Bold',
      'Meiryo UI': 'Meiryo UI Bold',
      'Yu Gothic': 'Yu Gothic Bold',
      'Yu Gothic UI': 'Yu Gothic UI Bold',
    }
    for (const [family, face] of Object.entries(gated)) {
      const alias = CELL_FONT_ALIASES.find((a) => a.family === family)
      expect(alias?.whenGenuine?.bold?.[0], family).toBe(face)
      expect(alias?.bold?.[0], family).toBe('HiraginoSans-W6')
    }
  })

  it('gives Yu Gothic / Yu Mincho bold chains a Hiragino fallback', () => {
    for (const family of ['游ゴシック', '游ゴシック体', 'Yu Gothic', 'Yu Gothic UI']) {
      const alias = CELL_FONT_ALIASES.find((a) => a.family === family)
      expect(alias?.bold, family).toContain('HiraginoSans-W6')
    }
    for (const family of ['游明朝', 'Yu Mincho']) {
      const alias = CELL_FONT_ALIASES.find((a) => a.family === family)
      expect(alias?.bold, family).toContain('HiraMinProN-W6')
    }
  })

  it('never lists a regular face as a bold source (would suppress synthetic bold)', () => {
    for (const alias of CELL_FONT_ALIASES) {
      if (alias.bold) for (const face of alias.bold) expect(alias.regular).not.toContain(face)
      if (alias.latin?.bold)
        for (const face of alias.latin.bold) expect(alias.latin.regular).not.toContain(face)
      if (alias.whenGenuine?.bold)
        for (const face of alias.whenGenuine.bold)
          expect(alias.whenGenuine.regular).not.toContain(face)
    }
  })

  it('pairs every size-adjusted substitute with a genuine-font gate', () => {
    for (const alias of CELL_FONT_ALIASES) {
      const adjusted =
        alias.sizeAdjust ??
        alias.boldSizeAdjust ??
        alias.latin?.sizeAdjust ??
        alias.latin?.boldSizeAdjust
      if (adjusted) expect(alias.skipIfLocal?.length, alias.family).toBeGreaterThan(0)
    }
  })

  it('width-corrects the missing-on-macOS families from the prod clusters', () => {
    for (const family of ['Bahnschrift', 'Segoe UI', 'Dosis', 'Aptos Narrow']) {
      const alias = CELL_FONT_ALIASES.find((a) => a.family === family)
      expect(alias?.sizeAdjust, family).toMatch(/^\d+(\.\d+)?%$/)
    }
    for (const family of [
      'Avenir Next LT Pro',
      'Avenir Next LT Pro Demi',
      'Avenir Next LT Pro Light',
    ]) {
      expect(
        CELL_FONT_ALIASES.some((a) => a.family === family),
        family,
      ).toBe(true)
    }
  })

  it('splits Malgun Gothic per script: hangul stays exact, latin is corrected', () => {
    for (const family of ['Malgun Gothic', '맑은 고딕']) {
      const alias = CELL_FONT_ALIASES.find((a) => a.family === family)
      expect(alias?.regular, family).toContain('AppleGothic')
      expect(alias?.sizeAdjust, family).toBeUndefined()
      expect(alias?.latin?.sizeAdjust, family).toBe('104%')
      // malgunbd.ttf digits 0.5796em over Helvetica Neue Bold 0.556em.
      expect(pct(alias?.latin?.boldSizeAdjust), family).toBeCloseTo(104.2, 1)
      expect(alias?.skipIfLocal, family).toContain('Malgun Gothic')
    }
  })

  it('keeps bold Malgun hangul in the family at Excel width', () => {
    // AppleGothic has no bold face; Apple SD Gothic Neo Bold sets hangul at
    // 0.865em where Malgun Gothic Bold sets 1.0em.
    for (const family of ['Malgun Gothic', '맑은 고딕']) {
      const alias = CELL_FONT_ALIASES.find((a) => a.family === family)
      expect(alias?.bold?.[0], family).toBe('Apple SD Gothic Neo Bold')
      expect(alias?.bold, family).not.toContain('Malgun Gothic Bold')
      expect(pct(alias?.boldSizeAdjust), family).toBeCloseTo(100 / 0.865, 0)
    }
  })

  it('pairs every Latin bold sub-face with a base bold face', () => {
    // Chromium selects a family's weight-700 faces before unicode-range, so
    // a Latin-only bold face sends bold non-Latin text to the system fallback.
    for (const alias of CELL_FONT_ALIASES) {
      if (alias.latin?.bold) expect(alias.bold?.length, alias.family).toBeGreaterThan(0)
    }
  })

  const pct = (value: string | undefined): number => {
    expect(value).toMatch(/^\d+(\.\d+)?%$/)
    return Number.parseFloat(value!)
  }

  it('width-corrects the Thai Office faces per script', () => {
    // Cordia New draws digits at 0.3645em and Thai at ~0.30em; Thonburi is
    // 0.666em / ~0.44em, so both scripts need their own size-adjust and the
    // Latin sub-face must leave the Thai block (inside U+0-2CFF) alone.
    const families = [
      'Cordia New',
      'CordiaUPC',
      'Angsana New',
      'AngsanaUPC',
      'TH SarabunPSK',
      'TH Sarabun New',
    ]
    for (const family of families) {
      const alias = CELL_FONT_ALIASES.find((a) => a.family === family)
      expect(alias, family).toBeDefined()
      expect(alias!.regular, family).toEqual(['Thonburi'])
      expect(alias!.bold, family).toContain('Thonburi-Bold')
      expect(alias!.skipIfLocal, family).toContain(family)
      const thai = pct(alias!.sizeAdjust)
      const thaiBold = pct(alias!.boldSizeAdjust)
      expect(thai, family).toBeGreaterThan(60)
      expect(thai, family).toBeLessThan(80)
      expect(thaiBold, family).toBeGreaterThan(60)
      expect(thaiBold, family).toBeLessThan(80)
      const latin = alias!.latin
      expect(latin, family).toBeDefined()
      expect(latin!.unicodeRange, family).toBe('U+0-DFF, U+E80-2CFF')
      const latinAdjust = pct(latin!.sizeAdjust)
      expect(latinAdjust, family).toBeGreaterThan(55)
      expect(latinAdjust, family).toBeLessThan(75)
      expect(pct(latin!.boldSizeAdjust), family).toBeGreaterThan(55)
      expect(latin!.bold?.length, family).toBeGreaterThan(0)
    }
    // Cordia / Sarabun Latin is a narrow sans (Helvetica Neue); Angsana's
    // Latin is Times New Roman at 66% and keeps the serif design.
    expect(CELL_FONT_ALIASES.find((a) => a.family === 'Cordia New')?.latin?.regular).toEqual([
      'Helvetica Neue',
    ])
    const angsana = CELL_FONT_ALIASES.find((a) => a.family === 'Angsana New')
    expect(angsana?.latin?.regular).toEqual(['Times New Roman'])
    expect(angsana?.latin?.sizeAdjust).toBe('66%')
    expect(angsana?.latin?.bold).toContain('Times New Roman Bold')
    // The UPC spellings are the same designs.
    for (const [a, b] of [
      ['Cordia New', 'CordiaUPC'],
      ['Angsana New', 'AngsanaUPC'],
      ['TH SarabunPSK', 'TH Sarabun New'],
    ]) {
      const first = CELL_FONT_ALIASES.find((x) => x.family === a)!
      const second = CELL_FONT_ALIASES.find((x) => x.family === b)!
      expect(second.sizeAdjust, b).toBe(first.sizeAdjust)
      expect(second.latin?.sizeAdjust, b).toBe(first.latin?.sizeAdjust)
    }
  })

  it('narrows the Latin runs of the JP gothic substitutes to the JIS half-width metrics', () => {
    // MS (P/UI) Gothic digits are 0.5em vs Helvetica Neue 0.556em (≈90%);
    // kana/kanji stay on the unadjusted Hiragino base face.
    for (const family of [
      'ＭＳ Ｐゴシック',
      'MS PGothic',
      'ＭＳ ゴシック',
      'MS Gothic',
      'MS UI Gothic',
    ]) {
      const alias = CELL_FONT_ALIASES.find((a) => a.family === family)
      expect(alias, family).toBeDefined()
      expect(alias!.regular[0], family).toBe('Hiragino Sans')
      expect(alias!.sizeAdjust, family).toBeUndefined()
      expect(alias!.boldSizeAdjust, family).toBeUndefined()
      expect(alias!.latin?.regular, family).toEqual(['Helvetica Neue'])
      expect(alias!.latin?.unicodeRange, family).toBeUndefined()
      const adjust = pct(alias!.latin?.sizeAdjust)
      expect(adjust, family).toBeGreaterThan(88)
      expect(adjust, family).toBeLessThan(93)
      const boldAdjust = pct(alias!.latin?.boldSizeAdjust)
      expect(boldAdjust, family).toBeGreaterThan(88)
      expect(boldAdjust, family).toBeLessThan(93)
      expect(alias!.skipIfLocal?.length, family).toBeGreaterThan(0)
      // The fullwidth spellings stay resolvable when the genuine font exists.
      expect(alias!.whenGenuine?.regular[0], family).toMatch(/^MS /)
    }
    // Yu Gothic's Latin already matches Helvetica Neue (0.5562em digits) —
    // the fix is leaving Hiragino's 0.657em digits, not the adjustment.
    for (const family of ['游ゴシック', '游ゴシック体', 'Yu Gothic']) {
      const alias = CELL_FONT_ALIASES.find((a) => a.family === family)
      expect(alias?.latin?.regular, family).toEqual(['Helvetica Neue'])
      expect(pct(alias?.latin?.sizeAdjust), family).toBeCloseTo(100.8, 5)
      expect(pct(alias?.latin?.boldSizeAdjust), family).toBeGreaterThan(100)
      expect(alias?.whenGenuine?.regular, family).toContain('YuGothic-Regular')
    }
    // Meiryo's Latin is the Verdana design.
    for (const family of ['メイリオ', 'Meiryo', 'Meiryo UI']) {
      const alias = CELL_FONT_ALIASES.find((a) => a.family === family)
      expect(alias?.latin?.regular, family).toEqual(['Verdana'])
      expect(alias?.latin?.bold?.[0], family).toBe('Verdana Bold')
      const adjust = pct(alias?.latin?.sizeAdjust)
      expect(adjust, family).toBeGreaterThan(94)
      expect(adjust, family).toBeLessThan(100)
    }
  })

  it('keeps chrome-stack families out of document.fonts via canvas scoping', () => {
    // The UI stack starts with 'Segoe UI'; an unscoped size-adjusted face
    // would restyle the ribbon on hosts without the genuine font.
    const alias = CELL_FONT_ALIASES.find((a) => a.family === 'Segoe UI')
    expect(alias?.scopeToCanvas).toBe(true)
  })

  it('rewrites a sole scoped cell family but never an explicit UI fallback stack', () => {
    const scoped = new Map([['segoe ui', '__cell-scope Segoe UI']])
    expect(rewriteScopedFamilies('italic bold 11pt "Segoe UI"', scoped)).toBe(
      'italic bold 11pt "__cell-scope Segoe UI"',
    )
    expect(rewriteScopedFamilies('16px "Segoe UI", monospace', scoped)).toBe(
      '16px "__cell-scope Segoe UI", monospace',
    )
    // UI measurements mirror a CSS stack with real fallbacks — they must fall
    // through natively like the DOM they match (truncateCardName).
    expect(
      rewriteScopedFamilies(
        "500 13px 'Segoe UI', -apple-system, BlinkMacSystemFont, 'PingFang SC', sans-serif",
        scoped,
      ),
    ).toBe("500 13px 'Segoe UI', -apple-system, BlinkMacSystemFont, 'PingFang SC', sans-serif")
    expect(rewriteScopedFamilies('16px Carlito', scoped)).toBe('16px Carlito')
  })

  it('keeps a plain rename mapping for the Hangul Malgun spelling when the genuine font exists', () => {
    // The OS matcher never resolves localized family names (probed on macOS),
    // so skipping the alias entirely would drop '맑은 고딕' to the sans
    // fallback on hosts that do have Malgun Gothic.
    const alias = CELL_FONT_ALIASES.find((a) => a.family === '맑은 고딕')
    expect(alias?.whenGenuine?.regular).toContain('Malgun Gothic')
    expect(alias?.whenGenuine?.bold).toContain('Malgun Gothic Bold')
  })

  it('backs Dosis and Aptos Narrow with the bundled Carlito, not local()-only', () => {
    for (const family of ['Dosis', 'Aptos Narrow']) {
      const alias = CELL_FONT_ALIASES.find((a) => a.family === family)
      expect(
        alias?.regular.some((s) => s.startsWith('url(')),
        family,
      ).toBe(true)
      expect(
        alias?.bold?.some((s) => s.startsWith('url(')),
        family,
      ).toBe(true)
    }
  })

  it('maps the Office-for-Mac DFonts onto the stock macOS designs', () => {
    const baskerville = CELL_FONT_ALIASES.find((a) => a.family === 'Baskerville Old Face')
    expect(baskerville?.regular[0]).toBe('Baskerville Old Face')
    expect(baskerville?.regular).toContain('Baskerville')
    expect(baskerville?.sizeAdjust).toBeUndefined()
    // No genuine bold exists: never hand bold cells to a Times face where the
    // regular resolves to the real font (Windows) — synthetic bold must win.
    expect(baskerville?.bold?.some((f) => /Times/.test(f))).toBe(false)
    const gill = CELL_FONT_ALIASES.find((a) => a.family === 'Gill Sans MT')
    expect(gill?.regular[0]).toBe('Gill Sans MT')
    expect(gill?.regular).toContain('Gill Sans')
    expect(gill?.bold?.[0]).toBe('Gill Sans MT Bold')
  })

  const find = (family: string) => CELL_FONT_ALIASES.find((a) => a.family === family)

  it('resolves the bundled Carlito through the shared package asset, not a dead relative path', () => {
    for (const family of ['Dosis', 'Aptos Narrow']) {
      const alias = find(family)
      const urls = [...alias!.regular, ...alias!.bold!].filter((s) => s.startsWith('url('))
      expect(urls, family).toHaveLength(2)
      for (const url of urls) {
        expect(url, family).toMatch(/Carlito-(Regular|Bold)[^)]*\.ttf\)$/)
        expect(url, family).not.toContain('./fonts/')
      }
    }
  })

  it('renames the localized BIZ UD spellings onto the installed English family', () => {
    const gothic = find('BIZ UD\u30b4\u30b7\u30c3\u30af')
    expect(gothic?.regular[0]).toBe('BIZ UDGothic')
    expect(gothic?.regular).toContain('Hiragino Sans')
    expect(gothic?.bold?.[0]).toBe('BIZ UDGothic Bold')
    expect(gothic?.sizeAdjust).toBeUndefined()
    // The proportional twin falls to the fixed-pitch design before Hiragino.
    const pGothic = find('BIZ UDP\u30b4\u30b7\u30c3\u30af')
    expect(pGothic?.regular.slice(0, 3)).toEqual([
      'BIZ UDPGothic',
      'BIZUDPGothic-Regular',
      'BIZ UDGothic',
    ])
    for (const family of ['BIZ UD\u660e\u671d', 'BIZ UDP\u660e\u671d']) {
      expect(find(family)?.regular, family).toContain('Hiragino Mincho ProN')
      expect(withSansSerifFallback(`11pt "${family}"`)).toBe(`11pt "${family}", serif, ${EMOJI}`)
    }
    expect(withSansSerifFallback('11pt "BIZ UD\u30b4\u30b7\u30c3\u30af"')).toBe(
      `11pt "BIZ UD\u30b4\u30b7\u30c3\u30af", sans-serif, ${EMOJI}`,
    )
  })

  it('draws the heavy-by-name HG faces from a heavy weight even without <b/>', () => {
    const soeiUB = '\u5275\u82f1\u89d2\uff7a\uff9e\uff7c\uff6f\uff78UB'
    for (const prefix of ['HGP', 'HGS', 'HG']) {
      const alias = find(`${prefix}${soeiUB}`)
      expect(alias?.regular[0], prefix).toMatch(/SoeiKakugothicUB$/)
      expect(alias?.regular[1], prefix).toBe('Hiragino Sans W8')
      expect(alias?.regular, prefix).toContain('HiraginoSans-W6')
      expect(alias?.bold?.[0], prefix).toBe('Hiragino Sans W9')
    }
    const gothicE = find('HG\uff7a\uff9e\uff7c\uff6f\uff78E')
    expect(gothicE?.regular.slice(0, 2)).toEqual(['HGGothicE', 'Hiragino Sans W7'])
    const minchoE = find('HG\u660e\u671dE')
    expect(minchoE?.regular).toEqual(['HGMinchoE', 'HiraMinProN-W6', 'Hiragino Mincho ProN W6'])
    expect(minchoE?.bold).toBeUndefined()
    expect(withSansSerifFallback('12pt "HG\u660e\u671dE"')).toBe(
      `12pt "HG\u660e\u671dE", serif, ${EMOJI}`,
    )
    const maru = find('HG\u4e38\uff7a\uff9e\uff7c\uff6f\uff78M-PRO')
    expect(maru?.regular.slice(0, 2)).toEqual(['HGMaruGothicMPRO', 'Hiragino Maru Gothic ProN'])
  })

  it('pins the Korean fixed-pitch twins to half-width Latin over an exact-hangul base', () => {
    for (const [family, genuine] of [
      ['GulimChe', 'GulimChe'],
      ['\uad74\ub9bc\uccb4', 'GulimChe'],
      ['DotumChe', 'DotumChe'],
      ['\ub3cb\uc6c0\uccb4', 'DotumChe'],
    ]) {
      const alias = find(family!)
      expect(alias?.regular[0], family).toBe('AppleGothic')
      expect(alias?.sizeAdjust, family).toBeUndefined()
      // 0.5em digits over Helvetica Neue's 0.556em, as for MS Gothic.
      expect(pct(alias?.latin?.sizeAdjust), family).toBeCloseTo(89.9, 5)
      expect(alias?.bold?.[0], family).toBe('Apple SD Gothic Neo Bold')
      expect(alias?.skipIfLocal, family).toEqual([genuine])
      expect(alias?.whenGenuine?.regular, family).toEqual([genuine])
      expect(withSansSerifFallback(`10pt "${family}"`)).toBe(
        `10pt "${family}", sans-serif, ${EMOJI}`,
      )
    }
    for (const family of ['BatangChe', '\ubc14\ud0d5\uccb4']) {
      const alias = find(family)
      expect(alias?.regular, family).toContain('AppleMyungjo')
      // Times New Roman digits are exactly 0.5em — no size-adjust needed.
      expect(alias?.latin?.regular, family).toEqual(['Times New Roman'])
      expect(alias?.latin?.sizeAdjust, family).toBeUndefined()
      expect(alias?.skipIfLocal, family).toEqual(['BatangChe'])
      expect(withSansSerifFallback(`10pt "${family}"`)).toBe(`10pt "${family}", serif, ${EMOJI}`)
    }
  })

  it('keeps serif intent for the cloud-only Google serif faces', () => {
    const expectations: Record<string, string> = {
      'Playfair Display': 'Didot',
      'EB Garamond': 'Garamond',
      Merriweather: 'Georgia',
      Lora: 'Georgia',
      'Libre Baskerville': 'Baskerville',
    }
    for (const [family, standIn] of Object.entries(expectations)) {
      const alias = find(family)
      expect(alias?.regular[0], family).toBe(family)
      expect(alias?.regular[1], family).toBe(standIn)
      expect(alias?.bold?.[0], family).toBe(`${family} Bold`)
      expect(withSansSerifFallback(`bold 11pt "${family}"`)).toBe(
        `bold 11pt "${family}", serif, ${EMOJI}`,
      )
    }
  })

  it('treats the Adobe Kozuka Mincho PostScript names as mincho', () => {
    for (const base of ['KozMinPro', 'KozMinPr6N']) {
      for (const weight of ['Regular', 'Medium']) {
        const family = `${base}-${weight}`
        const alias = find(family)
        expect(alias?.regular, family).toEqual([family, 'Hiragino Mincho ProN', 'HiraMinProN-W3'])
        expect(alias?.bold?.[0], family).toBe('HiraMinProN-W6')
        expect(withSansSerifFallback(`10pt ${family}`)).toBe(`10pt ${family}, serif, ${EMOJI}`)
      }
      const bold = find(`${base}-Bold`)
      expect(bold?.regular.slice(0, 2)).toEqual([`${base}-Bold`, 'HiraMinProN-W6'])
      expect(bold?.bold).toBeUndefined()
    }
    // Unrecognized gothic siblings keep Excel's sans substitution.
    expect(withSansSerifFallback('10pt KozGoPro-Regular')).toBe(
      `10pt KozGoPro-Regular, sans-serif, ${EMOJI}`,
    )
  })

  it('keeps serif intent for mincho/song/ming/batang names', () => {
    const serifFaces =
      /Mincho|Song|Myungjo|Myeongjo|LiSung|Times|Georgia|Palatino|Antiqua|PMingLiU|MingLiU/
    for (const family of ['ＭＳ 明朝', '游明朝', '宋体', 'SimSun', '新細明體', 'Batang', '바탕']) {
      const alias = CELL_FONT_ALIASES.find((a) => a.family === family)
      expect(alias, family).toBeDefined()
      expect(
        alias!.regular.some((f) => serifFaces.test(f)),
        family,
      ).toBe(true)
    }
  })
})
