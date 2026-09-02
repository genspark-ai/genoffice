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
  })

  it('keeps sans-serif for sans families whose names contain "serif"', () => {
    expect(withSansSerifFallback('12px "Microsoft Sans Serif"')).toBe(
      `12px "Microsoft Sans Serif", sans-serif, ${EMOJI}`,
    )
  })

  // Excel substitutes its sans default for names it cannot resolve — even
  // myeongjo/mincho-keyworded ones. Hancom composite chains and the single
  // names Univer's per-glyph fallback re-probes must both come out sans
  // (prod_059 Excel reference).
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
      expect(alias?.latin?.boldSizeAdjust, family).toBe('109.4%')
      expect(alias?.skipIfLocal, family).toContain('Malgun Gothic')
    }
  })

  const pct = (value: string | undefined): number => {
    expect(value).toMatch(/^\d+(\.\d+)?%$/)
    return Number.parseFloat(value!)
  }

  it('width-corrects the Thai Office faces per script (prod_066)', () => {
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
