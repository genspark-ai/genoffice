import { describe, expect, it } from 'vitest'

import { CELL_FONT_ALIASES, withSansSerifFallback } from '../src/renderer/cell-font-fallback'

const EMOJI = '"Apple Color Emoji", "Segoe UI Emoji", "Noto Color Emoji"'

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
      Meiryo: 'Meiryo Bold',
      メイリオ: 'Meiryo Bold',
      'Meiryo UI': 'Meiryo UI Bold',
      'Yu Gothic': 'Yu Gothic Bold',
      'Yu Gothic UI': 'Yu Gothic UI Bold',
    }
    for (const [family, face] of Object.entries(expectations)) {
      const alias = CELL_FONT_ALIASES.find((a) => a.family === family)
      expect(alias?.bold?.[0], family).toBe(face)
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
      if (!alias.bold) continue
      for (const face of alias.bold) expect(alias.regular).not.toContain(face)
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
