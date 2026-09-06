/**
 * Meiryo / Meiryo UI metric aliases (fonts.css): Word renders the real faces
 * (probe 2026-09-03, advances read from Word's meiryo.ttc). Kana and JP
 * punctuation ride the bundled GenOffice UI Kana JP faces (Noto Sans JP
 * outlines condensed to Meiryo UI's exact per-glyph advances at full height;
 * tools/build-meiryo-ui-kana-font.py) and Latin rides a Verdana size-adjust,
 * so text advances like Meiryo instead of 1em / Helvetica widths.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { advanceEm, readWoff2 } from './helpers/woff2-metrics'

const FONTS = join(__dirname, '../src/renderer/fonts')
const css = readFileSync(join(FONTS, 'fonts.css'), 'utf8')
const KANA_RANGE = 'U+3000-3002, U+3008-3011, U+3014-3015, U+3041-3096, U+3099-30FF'

function faces(family: string) {
  return [...css.matchAll(/@font-face \{([^}]*)\}/g)]
    .map((m) => m[1])
    .filter((body) => body.includes(`font-family: '${family}'`))
    .map((body) => ({
      range: /unicode-range:\s*([^;]+);/.exec(body)?.[1].replace(/\s+/g, ' ').trim(),
      adjust: /size-adjust:\s*([\d.]+)%/.exec(body)?.[1],
      bold: /font-weight: bold/.test(body),
      src: /src:\s*([^;]+);/.exec(body)?.[1] ?? '',
    }))
}

describe('Meiryo UI GO', () => {
  const rules = faces('Meiryo UI GO')
  const find = (range: string, bold = false) =>
    rules.find((r) => r.range === range && r.bold === bold)

  it('serves kana and JP punctuation from the bundled faces without size-adjust', () => {
    expect(find(KANA_RANGE)?.src).toBe("url('./GenOfficeUIKanaJP-Regular.woff2') format('woff2')")
    expect(find(KANA_RANGE, true)?.src).toBe(
      "url('./GenOfficeUIKanaJP-Bold.woff2') format('woff2')",
    )
    expect(find(KANA_RANGE)?.adjust).toBeUndefined()
    expect(find(KANA_RANGE, true)?.adjust).toBeUndefined()
    expect(rules.filter((r) => /Hiragino/.test(r.src) && !/Verdana/.test(r.src))).toHaveLength(0)
  })

  it('leaves 1em glyphs (kanji, fullwidth parens/colon) to the chain', () => {
    for (const r of rules) {
      expect(r.range).not.toMatch(/U\+4E00|U\+FF08|U\+FF1A/)
    }
  })

  it('Latin range is Verdana-based with a Hiragino backstop, PostScript names for local()', () => {
    const latin = find('U+0020-007E, U+00A0-024F')
    expect(latin?.adjust).toBe('97.05')
    expect(latin?.src).toContain("local('Verdana')")
    expect(latin?.src).toContain("local('HiraginoSans-W3')")
    expect(find('U+0020-007E, U+00A0-024F', true)?.adjust).toBe('94.8')
  })
})

describe('GenOffice UI Kana JP (Meiryo UI advances, full height)', () => {
  const regular = readWoff2(join(FONTS, 'GenOfficeUIKanaJP-Regular.woff2'))
  const bold = readWoff2(join(FONTS, 'GenOfficeUIKanaJP-Bold.woff2'))
  const near = (font: typeof regular, cp: number, em: number) =>
    expect(advanceEm(font, cp)).toBeCloseTo(em, 3)

  it('regular kana carry Meiryo UI regular advances', () => {
    near(regular, 0x3042, 0.8164) // あ
    near(regular, 0x3046, 0.6392) // う
    near(regular, 0x30a2, 0.7539) // ア
    near(regular, 0x30fc, 0.8164) // ー
    near(regular, 0x30fd, 0.7461) // ヽ
  })

  it('bold kana carry the wider Meiryo UI Bold advances', () => {
    near(bold, 0x3042, 0.8999)
    near(bold, 0x3046, 0.7134)
    near(bold, 0x30a2, 0.8247)
    near(bold, 0x30fc, 0.8818)
  })

  it('ideographic space, comma/period and brackets match at both weights', () => {
    for (const font of [regular, bold]) {
      near(font, 0x3000, 0.6641)
      near(font, 0x3001, 0.6641)
      near(font, 0x3002, 0.6641)
      near(font, 0x300c, 0.5)
      near(font, 0x300e, 0.5)
      near(font, 0x3010, 0.5)
    }
    near(regular, 0x30fb, 0.5)
    near(bold, 0x30fb, 0.5205)
  })

  it('keeps the Hiragino-class vertical metrics so mixed lines share a baseline box', () => {
    for (const font of [regular, bold]) {
      const hhea = font.tables.get('hhea')!
      expect(hhea.readInt16BE(4) / font.unitsPerEm).toBeCloseTo(0.88, 3)
      expect(hhea.readInt16BE(6) / font.unitsPerEm).toBeCloseTo(-0.12, 3)
    }
  })
})

describe('Meiryo GO', () => {
  it('only rescales Latin (Meiryo kana/kanji/punctuation are 1em like Hiragino)', () => {
    const rules = faces('Meiryo GO')
    expect(rules).toHaveLength(2)
    expect(rules.every((r) => r.range === 'U+0020-007E, U+00A0-024F')).toBe(true)
    expect(rules.map((r) => r.adjust).sort()).toEqual(['94.8', '97.05'])
  })
})
