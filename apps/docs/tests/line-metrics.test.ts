/**
 * Unit tests for the F1 line-box metrics engine
 *
 * Covers:
 *   - HeuristicMetrics measurements
 *   - computeLineHeight's three line-height modes
 *   - snapSpacingToGrid grid rounding
 *   - simulateLines line-wrapping simulation
 *   - computeLineMetrics main entry (various font sizes/rules/grids)
 */

import { describe, it, expect } from 'vitest'
import {
  HeuristicMetrics,
  computeLineHeight,
  cssFontFamily,
  cssLineHeight,
  snapSpacingToGrid,
  simulateLines,
  computeLineMetrics,
  textHasCjk,
} from '../src/renderer/line-metrics'

const TWIPS_TO_PX = 96 / 1440

// ─── HeuristicMetrics ─────────────────────────────────────────────────────

describe('HeuristicMetrics', () => {
  const m = new HeuristicMetrics()

  it('natural line height of 12pt text is about 14.4px (1.2em)', () => {
    const fontSizePx = 12 * (96 / 72) // 16px
    const style = { fontFamily: 'Arial', fontSizePx, bold: false, italic: false }
    const metrics = m.metrics(style)
    expect(metrics.lineHeight).toBeCloseTo(fontSizePx * 1.2, 1)
    expect(metrics.ascent).toBeCloseTo(fontSizePx * 0.8, 1)
    expect(metrics.descent).toBeCloseTo(fontSizePx * 0.2, 1)
  })

  it('CJK character width ≈ 1em', () => {
    const fontSizePx = 16
    const style = { fontFamily: 'SimSun', fontSizePx, bold: false, italic: false }
    const w = m.measure('中文', style)
    expect(w).toBeCloseTo(fontSizePx * 2, 0)
  })

  it('Latin character width ≈ 0.52em (average)', () => {
    const fontSizePx = 16
    const style = { fontFamily: 'Calibri', fontSizePx, bold: false, italic: false }
    const w = m.measure('hello', style)
    // 5 chars × 0.52em × 16px ≈ 41.6
    expect(w).toBeGreaterThan(30)
    expect(w).toBeLessThan(60)
  })

  it('bold is slightly wider than regular (×1.04)', () => {
    const fontSizePx = 16
    const normal = m.measure('ABC', { fontFamily: 'Arial', fontSizePx, bold: false, italic: false })
    const bold = m.measure('ABC', { fontFamily: 'Arial', fontSizePx, bold: true, italic: false })
    expect(bold).toBeCloseTo(normal * 1.04, 2)
  })
})

// ─── computeLineHeight ─────────────────────────────────────────────────────

describe('computeLineHeight', () => {
  const naturalH = 20 // px, assume the font's natural line height is 20px

  it('auto mode single spacing (240) = natural line height', () => {
    const h = computeLineHeight(naturalH, 'auto', 240, undefined)
    expect(h).toBeCloseTo(naturalH, 2)
  })

  it('auto mode 1.5x (360) = naturalH × 1.5', () => {
    const h = computeLineHeight(naturalH, 'auto', 360, undefined)
    expect(h).toBeCloseTo(naturalH * 1.5, 2)
  })

  it('auto mode 2x (480) = naturalH × 2', () => {
    const h = computeLineHeight(naturalH, 'auto', 480, undefined)
    expect(h).toBeCloseTo(naturalH * 2, 2)
  })

  it('atLeast mode: uses the specified value when it is larger than the natural height', () => {
    const atLeast = 400 // twips = 400/1440*96 ≈ 26.67px
    const atLeastPx = atLeast * TWIPS_TO_PX
    const h = computeLineHeight(naturalH, 'atLeast', atLeast, undefined)
    expect(h).toBeCloseTo(atLeastPx, 1)
    expect(h).toBeGreaterThan(naturalH)
  })

  it('atLeast mode: uses the natural height when the specified value is smaller', () => {
    const atLeast = 200 // twips = 200/1440*96 ≈ 13.33px < 20px
    const h = computeLineHeight(naturalH, 'atLeast', atLeast, undefined)
    expect(h).toBeCloseTo(naturalH, 2)
  })

  it('exact mode: fixed line height (independent of natural height)', () => {
    const exact = 300 // twips = 300/1440*96 = 20px
    const exactPx = exact * TWIPS_TO_PX
    const h = computeLineHeight(naturalH * 2, 'exact', exact, undefined)
    expect(h).toBeCloseTo(exactPx, 1)
  })

  it('no lineRule = default single spacing (returns natural line height)', () => {
    const h = computeLineHeight(naturalH, undefined, undefined, undefined)
    expect(h).toBeCloseTo(naturalH, 2)
  })

  it('docGrid lines mode: line height is not rounded (measured on real Word for Mac)', () => {
    // Precise Word baseline (line-by-line from the 02/10 corpus PDFs): 12pt/1.15×
    // lays out at 18pt/line, not a multiple of linePitch (15.6pt) — docGrid does not affect line height
    const docGrid = { type: 'lines' as const, linePitch: 312 }
    const h = computeLineHeight(20, 'auto', 240, docGrid)
    expect(h).toBeCloseTo(20, 2)
  })

  it('docGrid lines mode: large line heights are not rounded either', () => {
    const docGrid = { type: 'lines' as const, linePitch: 300 }
    const h = computeLineHeight(55, 'auto', 240, docGrid)
    expect(h).toBeCloseTo(55, 2)
  })

  it('type=default does no grid rounding', () => {
    const docGrid = { type: 'default' as const, linePitch: 312 }
    const h = computeLineHeight(naturalH, 'auto', 240, docGrid)
    expect(h).toBeCloseTo(naturalH, 2)
  })
})

// ─── snapSpacingToGrid ─────────────────────────────────────────────────────

describe('snapSpacingToGrid', () => {
  it('converts straight to px without a docGrid', () => {
    const px = snapSpacingToGrid(160, undefined)
    expect(px).toBeCloseTo(160 * TWIPS_TO_PX, 2)
  })

  it('type=default does not round', () => {
    const px = snapSpacingToGrid(160, { type: 'default', linePitch: 312 })
    expect(px).toBeCloseTo(160 * TWIPS_TO_PX, 2)
  })

  it('type=lines does not round either (real Word: 240tw space-after lays out at the raw 8pt)', () => {
    const px = snapSpacingToGrid(160, { type: 'lines', linePitch: 312 })
    expect(px).toBeCloseTo(160 * TWIPS_TO_PX, 2)
  })
})

// ─── simulateLines ─────────────────────────────────────────────────────────

describe('simulateLines', () => {
  const metrics = new HeuristicMetrics()
  const defaultSize = 12 // pt
  const defaultFamily = 'Calibri'

  it('empty runs return 1 line', () => {
    const lines = simulateLines([], 200, metrics, defaultSize, defaultFamily)
    expect(lines.length).toBe(1)
  })

  it('short text in a wide container stays on 1 line', () => {
    const runs = [{ text: 'Hello world', sizeHalfPoints: 24 }]
    const lines = simulateLines(runs, 500, metrics, defaultSize, defaultFamily)
    expect(lines.length).toBe(1)
  })

  it('long English paragraph wraps by word', () => {
    // Each word is ~5 chars × 0.52em × 16px ≈ 41.6px; 100px container → ~2 words/line
    const longText = 'word '.repeat(20).trim()
    const runs = [{ text: longText, sizeHalfPoints: 24 }]
    const lines = simulateLines(runs, 100, metrics, defaultSize, defaultFamily)
    expect(lines.length).toBeGreaterThan(3)
    expect(lines.length).toBeLessThan(20)
  })

  it('CJK per-character wrapping: Chinese chars × width ≈ 16px, 60px container → about 4 lines', () => {
    const cjkText = '中国政府工作报告示例文字' // 12 chars
    const runs = [{ text: cjkText, sizeHalfPoints: 24 }] // 16px per char
    const lines = simulateLines(runs, 60, metrics, defaultSize, defaultFamily)
    // 12 chars × 16px / 60px ≈ 3.2 → 4 lines
    expect(lines.length).toBeGreaterThanOrEqual(3)
    expect(lines.length).toBeLessThanOrEqual(5)
  })

  it('newline \\n forces a line break', () => {
    const runs = [{ text: 'line one\nline two\nline three', sizeHalfPoints: 24 }]
    const lines = simulateLines(runs, 500, metrics, defaultSize, defaultFamily)
    expect(lines.length).toBe(3)
  })
})

// ─── computeLineMetrics main entry ─────────────────────────────────────────

describe('computeLineMetrics', () => {
  it('empty paragraph returns 1 line with totalHeight > 0', () => {
    const result = computeLineMetrics({
      runs: [],
      availWidthPx: 400,
      isEmpty: true,
    })
    expect(result.lineCount).toBe(1)
    expect(result.totalHeight).toBeGreaterThan(0)
  })

  it('space before/after counts toward totalHeight', () => {
    const result = computeLineMetrics({
      runs: [{ text: 'test', sizeHalfPoints: 24 }],
      availWidthPx: 400,
      spaceBefore: 160, // ~10.7px
      spaceAfter: 200, // ~13.3px
    })
    const spBefore = result.spaceBeforePx
    const spAfter = result.spaceAfterPx
    expect(spBefore).toBeGreaterThan(0)
    expect(spAfter).toBeGreaterThan(0)
    expect(result.totalHeight).toBeCloseTo(
      result.lineHeights.reduce((s, h) => s + h, 0) + spBefore + spAfter,
      2,
    )
  })

  it('docGrid does not change line height: Chinese official doc linePitch=312 (A4 page, 12pt SimSun)', () => {
    // Real-Word baseline: line height is identical with or without docGrid (no snapping)
    const docGrid = { type: 'lines' as const, linePitch: 312 }
    const input = {
      runs: [{ text: '中华人民共和国', sizeHalfPoints: 24 }], // 12pt
      availWidthPx: 400,
      lineRule: 'auto' as const,
      lineRawTwips: 312,
    }
    const withGrid = computeLineMetrics({ ...input, docGrid })
    const without = computeLineMetrics(input)
    expect(withGrid.lineHeights).toEqual(without.lineHeights)
    expect(withGrid.totalHeight).toBeCloseTo(without.totalHeight, 2)
  })

  it('exact line-height mode: total height = lineCount × exactPx + spacing', () => {
    const exactTwips = 360 // 360 twips = 24px
    const exactPx = exactTwips * TWIPS_TO_PX
    const result = computeLineMetrics({
      runs: [{ text: 'one two three', sizeHalfPoints: 24 }],
      availWidthPx: 400,
      lineRule: 'exact',
      lineRawTwips: exactTwips,
    })
    for (const h of result.lineHeights) {
      expect(h).toBeCloseTo(exactPx, 1)
    }
  })

  it('atLeast line-height mode: line height >= the specified value', () => {
    const atLeastTwips = 400
    const atLeastPx = atLeastTwips * TWIPS_TO_PX
    const result = computeLineMetrics({
      runs: [{ text: 'test', sizeHalfPoints: 24 }],
      availWidthPx: 400,
      lineRule: 'atLeast',
      lineRawTwips: atLeastTwips,
    })
    for (const h of result.lineHeights) {
      expect(h).toBeGreaterThanOrEqual(atLeastPx - 0.01)
    }
  })

  it('lineCount matches the length of the line-height array', () => {
    const result = computeLineMetrics({
      runs: [{ text: '这是一段较长的中文内容，需要换行处理' }],
      availWidthPx: 100,
      lineRule: 'auto',
      lineRawTwips: 240,
    })
    expect(result.lineCount).toBe(result.lineHeights.length)
    expect(result.lineCount).toBeGreaterThan(1)
  })
})

// ─── lineTexts (locating the page-start line for in-block page splits) ─────

describe('lineTexts', () => {
  it('CJK per-character wrapping: concatenated line texts restore the original and align with lineHeights', () => {
    const text = '这是一段很长的文字内容用于测试大段落的分页处理行为'
    const result = computeLineMetrics({
      runs: [{ text, sizeHalfPoints: 24 }],
      availWidthPx: 16 * 8, // 8 CJK chars per line
    })
    expect(result.lineTexts.length).toBe(result.lineHeights.length)
    expect(result.lineTexts.join('')).toBe(text)
    expect(result.lineTexts[0]).toBe('这是一段很长的文')
    expect(result.lineTexts[1]).toBe('字内容用于测试大')
  })

  it('English wraps by word: the space at the break is consumed and words are not split', () => {
    const result = computeLineMetrics({
      runs: [{ text: 'aaaa bbbb cccc dddd', sizeHalfPoints: 24 }],
      availWidthPx: 16 * 0.52 * 9, // fits roughly one word + a space
    })
    expect(result.lineTexts.length).toBe(result.lineHeights.length)
    expect(result.lineTexts.join(' ').replace(/\s+/g, ' ')).toBe('aaaa bbbb cccc dddd')
    for (const t of result.lineTexts) expect(t.trim().length).toBeGreaterThan(0)
  })

  it('empty paragraph lineTexts is a single empty string', () => {
    const result = computeLineMetrics({ runs: [], availWidthPx: 400, isEmpty: true })
    expect(result.lineTexts).toEqual([''])
  })
})

// ─── cssLineHeight (canvas line height per the document's spacing rules) ───

describe('cssLineHeight', () => {
  it('auto multiple → calc(coefficient variable × multiple)', () => {
    expect(cssLineHeight('auto', 276, 1.15)).toBe('calc(var(--doc-line-factor,1.2) * 1.15)')
  })

  it('derives the multiple from auto twips when lineSpacing is absent', () => {
    expect(cssLineHeight('auto', 360, undefined)).toBe('calc(var(--doc-line-factor,1.2) * 1.5)')
  })

  it('exact → fixed pt', () => {
    expect(cssLineHeight('exact', 320, undefined)).toBe('16.0pt')
  })

  it('atLeast → max(pt, natural line height)', () => {
    expect(cssLineHeight('atLeast', 360, undefined)).toBe(
      'max(18.0pt, calc(var(--doc-line-factor,1.2) * 1em))',
    )
  })

  it('no line spacing settings at all → null (inherit)', () => {
    expect(cssLineHeight(undefined, undefined, undefined)).toBeNull()
  })
})

describe('cssFontFamily', () => {
  it('common Word fonts → metric-compatible fallback + CJK safety net', () => {
    expect(cssFontFamily('Calibri')).toBe("'Calibri','Carlito','Noto Sans CJK SC',sans-serif")
    expect(cssFontFamily('Times New Roman')).toBe(
      "'Times New Roman','Liberation Serif','Noto Serif CJK SC',serif",
    )
    expect(cssFontFamily('宋体')).toBe(
      "'宋体','Songti SC','STSong','SimSun','Noto Serif CJK SC',serif",
    )
  })

  it('Simplified-Chinese office fonts map to real macOS/Windows families', () => {
    expect(cssFontFamily('仿宋_GB2312')).toBe(
      "'仿宋_GB2312','STFangsong','FangSong','Noto Serif CJK SC',serif",
    )
    expect(cssFontFamily('楷体_GB2312')).toBe(
      "'楷体_GB2312','STKaiti','Kaiti SC','KaiTi','Noto Serif CJK SC',serif",
    )
    expect(cssFontFamily('黑体')).toBe(
      "'黑体','Heiti SC','STHeiti','SimHei','PingFang SC','Noto Sans CJK SC',sans-serif",
    )
    expect(cssFontFamily('方正小标宋_GBK')).toBe(
      "'方正小标宋_GBK','STZhongsong','Songti SC','STSong','SimSun','Noto Serif CJK SC',serif",
    )
    expect(cssFontFamily('方正小标宋简体')).toContain("'STZhongsong'")
    expect(cssFontFamily('华文中宋')).toBe(
      "'华文中宋','STZhongsong','Songti SC','STSong','SimSun','Noto Serif CJK SC',serif",
    )
    expect(cssFontFamily('等线')).toBe(
      "'等线','DengXian','PingFang SC','Microsoft YaHei','Noto Sans CJK SC',sans-serif",
    )
  })

  it('unknown font family guesses serif-ness by name', () => {
    expect(cssFontFamily('SomeCustomFont')).toBe("'SomeCustomFont','Noto Sans CJK SC',sans-serif")
  })

  it('Japanese fonts → same-script fallback chain, never falls back to Simplified Chinese', () => {
    expect(cssFontFamily('游ゴシック')).toBe(
      "'游ゴシック','Yu Gothic','Hiragino Sans','Meiryo','Noto Sans JP',sans-serif",
    )
    expect(cssFontFamily('ＭＳ Ｐ明朝')).toBe(
      "'ＭＳ Ｐ明朝','Yu Mincho','Hiragino Mincho ProN','MS Mincho','Noto Serif JP',serif",
    )
    expect(cssFontFamily('Meiryo')).toContain("'Hiragino Sans'")
    expect(cssFontFamily('Meiryo')).not.toContain('CJK SC')
  })

  it('Korean/Traditional Chinese fonts → same-script fallback chain', () => {
    expect(cssFontFamily('맑은 고딕')).toBe(
      "'맑은 고딕','Malgun Gothic','Apple SD Gothic Neo','Noto Sans KR',sans-serif",
    )
    expect(cssFontFamily('Batang')).toBe("'Batang','AppleMyungjo','Noto Serif KR',serif")
    expect(cssFontFamily('微軟正黑體')).toBe(
      "'微軟正黑體','Microsoft JhengHei','PingFang TC','Heiti TC','Noto Sans TC',sans-serif",
    )
    expect(cssFontFamily('新細明體')).toBe(
      "'新細明體','PMingLiU','MingLiU','Songti TC','Noto Serif TC',serif",
    )
  })
})

describe('textHasCjk', () => {
  it('detects CJK vs pure Western text', () => {
    expect(textHasCjk('中文 abc')).toBe(true)
    expect(textHasCjk('English only, 123.')).toBe(false)
    expect(textHasCjk('')).toBe(false)
  })
})
