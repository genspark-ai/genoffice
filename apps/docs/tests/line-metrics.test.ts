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

import { afterEach, describe, it, expect, vi } from 'vitest'
import {
  HeuristicMetrics,
  autospaceBoundaries,
  autospacePadBetween,
  computeLineHeight,
  cssCsFontFamily,
  cssDualFontFamily,
  cssFontFamily,
  cssLineHeight,
  krLineFactor,
  lineHeightFactor,
  paraLineFactorCss,
  snapSpacingToGrid,
  simulateLines,
  computeLineMetrics,
  textHasCjk,
  textHasComplexScript,
  textHasHangul,
} from '../src/renderer/line-metrics'

const TWIPS_TO_PX = 96 / 1440

// ─── HeuristicMetrics ─────────────────────────────────────────────────────

describe('HeuristicMetrics', () => {
  const m = new HeuristicMetrics()

  it('natural line height of 12pt Arial follows its hhea factor (1.15em)', () => {
    const fontSizePx = 12 * (96 / 72) // 16px
    const style = { fontFamily: 'Arial', fontSizePx, bold: false, italic: false }
    const metrics = m.metrics(style)
    expect(metrics.lineHeight).toBeCloseTo(fontSizePx * 1.15, 1)
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

  it('docGrid lines mode: single/auto lines snap up to whole linePitch cells (LO probe)', () => {
    // LO baseline (probe docx, 2026-08-10): natural 20px on a 15.6pt (20.8px)
    // grid takes one cell; the auto multiple applies AFTER snapping
    const docGrid = { type: 'lines' as const, linePitch: 312 }
    const h = computeLineHeight(20, 'auto', 240, docGrid)
    expect(h).toBeCloseTo(312 * (96 / 1440), 2)
  })

  it('docGrid lines mode: taller lines take more whole cells', () => {
    const docGrid = { type: 'lines' as const, linePitch: 300 }
    const h = computeLineHeight(55, 'auto', 240, docGrid)
    // 300tw = 20px cell; 55px needs 3 cells = 60px
    expect(h).toBeCloseTo(60, 2)
  })

  it('docGrid lines mode: exact and atLeast never snap', () => {
    const docGrid = { type: 'lines' as const, linePitch: 312 }
    expect(computeLineHeight(20, 'exact', 260, docGrid)).toBeCloseTo(260 * (96 / 1440), 2)
    expect(computeLineHeight(20, 'atLeast', 0, docGrid)).toBeCloseTo(20, 2)
    expect(computeLineHeight(20, 'atLeast', 480, docGrid)).toBeCloseTo(480 * (96 / 1440), 2)
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

  it('type=lines quantizes spacing DOWN to whole cells (LO probe: 6pt vanishes on a 15.6pt grid)', () => {
    expect(snapSpacingToGrid(160, { type: 'lines', linePitch: 312 })).toBeCloseTo(0, 2)
    expect(snapSpacingToGrid(480, { type: 'lines', linePitch: 312 })).toBeCloseTo(
      312 * TWIPS_TO_PX,
      2,
    )
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

  it('docGrid snaps line heights: Chinese official doc linePitch=312 (A4 page, 12pt SimSun)', () => {
    // LO baseline: 12pt SimSun natural (1.7em = 27.2px) exceeds one 20.8px cell,
    // so each line takes two cells before the 1.3 auto multiple applies
    const docGrid = { type: 'lines' as const, linePitch: 312 }
    const input = {
      runs: [{ text: '中华人民共和国', sizeHalfPoints: 24 }], // 12pt
      availWidthPx: 400,
      lineRule: 'auto' as const,
      lineRawTwips: 312,
    }
    const withGrid = computeLineMetrics({ ...input, docGrid })
    const without = computeLineMetrics(input)
    const cell = 312 * (96 / 1440)
    for (const [i, h] of withGrid.lineHeights.entries()) {
      expect(h).toBeCloseTo(
        Math.ceil(without.lineHeights[i] / (312 / 240) / cell - 0.001) * cell * (312 / 240),
        2,
      )
    }
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
    expect(cssLineHeight('auto', 276, 1.15)).toBe(
      'calc(round(up, calc(var(--doc-line-factor,1.2) * 1em), var(--doc-grid-pitch,0.0001px)) * 1.15)',
    )
  })

  it('derives the multiple from auto twips when lineSpacing is absent', () => {
    expect(cssLineHeight('auto', 360, undefined)).toBe(
      'calc(round(up, calc(var(--doc-line-factor,1.2) * 1em), var(--doc-grid-pitch,0.0001px)) * 1.5)',
    )
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
    expect(cssFontFamily('Calibri')).toBe("'Calibri','Carlito GO','Noto Sans CJK SC',sans-serif")
    expect(cssFontFamily('Times New Roman')).toBe(
      "'Times New Roman','Liberation Serif','Noto Serif CJK SC',serif",
    )
    expect(cssFontFamily('宋体')).toBe(
      "'宋体','GenOffice Songti SC','STSong','SimSun','Noto Serif CJK SC',serif",
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

  describe('SC-variant declares (Word substitutes missing East Asian fonts with a serif)', () => {
    function stubCanvas(availableFamilies: string[]) {
      const known = new Set(availableFamilies)
      let width = 50
      const fake = {
        set font(spec: string) {
          const family = /"([^"]+)"/.exec(spec)?.[1]
          width = family !== undefined && known.has(family) ? 100 : 50
        },
        measureText: () => ({ width }),
      }
      vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(
        fake as unknown as CanvasRenderingContext2D,
      )
    }

    afterEach(() => vi.restoreAllMocks())

    it('missing SC sans routes to the SimSun-class serif chain', () => {
      expect(cssFontFamily('Noto Sans SC')).toBe(
        "'Noto Sans SC','GenOffice Songti SC','STSong','SimSun','Noto Serif CJK SC',serif",
      )
    })

    it('bundled subset faces count as missing and never lead the chain', () => {
      expect(cssFontFamily('Noto Sans CJK SC')).toBe(
        "'GenOffice Songti SC','STSong','SimSun','Noto Serif CJK SC',serif",
      )
      expect(cssFontFamily('Noto Serif CJK SC')).toBe(
        "'GenOffice Songti SC','STSong','SimSun','Noto Serif CJK SC',serif",
      )
    })

    it('true SC serif declares keep their name at the head', () => {
      expect(cssFontFamily('Noto Serif SC')).toBe(
        "'Noto Serif SC','GenOffice Songti SC','STSong','SimSun','Noto Serif CJK SC',serif",
      )
    })

    it('locally installed SC sans keeps the declared name and a sans chain', () => {
      stubCanvas(['Source Han Sans CN'])
      expect(cssFontFamily('Source Han Sans CN')).toBe(
        "'Source Han Sans CN','PingFang SC','Microsoft YaHei','Noto Sans CJK SC',sans-serif",
      )
    })

    it('jp/kr/tc variants keep their same-script substitution', () => {
      expect(cssFontFamily('Noto Sans CJK JP')).toBe(
        "'Noto Sans CJK JP','Yu Mincho','Hiragino Mincho ProN','MS Mincho','Noto Serif JP',serif",
      )
      expect(cssFontFamily('Source Han Sans K')).toBe(
        "'Source Han Sans K','Batang','GenOffice Serif KR','AppleMyungjo','Noto Serif KR',serif",
      )
      expect(cssFontFamily('Noto Sans CJK TC')).toBe(
        "'Noto Sans CJK TC','PMingLiU','MingLiU','GenOffice Fullwidth TC','Songti TC','Noto Serif TC',serif",
      )
    })
  })

  it('Korean/Traditional Chinese fonts → same-script fallback chain', () => {
    expect(cssFontFamily('맑은 고딕')).toBe(
      "'맑은 고딕','Malgun Gothic','GenOffice Sans KR','Apple SD Gothic Neo','Noto Sans KR',sans-serif",
    )
    expect(cssFontFamily('Batang')).toBe(
      "'Batang','GenOffice Serif KR','AppleMyungjo','Noto Serif KR',serif",
    )
    expect(cssFontFamily('微軟正黑體')).toBe(
      "'微軟正黑體','Microsoft JhengHei','PingFang TC','Heiti TC','Noto Sans TC',sans-serif",
    )
    expect(cssFontFamily('新細明體')).toBe(
      "'新細明體','PMingLiU','MingLiU','GenOffice Fullwidth TC','Songti TC','Noto Serif TC',serif",
    )
  })
})

describe('textHasCjk', () => {
  it('detects CJK vs pure Western text', () => {
    expect(textHasCjk('中文 abc')).toBe(true)
    expect(textHasCjk('English only, 123.')).toBe(false)
    expect(textHasCjk('')).toBe(false)
  })

  it('hangul counts as CJK (syllables + jamo)', () => {
    expect(textHasCjk('한국어 문서')).toBe(true)
    expect(textHasCjk('가')).toBe(true)
    expect(textHasCjk('ㄱㄴ')).toBe(true)
  })
})

// ─── Korean fidelity ────────────────────────────────────────────────────────

describe('Korean line metrics', () => {
  it('hangul advances 1.0em in the heuristic model', () => {
    const m = new HeuristicMetrics()
    const style = { fontFamily: 'Batang', fontSizePx: 16, bold: false, italic: false }
    expect(m.measure('한', style)).toBeCloseTo(16, 5)
    expect(m.measure('한글날', style)).toBeCloseTo(48, 5)
  })

  it('Korean font line factors: Batang-class 1.4583, Malgun 1.775 (LO probe)', () => {
    expect(lineHeightFactor('Batang')).toBe(1.4583)
    expect(lineHeightFactor('바탕')).toBe(1.4583)
    expect(lineHeightFactor('Gulim')).toBe(1.4583)
    expect(lineHeightFactor('Dotum')).toBe(1.4583)
    expect(lineHeightFactor('NanumMyeongjo')).toBe(1.4583)
    expect(lineHeightFactor('Malgun Gothic')).toBe(1.775)
    expect(lineHeightFactor('맑은 고딕')).toBe(1.775)
    expect(lineHeightFactor('Noto Sans CJK KR')).toBe(1.4583)
    expect(lineHeightFactor('Noto Serif KR')).toBe(1.4583)
    expect(lineHeightFactor('Source Han Sans K')).toBe(1.4583)
  })

  it('Chinese/Japanese factors follow the LO substitution probe', () => {
    expect(lineHeightFactor('SimSun')).toBe(1.7)
    expect(lineHeightFactor('宋体')).toBe(1.7)
    expect(lineHeightFactor('DengXian')).toBe(1.775)
    expect(lineHeightFactor('等线')).toBe(1.775)
    expect(lineHeightFactor('仿宋_GB2312')).toBe(1.775)
    expect(lineHeightFactor('黑体')).toBe(1.0)
    expect(lineHeightFactor('楷体')).toBe(1.0)
    expect(lineHeightFactor('楷体_GB2312')).toBe(1.775)
    expect(lineHeightFactor('ＭＳ 明朝')).toBe(1.7)
    expect(lineHeightFactor('MS Gothic')).toBe(1.7)
    expect(lineHeightFactor('游明朝')).toBe(2.2667)
    expect(lineHeightFactor('Meiryo')).toBe(1.775)
    expect(lineHeightFactor('PMingLiU')).toBe(1.0)
    expect(lineHeightFactor('Microsoft JhengHei')).toBe(1.775)
    expect(lineHeightFactor('Calibri')).toBe(1.22)
    expect(lineHeightFactor('Century Gothic')).toBe(1.2)
  })

  it('SC-variant declares take the PingFang-class factor (their substitution target when missing)', () => {
    expect(lineHeightFactor('Noto Sans CJK SC')).toBe(1.775)
    expect(lineHeightFactor('Noto Serif SC')).toBe(1.775)
    expect(lineHeightFactor('Source Han Sans CN')).toBe(1.775)
    expect(lineHeightFactor('Noto Sans SC')).toBe(1.8375)
  })

  it('textHasHangul separates Korean from other CJK', () => {
    expect(textHasHangul('보고서 2026')).toBe(true)
    expect(textHasHangul('\u4e2d\u6587')).toBe(false)
    expect(textHasHangul('かな')).toBe(false)
  })

  it('paraLineFactorCss routes by script', () => {
    expect(paraLineFactorCss('한국어')).toBe('var(--doc-line-factor-kr,1.4583)')
    expect(paraLineFactorCss('\u4e2d\u6587')).toBe('var(--doc-line-factor-cjk,1.7)')
    expect(paraLineFactorCss('latin')).toBe('var(--doc-line-factor-latin,1.2)')
  })

  it('krLineFactor follows the EA face, defaulting to Batang-class', () => {
    expect(krLineFactor('Batang')).toBe(1.4583)
    expect(krLineFactor('맑은 고딕')).toBe(1.775)
    expect(krLineFactor(undefined)).toBe(1.4583)
  })

  it('Korean ascii face in a dual-slot chain keeps only the literal family', () => {
    expect(cssDualFontFamily('맑은 고딕', 'Batang')).toBe(
      "'맑은 고딕','Batang','GenOffice Serif KR','AppleMyungjo','Noto Serif KR',serif",
    )
  })

  it('hangul wraps at word boundaries like Word, not per syllable', () => {
    const m = new HeuristicMetrics()
    const lines = simulateLines(
      [{ text: '가나다 라마바 사아자', sizeHalfPoints: 24 }],
      16 * 6 + 3, // fits 6 syllables (96px) but not word+space+word (101px)
      m,
      12,
      'Batang',
    )
    expect(lines.map((ln) => ln.text.trim())).toEqual(['가나다', '라마바', '사아자'])
  })

  it('an overlong hangul word still hard-breaks inside the word', () => {
    const m = new HeuristicMetrics()
    const lines = simulateLines(
      [{ text: '한글한글한글', sizeHalfPoints: 24 }],
      16 * 4 + 1, // 4 syllables per line at 12pt (16px)
      m,
      12,
      'Batang',
    )
    expect(lines.map((ln) => ln.text)).toEqual(['한글한글', '한글'])
  })

  it('hangul words keep the CJK line-height floor of the per-syllable model', () => {
    const m = new HeuristicMetrics()
    const lines = simulateLines([{ text: '한글 문서', sizeHalfPoints: 24 }], 500, m, 12, 'Calibri')
    // Latin-font run: CJK fallback factor 1.3 beats Calibri's 1.22
    expect(lines[0].naturalLineH).toBeCloseTo(16 * 1.3, 5)
  })

  it('Chinese keeps per-character wrapping even with spaces present', () => {
    const m = new HeuristicMetrics()
    const lines = simulateLines(
      [{ text: '中中中 中中中中中', sizeHalfPoints: 24 }],
      16 * 6 + 3,
      m,
      12,
      'SimSun',
    )
    // char-level fill: the second word splits across the line boundary
    expect(lines.length).toBe(2)
    expect(lines[0].text).toBe('中中中 中中')
  })

  it('Korean paragraphs pull mixed-in Han into the word buffer (keep-all)', () => {
    const m = new HeuristicMetrics()
    const lines = simulateLines(
      [{ text: '한글漢字한글 다음', sizeHalfPoints: 24 }],
      16 * 6 + 3, // fits 6 chars; the 7th ('다') would split a per-char line
      m,
      12,
      'Batang',
    )
    expect(lines.map((ln) => ln.text.trim())).toEqual(['한글漢字한글', '다음'])
  })
})

describe('autospaceBoundaries', () => {
  it('finds kana-Latin and kana-digit boundaries', () => {
    expect(autospaceBoundaries('ペン12')).toEqual([2])
    expect(autospaceBoundaries('12ペン')).toEqual([2])
    expect(autospaceBoundaries('テスト17.0km')).toEqual([3])
  })

  it('covers Han and hangul on the CJK side', () => {
    expect(autospaceBoundaries('A漢B')).toEqual([1, 2])
    expect(autospaceBoundaries('한글A')).toEqual([2])
  })

  it('needs direct adjacency: spaces and punctuation get no pad', () => {
    expect(autospaceBoundaries('ペン 12')).toEqual([])
    expect(autospaceBoundaries('ペン、12')).toEqual([])
    expect(autospaceBoundaries('。A')).toEqual([])
    expect(autospaceBoundaries('あ・A')).toEqual([])
  })

  it('ignores fullwidth/halfwidth forms and non-CJK astral chars', () => {
    expect(autospaceBoundaries('Ａ' + '1')).toEqual([])
    expect(autospaceBoundaries('ｱA')).toEqual([])
    expect(autospaceBoundaries('あ\u{1F600}A')).toEqual([])
  })
})

describe('autospacePadBetween', () => {
  it('pads only when the seam chars are directly adjacent CJK and Latin', () => {
    expect(autospacePadBetween('ペン', '12')).toBe(true)
    expect(autospacePadBetween('12', 'ペン')).toBe(true)
    expect(autospacePadBetween('ペン ', '12')).toBe(false)
    expect(autospacePadBetween('ペン', ' 12')).toBe(false)
    expect(autospacePadBetween('', '12')).toBe(false)
    expect(autospacePadBetween('ペン', '')).toBe(false)
  })
})

// ─── Arabic fidelity ────────────────────────────────────────────────────────

describe('cssFontFamily Arabic', () => {
  it('naskh/serif-class names get the bundled Naskh chain', () => {
    expect(cssFontFamily('Noto Naskh Arabic')).toBe(
      "'Noto Naskh Arabic','Geeza Pro','Al Bayan',serif",
    )
    expect(cssFontFamily('Arabic Typesetting')).toBe(
      "'Arabic Typesetting','Noto Naskh Arabic','Geeza Pro','Al Bayan',serif",
    )
    expect(cssFontFamily('Amiri')).toContain("'Noto Naskh Arabic'")
    expect(cssFontFamily('Scheherazade New')).toContain("'Noto Naskh Arabic'")
  })

  it('Traditional/Simplified Arabic map to the size-adjusted alias', () => {
    expect(cssFontFamily('Traditional Arabic')).toBe(
      "'Traditional Arabic','Noto Naskh Arabic TA','Geeza Pro','Al Bayan',serif",
    )
    expect(cssFontFamily('Simplified Arabic')).toContain("'Noto Naskh Arabic TA'")
    // other naskh-class names keep the unscaled subset
    expect(cssFontFamily('Amiri')).not.toContain("'Noto Naskh Arabic TA'")
    expect(cssFontFamily('Arabic Typesetting')).not.toContain("'Noto Naskh Arabic TA'")
  })

  it('kufi/sans-class names get the Sans Arabic chain', () => {
    expect(cssFontFamily('Noto Sans Arabic')).toBe("'Noto Sans Arabic','Geeza Pro',sans-serif")
    expect(cssFontFamily('Noto Kufi Arabic')).toBe(
      "'Noto Kufi Arabic','Noto Sans Arabic','Geeza Pro',sans-serif",
    )
  })

  it('unknown Arabic names (by script in the name) default to the naskh chain', () => {
    expect(cssFontFamily('الخط الديواني')).toBe(
      "'الخط الديواني','Noto Naskh Arabic','Geeza Pro','Al Bayan',serif",
    )
    expect(cssFontFamily('Urdu Typesetting')).toContain("'Noto Naskh Arabic'")
  })

  it('does not capture CJK or Latin families', () => {
    expect(cssFontFamily('SimSun')).not.toContain('Arabic')
    expect(cssFontFamily('Calibri')).not.toContain('Arabic')
    expect(cssFontFamily('Batang')).not.toContain('Arabic')
    expect(cssFontFamily('SomeCustomFont')).not.toContain('Arabic')
  })
})

describe('textHasComplexScript', () => {
  it('detects Arabic, Hebrew and presentation forms', () => {
    expect(textHasComplexScript('مرحبا')).toBe(true)
    expect(textHasComplexScript('שלום')).toBe(true)
    expect(textHasComplexScript('ﻻ')).toBe(true)
  })

  it('is false for Latin and CJK', () => {
    expect(textHasComplexScript('hello 123')).toBe(false)
    expect(textHasComplexScript('\u4e2d\u6587')).toBe(false)
    expect(textHasComplexScript('かな한글')).toBe(false)
  })
})

describe('cssCsFontFamily', () => {
  it('cs chain leads; base Latin chain slots in after the bundled Naskh subset (no Latin coverage) and before Geeza Pro', () => {
    expect(cssCsFontFamily('Arabic Typesetting', 'Calibri', 'Calibri')).toBe(
      "'Arabic Typesetting','Noto Naskh Arabic','Calibri','Carlito GO','Noto Sans CJK SC','Geeza Pro','Al Bayan',sans-serif",
    )
  })

  it('Traditional Arabic cs run keeps the size-adjusted alias ahead of the Latin chain', () => {
    const chain = cssCsFontFamily('Traditional Arabic', 'Times New Roman')
    expect(chain.startsWith("'Traditional Arabic','Noto Naskh Arabic TA','Times New Roman'")).toBe(
      true,
    )
    expect(chain.indexOf("'Times New Roman'")).toBeLessThan(chain.indexOf("'Geeza Pro'"))
  })

  it('keeps a dual-slot base after the cs chain', () => {
    const chain = cssCsFontFamily('Amiri', 'Times New Roman', 'SimSun')
    expect(chain.startsWith("'Amiri','Noto Naskh Arabic'")).toBe(true)
    expect(chain).toContain("'Times New Roman'")
    expect(chain).toContain("'SimSun'")
    expect(chain.indexOf("'Times New Roman'")).toBeLessThan(chain.indexOf("'Geeza Pro'"))
  })

  it('non-Arabic cs font keeps head-then-base order', () => {
    const chain = cssCsFontFamily('David', 'Calibri', 'Calibri')
    expect(chain.startsWith("'David'")).toBe(true)
    expect(chain).toContain("'Calibri'")
  })

  it('cs-only run falls back to the plain cs chain', () => {
    expect(cssCsFontFamily('Noto Naskh Arabic')).toBe(
      "'Noto Naskh Arabic','Geeza Pro','Al Bayan',serif",
    )
  })

  it('deduplicates shared families', () => {
    const chain = cssCsFontFamily('Noto Naskh Arabic', 'Noto Naskh Arabic')
    expect(chain.match(/'Noto Naskh Arabic'/g)?.length).toBe(1)
  })
})
