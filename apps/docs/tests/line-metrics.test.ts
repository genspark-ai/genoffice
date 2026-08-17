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
  cssGridLineExpr,
  snapLineToPitch,
  estimateFootnoteHeight,
  footnoteLineHeightPx,
  cssCsFontFamily,
  cssDualFontFamily,
  cssFontFamily,
  cssLineHeight,
  krLineFactor,
  cjkDeclaredLineFactor,
  lineHeightFactor,
  simsunGapLineFactor,
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

  it('monospace Latin chars share one constant advance; proportional do not', () => {
    const fontSizePx = 16
    const mono = { fontFamily: 'Consolas', fontSizePx, bold: false, italic: false }
    expect(m.measure('iiii', mono)).toBeCloseTo(m.measure('mmmm', mono), 5)
    expect(m.measure('iiii', mono)).toBeCloseTo(4 * 0.55 * fontSizePx, 5)
    const sans = { fontFamily: 'Calibri', fontSizePx, bold: false, italic: false }
    expect(m.measure('iiii', sans)).toBeLessThan(m.measure('mmmm', sans))
  })

  it('monospace advance is judged by chain head; CJK stays 1em; Courier tier is 0.6em', () => {
    const fontSizePx = 16
    const chain = {
      fontFamily: "'Consolas','Menlo','Courier New','Liberation Mono',monospace",
      fontSizePx,
      bold: false,
      italic: false,
    }
    expect(m.measure('im', chain)).toBeCloseTo(2 * 0.55 * fontSizePx, 5)
    expect(m.measure('中', chain)).toBeCloseTo(fontSizePx, 5)
    const courier = { fontFamily: 'Courier New', fontSizePx, bold: false, italic: false }
    expect(m.measure('im', courier)).toBeCloseTo(2 * 0.6 * fontSizePx, 5)
  })

  it('bold does not widen monospace text', () => {
    const fontSizePx = 16
    const normal = m.measure('ab', { fontFamily: 'Menlo', fontSizePx, bold: false, italic: false })
    const bold = m.measure('ab', { fontFamily: 'Menlo', fontSizePx, bold: true, italic: false })
    expect(bold).toBe(normal)
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

  // Word probe (sample 32, 18pt grid, Yu Mincho 1.44): 10.5pt lines take one
  // cell, 13pt lines flip to two — the calibrated flip threshold.
  it('18pt grid: 10.5pt/1.44 snaps to one 18pt cell, 13pt/1.44 to two (Word probe)', () => {
    const docGrid = { type: 'lines' as const, linePitch: 360 } // 18pt = 24px
    const pt = (v: number) => (v * 96) / 72
    expect(computeLineHeight(1.44 * pt(10.5), undefined, undefined, docGrid)).toBeCloseTo(pt(18), 2)
    expect(computeLineHeight(1.44 * pt(13), undefined, undefined, docGrid)).toBeCloseTo(pt(36), 2)
    // no grid: pure factor, unsnapped
    expect(computeLineHeight(1.44 * pt(10.5), undefined, undefined, undefined)).toBeCloseTo(
      1.44 * pt(10.5),
      2,
    )
  })
})

// ─── snapLineToPitch ────────────────────────────────────────────────────────

describe('snapLineToPitch', () => {
  it('ceils to whole cells; zero pitch is identity', () => {
    expect(snapLineToPitch(20.16, 24)).toBe(24)
    expect(snapLineToPitch(24.96, 24)).toBe(48)
    expect(snapLineToPitch(20.16, 0)).toBe(20.16)
  })

  it('ε keeps float noise just past a cell boundary in the lower cell', () => {
    expect(snapLineToPitch(24.02, 24)).toBe(24)
    expect(snapLineToPitch(24.03, 24)).toBe(48)
  })

  it('cssGridLineExpr mirrors the same formula and ε', () => {
    expect(cssGridLineExpr()).toBe(
      'round(up, calc(var(--doc-line-factor,1.2) * 1em - var(--doc-grid-pitch,0.0001px) * 0.001), var(--doc-grid-pitch,0.0001px))',
    )
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

  it('type=lines keeps spacing at face value (Word probe 2026-08-13)', () => {
    expect(snapSpacingToGrid(160, { type: 'lines', linePitch: 312 })).toBeCloseTo(
      160 * TWIPS_TO_PX,
      2,
    )
    expect(snapSpacingToGrid(480, { type: 'lines', linePitch: 312 })).toBeCloseTo(
      480 * TWIPS_TO_PX,
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

// ─── SimSun-substitution ・/〜 line lift (Word probe 2026-08-13) ─────────────

describe('SimSun-substitution ・/〜 line lift', () => {
  const metrics = new HeuristicMetrics()
  const sizePt = 10.5
  const sizePx = sizePt * (96 / 72)
  const jpRuns = (text: string) => [
    { text, fontFamily: 'Noto Sans JP', sizeHalfPoints: sizePt * 2 },
  ]

  it('simsunGapLineFactor hits JP/SC/TC substitution and SimSun, not KR/YaHei', () => {
    expect(simsunGapLineFactor('Noto Sans JP')).toBe(1.7143)
    expect(simsunGapLineFactor('Noto Sans CJK SC')).toBe(1.7143)
    expect(simsunGapLineFactor('SimSun')).toBe(1.7143)
    expect(simsunGapLineFactor('Batang')).toBeNull()
    expect(simsunGapLineFactor('Noto Sans KR')).toBeNull()
    expect(simsunGapLineFactor('Microsoft YaHei')).toBeNull()
  })

  it('a line with ・ lifts to 1.7143 × size; a plain line stays at 1.3029', () => {
    const [plain] = simulateLines(jpRuns('日本語の本文'), 1000, metrics, sizePt, 'Noto Sans JP')
    const [lifted] = simulateLines(jpRuns('日本・語本文'), 1000, metrics, sizePt, 'Noto Sans JP')
    expect(plain.naturalLineH).toBeCloseTo(sizePx * 1.3029, 1)
    expect(lifted.naturalLineH).toBeCloseTo(sizePx * 1.7143, 1)
  })

  it('〜 lifts too; a YaHei-declared run does not change', () => {
    const [lifted] = simulateLines(jpRuns('期間は9時〜17時'), 1000, metrics, sizePt, 'Noto Sans JP')
    expect(lifted.naturalLineH).toBeCloseTo(sizePx * 1.7143, 1)
    const yahei = [{ text: '項目・内容', fontFamily: 'Microsoft YaHei', sizeHalfPoints: 21 }]
    const [unchanged] = simulateLines(yahei, 1000, metrics, sizePt, 'Microsoft YaHei')
    expect(unchanged.naturalLineH).toBeCloseTo(sizePx * 1.7143, 1) // YaHei's own factor, no extra lift
  })

  it('fixed-factor table-cell path keeps its pinned factor', () => {
    const [line] = simulateLines(jpRuns('項目・内容'), 1000, metrics, sizePt, 'Noto Sans JP', 1.0)
    expect(line.naturalLineH).toBeCloseTo(sizePx * 1.0, 1)
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

  it('empty paragraph line height follows its mark run size/font (Word rule)', () => {
    const base = { availWidthPx: 400, isEmpty: true, defaultFontSizePt: 10.5 }
    const def = computeLineMetrics({ ...base, runs: [] })
    expect(def.totalHeight).toBeCloseTo(10.5 * (96 / 72) * 1.22, 3)
    const sized = computeLineMetrics({ ...base, runs: [{ text: '', sizeHalfPoints: 28 }] })
    expect(sized.totalHeight).toBeCloseTo(14 * (96 / 72) * 1.22, 3)
    const kr = computeLineMetrics({ ...base, runs: [{ text: ' ', fontFamily: 'Malgun Gothic' }] })
    expect(kr.totalHeight).toBeCloseTo(10.5 * (96 / 72) * 1.7371, 3)
    // unstyled whitespace runs keep the document default
    const plain = computeLineMetrics({ ...base, runs: [{ text: ' ' }] })
    expect(plain.totalHeight).toBeCloseTo(def.totalHeight, 3)
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
      'calc(var(--doc-line-grid,var(--doc-line-factor,1.2)) * 1.15)',
    )
  })

  // unitless numbers inherit by value: a 36pt run in an 11pt paragraph keeps its
  // own line box instead of the paragraph's absolute height (regression sample 33)
  it('auto multiple is unitless outside typed-grid docs (no em/pt to inherit)', () => {
    const lh = cssLineHeight('auto', 276, 1.15)!
    expect(lh).not.toContain('em')
    expect(lh).not.toContain('pt')
  })

  it('derives the multiple from auto twips when lineSpacing is absent', () => {
    expect(cssLineHeight('auto', 360, undefined)).toBe(
      'calc(var(--doc-line-grid,var(--doc-line-factor,1.2)) * 1.5)',
    )
  })

  it('single spacing resolves the grid var ahead of the unitless factor', () => {
    expect(cssLineHeight('auto', 240, undefined)).toBe(
      'var(--doc-line-grid,var(--doc-line-factor,1.2))',
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

/** canvas stub for isFontAvailable: listed families measure differently from the generic fallbacks */
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

describe('cssFontFamily', () => {
  it('common Word fonts → metric-compatible fallback + CJK safety net', () => {
    expect(cssFontFamily('Calibri')).toBe("'Calibri','Carlito GO','Noto Sans CJK SC',sans-serif")
    expect(cssFontFamily('Times New Roman')).toBe(
      "'Times New Roman','Liberation Serif','GenOffice Box Drawing','Noto Serif CJK SC',serif",
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

  it('monospace families map to a mono chain (Menlo first on macOS)', () => {
    expect(cssFontFamily('Consolas')).toBe(
      "'Consolas','Menlo','Courier New','Liberation Mono','Noto Sans CJK SC',monospace",
    )
    expect(cssFontFamily('Courier New')).toBe(
      "'Courier New','Menlo','Liberation Mono','Noto Sans CJK SC',monospace",
    )
    for (const f of [
      'Menlo',
      'Monaco',
      'Cascadia Code',
      'SF Mono',
      'JetBrains Mono',
      'Source Code Pro',
      'Fira Code',
      'DejaVu Sans Mono',
      'Lucida Console',
    ]) {
      expect(cssFontFamily(f)).toMatch(/,monospace$/)
    }
  })

  it('mono css chain resolves the Consolas line factor, not Courier fallback', () => {
    expect(lineHeightFactor("'Consolas','Menlo','Courier New','Liberation Mono'")).toBe(1.1667)
  })

  it('Japanese fonts → same-script fallback chain, never falls back to Simplified Chinese', () => {
    expect(cssFontFamily('游ゴシック')).toBe(
      "'游ゴシック','Yu Gothic','GenOffice Hiragino Sans','Meiryo','Noto Sans JP',sans-serif",
    )
    expect(cssFontFamily('ＭＳ Ｐ明朝')).toBe(
      "'ＭＳ Ｐ明朝','Yu Mincho','GenOffice Hiragino Mincho','GenOffice MS Mincho','Noto Serif JP',serif",
    )
    expect(cssFontFamily('Meiryo')).toContain("'GenOffice Hiragino Sans'")
    expect(cssFontFamily('Meiryo')).not.toContain('CJK SC')
  })

  describe('SC-variant declares (Word substitutes missing East Asian fonts with a serif)', () => {
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
        "'Noto Sans CJK JP','Yu Mincho','GenOffice Hiragino Mincho','GenOffice MS Mincho','Noto Serif JP',serif",
      )
      expect(cssFontFamily('Source Han Sans K')).toBe(
        "'Source Han Sans K','KR Theme Latin GO','GenOffice Batang','GenOffice Serif KR','GenOffice Myungjo','Noto Serif KR',serif",
      )
      expect(cssFontFamily('Noto Sans CJK TC')).toBe(
        "'Noto Sans CJK TC','GenOffice MingLiU','GenOffice Fullwidth TC','Songti TC','Noto Serif TC',serif",
      )
    })
  })

  it('Korean/Traditional Chinese fonts → same-script fallback chain', () => {
    expect(cssFontFamily('맑은 고딕')).toBe(
      "'맑은 고딕','Malgun Gothic','GenOffice Sans KR','Apple SD Gothic Neo','Noto Sans KR',sans-serif",
    )
    expect(cssFontFamily('Batang')).toBe(
      "'Batang','GenOffice Batang','GenOffice Serif KR','GenOffice Myungjo','Noto Serif KR',serif",
    )
    expect(cssFontFamily('微軟正黑體')).toBe(
      "'微軟正黑體','Microsoft JhengHei','PingFang TC','GenOffice Heiti TC','Noto Sans TC',sans-serif",
    )
    expect(cssFontFamily('新細明體')).toBe(
      "'新細明體','GenOffice MingLiU','GenOffice Fullwidth TC','Songti TC','Noto Serif TC',serif",
    )
  })

  describe('KR chains (M3 probe: theme Latin head, real Nanum)', () => {
    afterEach(() => vi.restoreAllMocks())

    it('missing KR variant gets the theme Latin head ahead of the Batang chain', () => {
      expect(cssFontFamily('Noto Sans CJK KR')).toBe(
        "'Noto Sans CJK KR','KR Theme Latin GO','GenOffice Batang','GenOffice Serif KR','GenOffice Myungjo','Noto Serif KR',serif",
      )
    })

    it('installed KR variant keeps its chain without the theme Latin head', () => {
      stubCanvas(['Noto Serif CJK KR'])
      expect(cssFontFamily('Noto Serif CJK KR')).toBe(
        "'Noto Serif CJK KR','GenOffice Batang','GenOffice Serif KR','GenOffice Myungjo','Noto Serif KR',serif",
      )
    })

    it('Malgun/Batang declares keep their Latin-normalized subsets (no theme head)', () => {
      expect(cssFontFamily('Malgun Gothic')).not.toContain("'KR Theme Latin GO'")
      expect(cssFontFamily('Batang')).not.toContain("'KR Theme Latin GO'")
    })

    it('installed NanumGothic resolves at the literal head of the sans chain', () => {
      stubCanvas(['NanumGothic'])
      expect(cssFontFamily('NanumGothic')).toBe(
        "'NanumGothic','Malgun Gothic','GenOffice Sans KR','Apple SD Gothic Neo','Noto Sans KR',sans-serif",
      )
    })

    it('missing Nanum names keep the Batang-normalized serif chain', () => {
      expect(cssFontFamily('NanumBarunGothic')).toBe(
        "'NanumBarunGothic','GenOffice Batang','GenOffice Serif KR','GenOffice Myungjo','Noto Serif KR',serif",
      )
    })
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

  it('Korean font line factors: Batang-class 1.3029, Malgun 1.7371 (Word probe)', () => {
    expect(lineHeightFactor('Batang')).toBe(1.3029)
    expect(lineHeightFactor('바탕')).toBe(1.3029)
    expect(lineHeightFactor('Gulim')).toBe(1.3029)
    expect(lineHeightFactor('Dotum')).toBe(1.3029)
    expect(lineHeightFactor('NanumMyeongjo')).toBe(1.3029)
    expect(lineHeightFactor('Malgun Gothic')).toBe(1.7371)
    expect(lineHeightFactor('맑은 고딕')).toBe(1.7371)
    expect(lineHeightFactor('Noto Sans CJK KR')).toBe(1.3029)
    expect(lineHeightFactor('Noto Serif KR')).toBe(1.3029)
    expect(lineHeightFactor('Source Han Sans K')).toBe(1.3029)
  })

  it('Chinese/Japanese factors follow the Word probe (unprobed names keep LO values)', () => {
    expect(lineHeightFactor('SimSun')).toBe(1.3029)
    expect(lineHeightFactor('宋体')).toBe(1.3029)
    expect(lineHeightFactor('DengXian')).toBe(1.775)
    expect(lineHeightFactor('等线')).toBe(1.775)
    expect(lineHeightFactor('仿宋_GB2312')).toBe(1.775)
    expect(lineHeightFactor('黑体')).toBe(1.0)
    expect(lineHeightFactor('楷体')).toBe(1.0)
    expect(lineHeightFactor('楷体_GB2312')).toBe(1.775)
    expect(lineHeightFactor('ＭＳ 明朝')).toBe(1.3029)
    expect(lineHeightFactor('MS Gothic')).toBe(1.3029)
    expect(lineHeightFactor('游明朝')).toBe(1.44)
    expect(lineHeightFactor('Meiryo')).toBe(1.9429)
    expect(lineHeightFactor('PMingLiU')).toBe(1.3029)
    expect(lineHeightFactor('Microsoft JhengHei')).toBe(1.775)
    expect(lineHeightFactor('Calibri')).toBe(1.22)
    expect(lineHeightFactor('Century Gothic')).toBe(1.2)
  })

  it('missing SC-variant declares take the Word SimSun-substitution factor', () => {
    expect(lineHeightFactor('Noto Sans CJK SC')).toBe(1.3029)
    expect(lineHeightFactor('Noto Sans CJK TC')).toBe(1.3029)
    expect(lineHeightFactor('Noto Sans HK')).toBe(1.3029)
    expect(lineHeightFactor('Noto Serif SC')).toBe(1.3029)
    expect(lineHeightFactor('Source Han Sans CN')).toBe(1.3029)
    expect(lineHeightFactor('Noto Sans SC')).toBe(1.3029)
  })

  it('cjkDeclaredLineFactor maps regional Noto/Source Han variants for CJK runs', () => {
    expect(cjkDeclaredLineFactor('Noto Sans CJK JP')).toBe(1.3029)
    expect(cjkDeclaredLineFactor('Noto Serif JP')).toBe(1.3029)
    expect(cjkDeclaredLineFactor('Source Han Sans JP')).toBe(1.3029)
    expect(cjkDeclaredLineFactor('Noto Sans CJK TC')).toBe(1.3029)
    expect(cjkDeclaredLineFactor('Noto Sans HK')).toBe(1.3029)
    expect(cjkDeclaredLineFactor('Noto Sans CJK KR')).toBe(1.3029)
    expect(cjkDeclaredLineFactor('Noto Sans CJK SC')).toBe(1.3029)
    expect(cjkDeclaredLineFactor('Source Han Sans CN')).toBe(1.3029)
    expect(cjkDeclaredLineFactor('SimSun')).toBeNull()
    expect(cjkDeclaredLineFactor('Calibri')).toBeNull()
  })

  it('Tamil faces take the Latha factor (Word probe 2026-08-13)', () => {
    expect(lineHeightFactor('Noto Sans Tamil')).toBe(1.6686)
    expect(lineHeightFactor('Latha')).toBe(1.6686)
    expect(lineHeightFactor('Vijaya')).toBe(1.6686)
    expect(lineHeightFactor('InaiMathi')).toBe(1.6686)
  })

  it('textHasHangul separates Korean from other CJK', () => {
    expect(textHasHangul('보고서 2026')).toBe(true)
    expect(textHasHangul('\u4e2d\u6587')).toBe(false)
    expect(textHasHangul('かな')).toBe(false)
  })

  it('paraLineFactorCss routes by script', () => {
    expect(paraLineFactorCss('한국어')).toBe('var(--doc-line-factor-kr,1.3029)')
    expect(paraLineFactorCss('\u4e2d\u6587')).toBe('var(--doc-line-factor-cjk,1.7)')
    expect(paraLineFactorCss('latin')).toBe('var(--doc-line-factor-latin,1.2)')
  })

  it('krLineFactor follows the EA face, defaulting to Batang-class', () => {
    expect(krLineFactor('Batang')).toBe(1.3029)
    expect(krLineFactor('맑은 고딕')).toBe(1.7371)
    expect(krLineFactor(undefined)).toBe(1.3029)
  })

  it('Korean ascii face in a dual-slot chain keeps only the literal family', () => {
    expect(cssDualFontFamily('맑은 고딕', 'Batang')).toBe(
      "'맑은 고딕','Batang','GenOffice Batang','GenOffice Serif KR','GenOffice Myungjo','Noto Serif KR',serif",
    )
  })

  it('mono ascii face keeps its mono faces, CJK falls through to the eastAsia font', () => {
    expect(cssDualFontFamily('Consolas', 'SimSun')).toBe(
      "'Consolas','Menlo','Courier New','Liberation Mono','SimSun','GenOffice Songti SC','STSong','Noto Serif CJK SC',serif",
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
  it('missing naskh/serif-class names get a Times Latin head and the 90% TNR alias', () => {
    expect(cssFontFamily('Noto Naskh Arabic')).toBe(
      "'Noto Naskh Arabic W','Geeza Pro','Al Bayan',serif",
    )
    expect(cssFontFamily('Arabic Typesetting')).toBe(
      "'Arabic Typesetting','Times New Roman','Liberation Serif','Naskh Digits GO','Noto Naskh Arabic TNR','Geeza Pro','Al Bayan',serif",
    )
    expect(cssFontFamily('Amiri')).toContain("'Noto Naskh Arabic TNR'")
    expect(cssFontFamily('Scheherazade New')).toContain("'Noto Naskh Arabic TNR'")
  })

  it('unscaled digits alias precedes the TNR alias only in missing-serif chains', () => {
    const missing = cssFontFamily('Arabic Typesetting')
    expect(missing.indexOf("'Naskh Digits GO'")).toBeLessThan(
      missing.indexOf("'Noto Naskh Arabic TNR'"),
    )
    expect(cssFontFamily('Traditional Arabic')).not.toContain("'Naskh Digits GO'")
    expect(cssFontFamily('Noto Naskh Arabic')).not.toContain("'Naskh Digits GO'")
    expect(cssFontFamily('Noto Kufi Arabic')).not.toContain("'Naskh Digits GO'")
  })

  it('Traditional/Simplified Arabic map to the size-adjusted alias (no Times head)', () => {
    expect(cssFontFamily('Traditional Arabic')).toBe(
      "'Traditional Arabic','Noto Naskh Arabic TA','Geeza Pro','Al Bayan',serif",
    )
    expect(cssFontFamily('Traditional Arabic')).not.toContain("'Times New Roman'")
    expect(cssFontFamily('Simplified Arabic')).toContain("'Noto Naskh Arabic TA'")
    // other missing naskh-class names take the TNR calibration, not TA
    expect(cssFontFamily('Amiri')).not.toContain("'Noto Naskh Arabic TA'")
    expect(cssFontFamily('Arabic Typesetting')).not.toContain("'Noto Naskh Arabic TA'")
  })

  it('kufi/sans-class names get the Sans Arabic chain (no Times head)', () => {
    expect(cssFontFamily('Noto Sans Arabic')).toBe("'Noto Sans Arabic W','Geeza Pro',sans-serif")
    expect(cssFontFamily('Noto Kufi Arabic')).toBe(
      "'Noto Kufi Arabic','Noto Sans Arabic','Geeza Pro',sans-serif",
    )
  })

  it('unknown Arabic names (by script in the name) default to the naskh chain', () => {
    expect(cssFontFamily('الخط الديواني')).toBe(
      "'الخط الديواني','Times New Roman','Liberation Serif','Naskh Digits GO','Noto Naskh Arabic TNR','Geeza Pro','Al Bayan',serif",
    )
    expect(cssFontFamily('Urdu Typesetting')).toContain("'Noto Naskh Arabic TNR'")
  })

  it('Iranian B/XB/IR faces take the naskh chain and factor (Word substitutes Times New Roman)', () => {
    for (const name of ['B Mitra', 'B Nazanin', 'B Titr', 'XB Zar', 'IRLotus', 'B Yekan']) {
      const chain = cssFontFamily(name)
      // Times before the subset: ASCII punctuation/digits take Times shapes as in Word
      expect(chain).toContain("'Times New Roman'")
      expect(chain).toContain("'Noto Naskh Arabic TNR'")
      expect(chain.indexOf("'Times New Roman'")).toBeLessThan(
        chain.indexOf("'Noto Naskh Arabic TNR'"),
      )
      expect(chain).toMatch(/,serif$/)
      expect(lineHeightFactor(name)).toBe(1.1429)
    }
  })

  it('does not capture CJK or Latin families', () => {
    expect(cssFontFamily('SimSun')).not.toContain('Arabic')
    expect(cssFontFamily('Calibri')).not.toContain('Arabic')
    expect(cssFontFamily('Batang')).not.toContain('Arabic')
    expect(cssFontFamily('SomeCustomFont')).not.toContain('Arabic')
    expect(cssFontFamily('Bahnschrift')).not.toContain('Arabic')
    expect(lineHeightFactor('Bahnschrift')).toBe(1.2)
  })
})

describe('textHasComplexScript', () => {
  it('detects Arabic, Hebrew and presentation forms', () => {
    expect(textHasComplexScript('مرحبا')).toBe(true)
    expect(textHasComplexScript('שלום')).toBe(true)
    expect(textHasComplexScript('ﻻ')).toBe(true)
  })

  it('detects Indic and Thai scripts (Word routes them through the cs slot)', () => {
    expect(textHasComplexScript('தமிழ்')).toBe(true)
    expect(textHasComplexScript('हिन्दी')).toBe(true)
    expect(textHasComplexScript('ไทย')).toBe(true)
  })

  it('is false for Latin and CJK', () => {
    expect(textHasComplexScript('hello 123')).toBe(false)
    expect(textHasComplexScript('\u4e2d\u6587')).toBe(false)
    expect(textHasComplexScript('かな한글')).toBe(false)
  })
})

describe('cssCsFontFamily', () => {
  it('cs chain leads; punct alias then subset, ascii chain wins Latin letters, Times backstops', () => {
    expect(cssCsFontFamily('Arabic Typesetting', 'Calibri', 'Calibri')).toBe(
      "'Arabic Typesetting','Naskh Digits GO','Times Punct GO','Noto Naskh Arabic TNR','Calibri','Carlito GO','Noto Sans CJK SC','Times New Roman','Liberation Serif','Geeza Pro','Al Bayan',sans-serif",
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
    expect(
      chain.startsWith(
        "'Amiri','Naskh Digits GO','Times Punct GO','Noto Naskh Arabic TNR','Times New Roman'",
      ),
    ).toBe(true)
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
      "'Noto Naskh Arabic W','Geeza Pro','Al Bayan',serif",
    )
  })

  it('deduplicates shared families', () => {
    const chain = cssCsFontFamily('Noto Naskh Arabic', 'Noto Naskh Arabic')
    expect(chain.match(/'Noto Naskh Arabic W'/g)?.length).toBe(1)
  })
})

// ─── Footnote estimate model ──────────────────────────────────────────────

describe('estimateFootnoteHeight / footnoteLineHeightPx', () => {
  it('footnoteLineHeightPx is the 10pt default-font single-line height', () => {
    expect(footnoteLineHeightPx(undefined)).toBeCloseTo(
      10 * (96 / 72) * lineHeightFactor('Calibri'),
      5,
    )
  })

  it('Latin estimate is a whole multiple of footnoteLineHeightPx (render line-height shares the value)', () => {
    const lineH = footnoteLineHeightPx(undefined)
    const short = estimateFootnoteHeight('note', 600, undefined)
    expect(short).toBeCloseTo(lineH, 5)
    const long = estimateFootnoteHeight('word '.repeat(80).trim(), 600, undefined)
    expect(long / lineH).toBeCloseTo(Math.round(long / lineH), 5)
    expect(long).toBeGreaterThan(short)
  })

  it('docGrid line pitch snaps estimate and line height together', () => {
    const docGrid = { type: 'lines' as const, linePitch: 312 }
    const lineH = footnoteLineHeightPx(docGrid)
    expect(lineH).toBeCloseTo(312 * TWIPS_TO_PX, 5)
    expect(estimateFootnoteHeight('note', 600, docGrid)).toBeCloseTo(lineH, 5)
  })
})
