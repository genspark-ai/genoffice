import {
  ConfigService,
  ContextService,
  type Injector,
  LocaleService,
  Styles,
  Worksheet,
  WrapStrategy,
} from '@univerjs/core'
import { DEFAULT_PADDING_DATA, FontCache, SpreadsheetSkeleton } from '@univerjs/engine-render'
import { beforeAll, describe, expect, it } from 'vitest'

import { excelRowPitchPx, installAutofitLinePitch } from '../src/renderer/autofit-line-pitch'

const noMetrics = { fontBoundingBoxAscent: 0, fontBoundingBoxDescent: 0 }

describe('excelRowPitchPx', () => {
  it('reproduces the Excel default row heights of the listed fonts', () => {
    expect(excelRowPitchPx('Calibri', 10, noMetrics)).toBe(17) // 12.75 pt
    expect(excelRowPitchPx('Calibri', 11, noMetrics)).toBe(20) // 15 pt
    expect(excelRowPitchPx('Calibri', 14, noMetrics)).toBe(25) // 18.75 pt
    expect(excelRowPitchPx('Arial', 10, noMetrics)).toBe(17) // 12.75 pt
  })

  it('uses the Mac Excel pitch where reference prints differ from the GDI formula', () => {
    // Windows: Arial 11 -> 14.25 pt, Calibri 12 -> 15.75 pt; Mac Excel opens
    // both at 15 pt (two-line Arial 11 wraps print 30 pt).
    expect(excelRowPitchPx('Arial', 11, noMetrics)).toBe(20)
    expect(excelRowPitchPx('Calibri', 12, noMetrics)).toBe(20)
  })

  it('covers the CJK office faces at their Excel row heights', () => {
    expect(excelRowPitchPx('Malgun Gothic', 10, noMetrics)).toBe(19) // 14.25 pt
    expect(excelRowPitchPx('Malgun Gothic', 11, noMetrics)).toBe(22) // 16.5 pt (Korean Excel default)
    expect(excelRowPitchPx('\uB9D1\uC740 \uACE0\uB515', 12, noMetrics)).toBe(23) // 17.25 pt
    expect(excelRowPitchPx('Meiryo', 10, noMetrics)).toBe(22) // 16.5 pt
    expect(excelRowPitchPx('\u30E1\u30A4\u30EA\u30AA', 11, noMetrics)).toBe(25) // 18.75 pt
    expect(excelRowPitchPx('Meiryo UI', 10, noMetrics)).toBe(19) // 14.25 pt
    expect(excelRowPitchPx('Noto Sans CJK SC', 10, noMetrics)).toBe(21) // 15.75 pt
    expect(excelRowPitchPx('Microsoft YaHei', 14, noMetrics)).toBe(27) // 20.25 pt
    expect(excelRowPitchPx('\uFF2D\uFF33 \uFF30\u30B4\u30B7\u30C3\u30AF', 11, noMetrics)).toBe(17)
  })

  it('counts the Yu family line gap instead of the 2 px padding', () => {
    expect(excelRowPitchPx('Yu Gothic', 11, noMetrics)).toBe(25) // 18.75 pt (Japanese Excel default)
    expect(excelRowPitchPx('\u6E38\u30B4\u30B7\u30C3\u30AF', 12, noMetrics)).toBe(26) // 19.5 pt
    expect(excelRowPitchPx('Yu Gothic UI', 11, noMetrics)).toBe(22) // no line gap
  })

  it('accepts quoted and comma-separated family strings', () => {
    expect(excelRowPitchPx('"Times New Roman", serif', 10, noMetrics)).toBe(17)
  })

  it('falls back to the canvas font box for unlisted families', () => {
    // 13.33 px em with a 0.9 / 0.25 box: round(12) + round(3.33) + 2
    const metrics = { fontBoundingBoxAscent: 12, fontBoundingBoxDescent: 3.333 }
    expect(excelRowPitchPx('Some Font', 10, metrics)).toBe(17)
    expect(excelRowPitchPx('Some Font', 10, noMetrics)).toBe(0)
  })
})

// Real Worksheet + SpreadsheetSkeleton driven by a fake 2D context whose
// font box is Calibri's hhea pair (0.75 / 0.25 em) and whose glyphs are half
// an em wide.
describe('installAutofitLinePitch on the real SpreadsheetSkeleton', () => {
  const COLUMN_WIDTH = 100
  const DEFAULT_ROW_HEIGHT = 20

  function fontPx(font: string): number {
    const match = /([\d.]+)pt/.exec(font)
    return match ? (Number(match[1]) * 96) / 72 : 0
  }

  beforeAll(() => {
    const ctx = {
      font: '',
      textBaseline: 'alphabetic',
      measureText(text: string) {
        const px = fontPx(this.font)
        return {
          width: text.length * px * 0.5,
          fontBoundingBoxAscent: px * 0.75,
          fontBoundingBoxDescent: px * 0.25,
          actualBoundingBoxAscent: px * 0.75,
          actualBoundingBoxDescent: px * 0.25,
        }
      },
    }
    ;(FontCache as unknown as { _context: unknown })._context = ctx
    installAutofitLinePitch()
  })

  function skeletonFor(cells: Record<number, Record<number, unknown>>) {
    const worksheet = new Worksheet(
      'wb',
      {
        id: 's1',
        name: 'S1',
        rowCount: 8,
        columnCount: 2,
        defaultRowHeight: DEFAULT_ROW_HEIGHT,
        columnData: { 0: { w: COLUMN_WIDTH } },
        cellData: cells as never,
      },
      new Styles(),
    )
    return new SpreadsheetSkeleton(
      worksheet,
      new Styles(),
      new LocaleService(),
      new ContextService(),
      new ConfigService(),
      undefined as unknown as Injector,
    )
  }

  // Each 12-char word is 80 px at 10 pt (86.7 px at 11 pt): one word per
  // 96 px line, so the word count is the line count.
  const word = 'ABCDEFGHIJKL'
  const wrap = (fs: number) => ({ tb: WrapStrategy.WRAP, ff: 'Calibri', fs })

  it('measures a five-line Calibri 10 cell as 5 x 12.75 pt', () => {
    const skeleton = skeletonFor({
      0: { 0: { v: Array(5).fill(word).join(' '), s: wrap(10) } },
    })
    expect(skeleton.calculateAutoHeightForCell(0, 0)).toBe(5 * 17)
  })

  it('measures a three-line Calibri 11 cell as 3 x 15 pt', () => {
    const skeleton = skeletonFor({
      0: { 0: { v: Array(3).fill(word).join(' '), s: wrap(11) } },
    })
    expect(skeleton.calculateAutoHeightForCell(0, 0)).toBe(3 * 20)
  })

  // A cell document the way the loader builds one for manual line breaks:
  // \r paragraph breaks, one run carrying the cell font.
  const doc = (lines: string[], fs: number, ff = 'Calibri') => {
    const text = lines.join('\r')
    const dataStream = `${text}\r\n`
    const paragraphs = [...dataStream].flatMap((ch, i) => (ch === '\r' ? [{ startIndex: i }] : []))
    return {
      id: 'rich-cell',
      body: {
        dataStream,
        textRuns: [{ st: 0, ed: text.length, ts: { ff, fs } }],
        paragraphs,
        sectionBreaks: [{ startIndex: dataStream.length - 1 }],
      },
      documentStyle: {},
    }
  }

  it('rescales a manual-break document in one font to lines x pitch', () => {
    const skeleton = skeletonFor({
      0: { 0: { p: doc(Array(4).fill(word), 10), s: wrap(10) } },
      1: { 0: { p: doc(Array(3).fill(word), 11, 'Malgun Gothic'), s: wrap(11) } },
    })
    expect(skeleton.calculateAutoHeightForCell(0, 0)).toBe(4 * 17)
    expect(skeleton.calculateAutoHeightForCell(1, 0)).toBe(3 * 22)
  })

  it('leaves mixed-font rich text on the stock measure', () => {
    const rich = doc(Array(2).fill(word), 10)
    rich.body.textRuns = [
      { st: 0, ed: word.length, ts: { ff: 'Calibri', fs: 10 } },
      { st: word.length, ed: rich.body.dataStream.length - 2, ts: { ff: 'Calibri', fs: 14 } },
    ]
    const skeleton = skeletonFor({ 0: { 0: { p: rich, s: wrap(10) } } })
    const stock = skeleton.calculateAutoHeightForCell(0, 0) ?? 0
    expect(stock).toBeGreaterThan(0)
    expect(stock % 17).not.toBe(0)
  })

  it('keeps a single-line wrap row at the sheet default', () => {
    const skeleton = skeletonFor({ 0: { 0: { v: word, s: wrap(10) } } })
    expect(skeleton.calculateAutoHeightForCell(0, 0)).toBe(17)
    expect(
      skeleton.calculateAutoHeightInRange([
        { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
      ]),
    ).toEqual([{ row: 0, autoHeight: DEFAULT_ROW_HEIGHT }])
  })

  it('leaves non-wrap cells on the stock measure', () => {
    const skeleton = skeletonFor({ 0: { 0: { v: word, s: { ff: 'Calibri', fs: 10 } } } })
    const em = (10 * 96) / 72
    expect(skeleton.calculateAutoHeightForCell(0, 0)).toBeCloseTo(
      em + DEFAULT_PADDING_DATA.t + DEFAULT_PADDING_DATA.b,
    )
  })
})
