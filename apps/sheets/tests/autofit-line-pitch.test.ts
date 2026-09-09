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
    expect(excelRowPitchPx('Calibri', 12, noMetrics)).toBe(21) // 15.75 pt
    expect(excelRowPitchPx('Calibri', 14, noMetrics)).toBe(25) // 18.75 pt
    expect(excelRowPitchPx('Arial', 10, noMetrics)).toBe(17) // 12.75 pt
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
