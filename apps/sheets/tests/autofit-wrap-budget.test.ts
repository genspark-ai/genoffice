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
import { describe, expect, it } from 'vitest'

import {
  AUTOFIT_EXTRA_PX,
  installAutofitWrapBudget,
  type AutofitSkeletonLike,
} from '../src/renderer/autofit-wrap-budget'

function fakeSkeleton(columnData: Record<number, { w?: number; hd?: number } | undefined>) {
  const seen: number[] = []
  const proto: AutofitSkeletonLike & { seen: number[]; fail?: boolean } = {
    seen,
    _worksheetData: { columnData, defaultColumnWidth: 72 },
    calculateAutoHeightForCell(_row, col) {
      if (this.fail) throw new Error('measure failed')
      const width = this._worksheetData.columnData[col]?.w ?? this._worksheetData.defaultColumnWidth
      seen.push(width)
      return width
    },
  }
  installAutofitWrapBudget(proto)
  return proto
}

describe('installAutofitWrapBudget', () => {
  it('narrows only the measured column by the Excel autofit slack', () => {
    const skeleton = fakeSkeleton({ 3: { w: 171, hd: 0 } })
    expect(skeleton.calculateAutoHeightForCell(29, 3)).toBe(171 - AUTOFIT_EXTRA_PX)
    expect(skeleton._worksheetData.columnData[3]).toEqual({ w: 171, hd: 0 })
  })

  it('handles default-width columns without leaving an entry behind', () => {
    const skeleton = fakeSkeleton({})
    expect(skeleton.calculateAutoHeightForCell(0, 5)).toBe(72 - AUTOFIT_EXTRA_PX)
    expect(5 in skeleton._worksheetData.columnData).toBe(false)
  })

  it('restores the column when the measure throws', () => {
    const skeleton = fakeSkeleton({ 1: { w: 100 } })
    skeleton.fail = true
    expect(() => skeleton.calculateAutoHeightForCell(0, 1)).toThrow('measure failed')
    expect(skeleton._worksheetData.columnData[1]).toEqual({ w: 100 })
  })
})

// The real measure reads `_worksheetData.columnData[col].w`, not
// columnWidthAccumulation; this drives it end to end with a fake 2D context.
describe('installAutofitWrapBudget on the real SpreadsheetSkeleton', () => {
  const PX_PER_CHAR = 7
  const ASCENT = 12
  const DESCENT = 3
  const COLUMN_WIDTH = 100
  // 13 chars = 91 px: under the stock 96 px budget, over the patched 84 px.
  const TEXT = 'ABCDEF GHIJKL'

  function realSkeleton() {
    ;(FontCache as unknown as { _context: unknown })._context = {
      font: '',
      textBaseline: 'alphabetic',
      measureText: (text: string) => ({
        width: text.length * PX_PER_CHAR,
        fontBoundingBoxAscent: ASCENT,
        fontBoundingBoxDescent: DESCENT,
        actualBoundingBoxAscent: ASCENT,
        actualBoundingBoxDescent: DESCENT,
      }),
    }
    const worksheet = new Worksheet(
      'wb',
      {
        id: 's1',
        name: 'S1',
        rowCount: 3,
        columnCount: 3,
        columnData: { 0: { w: COLUMN_WIDTH } },
        cellData: { 0: { 0: { v: TEXT, s: { tb: WrapStrategy.WRAP } } } },
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

  it('wraps a 93-96% wide text to two lines only with the budget installed', () => {
    const skeleton = realSkeleton()
    const oneLine = ASCENT + DESCENT + DEFAULT_PADDING_DATA.t + DEFAULT_PADDING_DATA.b
    expect(skeleton.calculateAutoHeightForCell(0, 0)).toBe(oneLine)

    installAutofitWrapBudget()
    expect(skeleton.calculateAutoHeightForCell(0, 0)).toBe(oneLine + ASCENT + DESCENT)
    expect(skeleton.worksheet.getConfig().columnData[0]).toEqual({ w: COLUMN_WIDTH })
    expect(skeleton.columnWidthAccumulation[0]).toBe(COLUMN_WIDTH)
  })
})
