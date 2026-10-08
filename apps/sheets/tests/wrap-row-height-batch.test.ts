import { FontCache } from '@univerjs/engine-render'
import { describe, expect, it, vi } from 'vitest'

import type { WorkbookCellStyle } from '../src/shared/desktop-api'
import {
  type BatchWrapHeightInput,
  batchWrapRowHeights,
  createUniverWrapMeasurer,
  type WrapCellMeasurer,
} from '../src/renderer/wrap-row-height-batch'
import { wrapAutoFitRows } from '../src/renderer/univer-sync'

const WRAP = { wrapText: true } as WorkbookCellStyle
const PLAIN = { wrapText: false } as WorkbookCellStyle
const BIG = { wrapText: false, fontSize: 24 } as WorkbookCellStyle
const styles = [PLAIN, WRAP, BIG]

/// 20px per line of 10 characters at a 100px column; plain cells one line,
/// a 24pt font two.
const measurer: WrapCellMeasurer = {
  cellHeight(cell, style, wraps, columnWidthPx) {
    if (cell.rich?.length) return null
    if (!wraps || typeof cell.value !== 'string') return style?.fontSize === 24 ? 40 : 20
    const perLine = Math.max(1, Math.floor(columnWidthPx / 10))
    return Math.ceil(cell.value.length / perLine) * 20
  },
}

function input(partial: Partial<BatchWrapHeightInput>): BatchWrapHeightInput {
  return {
    cells: [],
    styles,
    rows: [],
    merges: undefined,
    inheritedWrap: () => false,
    columnWidthPx: () => 100,
    defaultRowHeightPx: 20,
    currentAutoHeight: () => undefined,
    keepTaller: true,
    ...partial,
  }
}

describe('batchWrapRowHeights', () => {
  it('fits each row once to its tallest cell, floored at the default', () => {
    const result = batchWrapRowHeights(
      input({
        rows: [0, 1, 2],
        cells: [
          { row: 0, column: 0, value: 'x'.repeat(25), styleIndex: 1 },
          { row: 0, column: 1, value: 'x'.repeat(12), styleIndex: 1 },
          { row: 1, column: 0, value: 'short', styleIndex: 1 },
          { row: 2, column: 3, value: 'x'.repeat(45), styleIndex: 1 },
        ],
      }),
      measurer,
    )
    expect(result.heights).toEqual([
      { row: 0, autoHeight: 60 },
      { row: 1, autoHeight: 20 },
      { row: 2, autoHeight: 100 },
    ])
    expect(result.fallbackRows).toEqual([])
  })

  it('counts the plain cells of a wrap row like Univer does', () => {
    const result = batchWrapRowHeights(
      input({
        rows: [0],
        cells: [
          { row: 0, column: 0, value: 'one line', styleIndex: 1 },
          { row: 0, column: 1, value: 12, styleIndex: 2 },
        ],
      }),
      measurer,
    )
    expect(result.heights).toEqual([{ row: 0, autoHeight: 40 }])
  })

  it('ignores rows it was not asked for and empty cells', () => {
    const result = batchWrapRowHeights(
      input({
        rows: [1],
        cells: [
          { row: 0, column: 0, value: 'x'.repeat(50), styleIndex: 1 },
          { row: 1, column: 0, value: '', styleIndex: 1 },
          { row: 1, column: 1, value: null, styleIndex: 1 },
        ],
      }),
      measurer,
    )
    expect(result.heights).toEqual([{ row: 1, autoHeight: 20 }])
  })

  it('never fits a merged wrap cell, even one spanning rows', () => {
    const result = batchWrapRowHeights(
      input({
        rows: [0, 1],
        cells: [
          { row: 0, column: 0, value: 'x'.repeat(50), styleIndex: 1 },
          { row: 1, column: 0, value: 'x'.repeat(50), styleIndex: 1 },
          { row: 1, column: 2, value: 'x'.repeat(15), styleIndex: 1 },
        ],
        merges: [{ startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }],
      }),
      measurer,
    )
    expect(result.heights).toEqual([
      { row: 0, autoHeight: 20 },
      { row: 1, autoHeight: 40 },
    ])
  })

  it('leaves customHeight rows alone through wrapAutoFitRows', () => {
    const cells = [
      { row: 0, column: 0, value: 'x'.repeat(50), styleIndex: 1 },
      { row: 1, column: 0, value: 'x'.repeat(50), styleIndex: 1 },
      { row: 2, column: 0, value: 'x'.repeat(50), styleIndex: 1 },
    ]
    const rows = wrapAutoFitRows(
      cells as never,
      styles,
      [
        { row: 0, height: 30, customHeight: true, hidden: false },
        { row: 1, height: 30, hidden: false },
        { row: 2, height: 3, hidden: false },
      ] as never,
      [],
      false,
      15,
      { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    )
    expect(rows).toEqual([1])
    const result = batchWrapRowHeights(input({ rows, cells }), measurer)
    expect(result.heights).toEqual([{ row: 1, autoHeight: 100 }])
  })

  it('uses the column width and inherited wrap for unstyled cells', () => {
    const result = batchWrapRowHeights(
      input({
        rows: [0],
        cells: [{ row: 0, column: 4, value: 'x'.repeat(40) }],
        inheritedWrap: (_row, column) => column === 4,
        columnWidthPx: (column) => (column === 4 ? 200 : 100),
      }),
      measurer,
    )
    expect(result.heights).toEqual([{ row: 0, autoHeight: 40 }])
  })

  it('keeps a taller earlier measure unless told to shrink', () => {
    const base = input({
      rows: [0, 1],
      cells: [
        { row: 0, column: 0, value: 'x'.repeat(15), styleIndex: 1 },
        { row: 1, column: 0, value: 'x'.repeat(15), styleIndex: 1 },
      ],
      currentAutoHeight: (row) => (row === 0 ? 90 : 25),
    })
    expect(batchWrapRowHeights(base, measurer).heights).toEqual([
      { row: 0, autoHeight: 90 },
      { row: 1, autoHeight: 40 },
    ])
    expect(batchWrapRowHeights({ ...base, keepTaller: false }, measurer).heights).toEqual([
      { row: 0, autoHeight: 40 },
      { row: 1, autoHeight: 40 },
    ])
  })

  it('hands rows with rich text to the per-row measure', () => {
    const result = batchWrapRowHeights(
      input({
        rows: [0, 1],
        cells: [
          { row: 0, column: 0, value: 'x'.repeat(15), styleIndex: 1 },
          { row: 1, column: 0, value: 'x'.repeat(15), styleIndex: 1 },
          { row: 1, column: 1, value: 'ab', styleIndex: 1, rich: [{ text: 'a' }, { text: 'b' }] },
        ] as never,
      }),
      measurer,
    )
    expect(result.heights).toEqual([{ row: 0, autoHeight: 40 }])
    expect(result.fallbackRows).toEqual([1])
  })

  it('clamps at the Univer row maximum', () => {
    const result = batchWrapRowHeights(
      input({ rows: [0], cells: [{ row: 0, column: 0, value: 'x'.repeat(5000), styleIndex: 1 }] }),
      measurer,
    )
    expect(result.heights).toEqual([{ row: 0, autoHeight: 2000 }])
  })
})

describe('createUniverWrapMeasurer', () => {
  it('resolves each style once and declines rich or rotated cells', () => {
    vi.spyOn(FontCache, 'getMeasureText').mockReturnValue({
      width: 8,
      fontBoundingBoxAscent: 12,
      fontBoundingBoxDescent: 3,
    } as never)
    let conversions = 0
    const measurer = createUniverWrapMeasurer(() => {
      conversions += 1
      return {}
    })
    const rotated = { wrapText: true, textRotation: 45 } as WorkbookCellStyle
    expect(
      measurer.cellHeight(
        { row: 0, column: 0, value: 'a', rich: [{ text: 'a' }] } as never,
        WRAP,
        true,
        100,
      ),
    ).toBeNull()
    expect(measurer.cellHeight({ row: 0, column: 0, value: 'a' }, rotated, true, 100)).toBeNull()
    expect(conversions).toBe(0)
    const cells = [
      { row: 0, column: 0, value: 'a', styleIndex: 1 },
      { row: 1, column: 0, value: 'b', styleIndex: 1 },
      { row: 2, column: 0, value: 'c', styleIndex: 0 },
    ]
    const heights = cells.map((cell) =>
      measurer.cellHeight(cell, styles[cell.styleIndex], cell.styleIndex === 1, 100),
    )
    expect(heights.every((height) => typeof height === 'number' && height > 0)).toBe(true)
    expect(conversions).toBe(2)
    expect(measurer.fontEntries).toBe(2)
  })
})
