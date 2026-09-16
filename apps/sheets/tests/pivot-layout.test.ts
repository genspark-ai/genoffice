import { describe, expect, it } from 'vitest'
import {
  areasOverlap,
  buildPivotLayout,
  PivotLayoutError,
  pivotOutputArea,
} from '@genoffice/xlsx-gateway/domain/pivot-layout'

const GRID = [
  ['region', 'product', 'amount'],
  ['North', 'A', 10],
  ['South', 'B', 20],
  ['North', 'B', 5],
  ['West', 'A', 7],
]

describe('buildPivotLayout', () => {
  it('bakes a single-level pivot with captions, a grand total and value formats', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: 'region',
      values: [
        { field: 'amount', agg: 'sum', numFmt: '#,##0' },
        { field: 'amount', agg: 'count' },
      ],
    })
    expect(layout.matrix).toEqual([
      ['region', 'Sum of amount', 'Count of amount'],
      ['North', 15, 2],
      ['South', 20, 1],
      ['West', 7, 1],
      ['Grand Total', 42, 4],
    ])
    expect(layout.width).toBe(3)
    expect(layout.height).toBe(5)
    expect(layout.numberFormats).toEqual([{ columnOffset: 1, format: '#,##0' }])
    expect(layout.definition).toMatchObject({
      fieldNames: ['region', 'product', 'amount'],
      rowFieldIndices: [0],
      rowItems: ['North', 'South', 'West'],
      rowLines: [
        { t: 'data', members: [0] },
        { t: 'data', members: [1] },
        { t: 'data', members: [2] },
      ],
      values: [
        { fieldIndex: 2, agg: 'sum', numFmt: '#,##0' },
        { fieldIndex: 2, agg: 'count' },
      ],
    })
    expect(pivotOutputArea({ row: 0, column: 4 }, layout)).toEqual({
      startRow: 0,
      startColumn: 4,
      endRow: 4,
      endColumn: 6,
    })
  })

  it('spreads a column field across columns and nests row levels with subtotals', () => {
    const columns = buildPivotLayout(GRID, {
      rowFields: 'region',
      columnField: 'product',
      values: [{ field: 'amount', agg: 'sum' }],
    })
    expect(columns.matrix).toEqual([
      ['region', 'A', 'B', 'Grand Total'],
      ['North', 10, 5, 15],
      ['South', null, 20, 20],
      ['West', 7, null, 7],
      ['Grand Total', 17, 25, 42],
    ])
    expect(columns.definition.columnFieldIndex).toBe(1)
    expect(columns.definition.columnItems).toEqual(['A', 'B'])
    expect(columns.numberFormats).toEqual([])

    const nested = buildPivotLayout(
      GRID,
      { rowFields: ['region', 'product'], values: [{ field: 'amount', agg: 'sum' }] },
      { subtotal: 'Sub', grandTotal: 'All' },
    )
    expect(nested.matrix).toEqual([
      ['region', 'product', 'Sum of amount'],
      ['North', 'A', 10],
      [null, 'B', 5],
      ['North', 'Sub', 15],
      ['South', 'B', 20],
      ['South', 'Sub', 20],
      ['West', 'A', 7],
      ['West', 'Sub', 7],
      ['All', '', 42],
    ])
    expect(nested.definition.rowLevelItems).toEqual([
      ['North', 'South', 'West'],
      ['A', 'B'],
    ])
    expect(nested.definition.rowLines.map((line) => line.t)).toEqual([
      'data',
      'data',
      'default',
      'data',
      'default',
      'data',
      'default',
    ])
  })

  it('keeps nested paths distinct when member labels concatenate alike', () => {
    const layout = buildPivotLayout(
      [
        ['a', 'b', 'n'],
        ['ab', 'c', 1],
        ['a', 'bc', 2],
      ],
      { rowFields: ['a', 'b'], values: [{ field: 'n', agg: 'sum' }] },
    )
    expect(layout.matrix).toEqual([
      ['a', 'b', 'Sum of n'],
      ['ab', 'c', 1],
      ['ab', 'Subtotal', 1],
      ['a', 'bc', 2],
      ['a', 'Subtotal', 2],
      ['Grand Total', '', 3],
    ])
  })

  it('rejects bad sources and fields with coded errors', () => {
    expect(() => buildPivotLayout([GRID[0]!], { rowFields: 'region', values: [] })).toThrow(
      PivotLayoutError,
    )
    expect(() =>
      buildPivotLayout(GRID, { rowFields: 'nope', values: [{ field: 'amount', agg: 'sum' }] }),
    ).toThrow('Field "nope" is not a source header.')
    try {
      buildPivotLayout(
        [
          ['a', 'a'],
          [1, 2],
        ],
        { rowFields: 'a', values: [{ field: 'a', agg: 'sum' }] },
      )
    } catch (err) {
      expect((err as PivotLayoutError).code).toBe('headerDuplicate')
    }
    expect(
      areasOverlap(
        { startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 },
        { startRow: 4, startColumn: 2, endRow: 9, endColumn: 9 },
      ),
    ).toBe(true)
    expect(
      areasOverlap(
        { startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 },
        { startRow: 5, startColumn: 0, endRow: 9, endColumn: 9 },
      ),
    ).toBe(false)
  })
})
