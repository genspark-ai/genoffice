import { describe, expect, it } from 'vitest'

import { chartCategoryVector, seriesOrientation } from '../src/renderer/chart-category-vector'

describe('chartCategoryVector', () => {
  it('flattens a single row or column unchanged', () => {
    expect(chartCategoryVector([['a', 'b', 'c']], 3)).toEqual(['a', 'b', 'c'])
    expect(chartCategoryVector([['a'], ['b']], 2)).toEqual(['a', 'b'])
  })

  it('keeps the innermost row of a two-row category range for a row series', () => {
    const grid = [
      ['', '', ''],
      ['Q1', 'Q2', 'Q3'],
    ]
    expect(chartCategoryVector(grid, 3)).toEqual(['Q1', 'Q2', 'Q3'])
  })

  it('keeps the innermost column for a column series', () => {
    const grid = [
      ['2024', 'Jan'],
      ['', 'Feb'],
      ['', 'Mar'],
    ]
    expect(chartCategoryVector(grid, 3)).toEqual(['Jan', 'Feb', 'Mar'])
  })

  it('preserves a cache that already held every cell', () => {
    const grid = [
      ['a', 'b'],
      ['c', 'd'],
    ]
    expect(chartCategoryVector(grid, 4)).toEqual(['a', 'b', 'c', 'd'])
  })

  it('falls back to the wider axis when the cache length matches neither', () => {
    const grid = [
      ['x', 'y', 'z'],
      ['1', '2', '3'],
    ]
    expect(chartCategoryVector(grid, 5)).toEqual(['1', '2', '3'])
  })

  it('uses the value orientation to disambiguate a square range', () => {
    const grid = [
      ['a', 'b'],
      ['c', 'd'],
    ]
    expect(chartCategoryVector(grid, 2, 'row')).toEqual(['c', 'd'])
    expect(chartCategoryVector(grid, 2, 'column')).toEqual(['b', 'd'])
  })
})

describe('seriesOrientation', () => {
  it('reads a single row or column, nothing for blocks', () => {
    expect(seriesOrientation({ startRow: 20, endRow: 20, startColumn: 1, endColumn: 13 })).toBe(
      'row',
    )
    expect(seriesOrientation({ startRow: 1, endRow: 9, startColumn: 3, endColumn: 3 })).toBe(
      'column',
    )
    expect(
      seriesOrientation({ startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 }),
    ).toBeUndefined()
    expect(
      seriesOrientation({ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }),
    ).toBeUndefined()
  })
})
