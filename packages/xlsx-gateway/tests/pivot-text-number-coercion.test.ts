import { describe, expect, it } from 'vitest'

import { buildPivotLayout } from '../src/domain/pivot-layout'

/// Amount is text-formatted: the cells hold the characters "100" and "50",
/// which is how a spreadsheet returns a number stored as text.
const TEXT_NUMBER_GRID = [
  ['Region', 'Amount'],
  ['East', '100'],
  ['West', '50'],
]

function totalFor(
  grid: (string | number | null)[][],
  agg: 'sum' | 'average' | 'max' | 'min' = 'sum',
): unknown {
  const layout = buildPivotLayout(grid, {
    rowFields: ['Region'],
    values: [{ field: 'Amount', agg }],
  })
  // Row layout: header, one line per region, then the grand total.
  return layout.matrix[layout.matrix.length - 1]?.[1]
}

describe('pivot layout number coercion', () => {
  it('bakes text-formatted numbers the way the refresh engine aggregates them', () => {
    // pivot-engine coerces with Number(String(value).trim()), so a refresh
    // totals 150 here. The baked grid dropped both cells instead and left the
    // value blank, so opening the file disagreed with refreshing it.
    expect(totalFor(TEXT_NUMBER_GRID)).toBe(150)
  })

  it('coerces a mix of stored numbers and text, ignoring unparseable cells', () => {
    // Same filter as pivot-engine: a raw value counts when it is not
    // null/'' and the coerced number is finite.
    expect(
      totalFor([
        ['Region', 'Amount'],
        ['East', 100],
        ['East', ' 50 '],
        ['East', 'n/a'],
        ['East', ''],
        ['East', null],
      ]),
    ).toBe(150)
  })

  it('keeps the numeric aggregates on text-formatted numbers', () => {
    expect(totalFor(TEXT_NUMBER_GRID, 'max')).toBe(100)
    expect(totalFor(TEXT_NUMBER_GRID, 'min')).toBe(50)
    expect(totalFor(TEXT_NUMBER_GRID, 'average')).toBe(75)
  })
})
