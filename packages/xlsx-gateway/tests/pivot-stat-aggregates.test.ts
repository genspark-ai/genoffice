/**
 * The four statistical aggregations (stdDev/stdDevp/var/varp) across the bake
 * path. The refresh engine (pivot-engine aggregate) carries the same arithmetic;
 * both switches must stay in sync.
 */
import { describe, expect, it } from 'vitest'
import { buildPivotLayout } from '../src/domain/pivot-layout'
import type { PivotAggregation } from '../src/domain/pivot-value-modes'

/** The grand-total cell of a one-row-field pivot over the given amounts */
function totalFor(agg: PivotAggregation, amounts: (number | string)[] = [10, 20, 30]): unknown {
  const grid: (string | number)[][] = [['Region', 'Amount']]
  amounts.forEach((v, i) => grid.push(['R' + i, v]))
  const layout = buildPivotLayout(grid, {
    rowFields: ['Region'],
    values: [{ field: 'Amount', agg }],
  })
  return layout.matrix[layout.matrix.length - 1]?.[1]
}

describe('statistical pivot aggregations', () => {
  it('sample variance and std dev use n-1 (STDEV/VAR)', () => {
    // {10, 20, 30}: mean 20, sum of squares 200, / (n-1) = 100, sqrt = 10
    expect(totalFor('var')).toBeCloseTo(100, 10)
    expect(totalFor('stdDev')).toBeCloseTo(10, 10)
  })

  it('population variance and std dev use n (VARP/STDEVP)', () => {
    expect(totalFor('varp')).toBeCloseTo(200 / 3, 10)
    expect(totalFor('stdDevp')).toBeCloseTo(Math.sqrt(200 / 3), 10)
  })

  it('a single-value bucket is blank for the sample statistics (Excel: #DIV/0!)', () => {
    const layout = buildPivotLayout(
      [
        ['Region', 'Amount'],
        ['East', 7],
      ],
      { rowFields: ['Region'], values: [{ field: 'Amount', agg: 'stdDev' }] },
    )
    // header row + one region row + grand total; the region cell is [1][1]
    expect(layout.matrix[1]?.[1]).toBeFalsy()
  })

  it('ignores non-numeric cells like every other aggregation', () => {
    // {10, 'n/a', 20}: the text cell drops out; sample var of {10,20} = 50
    expect(totalFor('var', [10, 'n/a', 20])).toBeCloseTo(50, 10)
  })
})
