import { describe, expect, it } from 'vitest'

import {
  comboBarIndices,
  comboCategorySeries,
  comboLineIndices,
} from '../src/renderer/WorkbookVisuals'

describe('comboLineIndices', () => {
  it('draws every lineChart-tagged series as a line regardless of position', () => {
    const series = [{ plot: 'barChart' }, { plot: 'lineChart' }, { plot: 'lineChart' }]
    expect([...comboLineIndices(series)]).toEqual([1, 2])
  })

  it('keeps bar series that follow the line group as bars', () => {
    const series = [{ plot: 'lineChart' }, { plot: 'barChart' }, { plot: 'barChart' }]
    expect([...comboLineIndices(series)]).toEqual([0])
  })

  it('falls back to the last series when no series carries a plot tag', () => {
    expect([...comboLineIndices([{}, {}, {}])]).toEqual([2])
    expect([...comboLineIndices([{}])]).toEqual([])
    expect([...comboLineIndices([])]).toEqual([])
  })
})

describe('comboBarIndices', () => {
  it('keeps the original positions of the bar series so clicks map back to chart.series', () => {
    const series = [{ plot: 'lineChart' }, { plot: 'barChart' }, { plot: 'barChart' }]
    expect(comboBarIndices(series, comboLineIndices(series))).toEqual([1, 2])
    const mixed = [{ plot: 'barChart' }, { plot: 'lineChart' }, { plot: 'barChart' }]
    expect(comboBarIndices(mixed, comboLineIndices(mixed))).toEqual([0, 2])
  })
})

describe('comboCategorySeries', () => {
  it('reads categories from the line series when only the first plot group caches them', () => {
    const line = { plot: 'lineChart', categories: ['Q1', 'Q2', 'Q3'], values: [1, 2, 3] }
    const bar = { plot: 'barChart', categories: [], values: [4, 5] }
    expect(comboCategorySeries([line, bar])).toBe(line)
    expect(comboCategorySeries([bar, line])).toBe(line)
  })

  it('falls back to the longest series when no series carries categories', () => {
    const short = { categories: [], values: [1] }
    const long = { categories: [], values: [1, 2, 3] }
    expect(comboCategorySeries([short, long])).toBe(long)
    expect(comboCategorySeries([])).toBeUndefined()
  })
})
