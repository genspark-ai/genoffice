import { describe, expect, it } from 'vitest'

import { recommendCharts } from '../src/domain/chart-recommend'

describe('recommendCharts', () => {
  it('puts line first for a time series', () => {
    const result = recommendCharts([
      ['Month', 'Sales'],
      ['Jan', 10],
      ['Feb', 12],
      ['Mar', 9],
      ['Apr', 15],
    ])
    expect(result?.items[0]).toEqual({ kind: 'line', reason: 'time' })
  })

  it('recommends pie for a few positive categories', () => {
    const result = recommendCharts([
      ['Region', 'Share'],
      ['North', 40],
      ['South', 25],
      ['East', 20],
      ['West', 15],
    ])
    expect(result?.items.map((item) => item.kind)).toContain('pie')
  })

  it('skips pie when values go negative', () => {
    const result = recommendCharts([
      ['Item', 'Delta'],
      ['A', 5],
      ['B', -3],
      ['C', 2],
    ])
    expect(result?.items.map((item) => item.kind)).not.toContain('pie')
  })

  it('puts line first for a Year/Value table with numeric years', () => {
    const result = recommendCharts([
      ['Year', 'Sales'],
      [2019, 10],
      [2020, 12],
      [2021, 9],
      [2022, 15],
      [2023, 18],
    ])
    expect(result?.items[0]).toEqual({ kind: 'line', reason: 'time' })
    expect(result?.parsed.hasCategoryColumn).toBe(true)
    expect(result?.parsed.categories).toEqual(['2019', '2020', '2021', '2022', '2023'])
  })

  it('puts scatter first for two bare numeric columns', () => {
    const rows: (string | number)[][] = [['X', 'Y']]
    for (let index = 0; index < 10; index++) rows.push([index * 2, index * index])
    const result = recommendCharts(rows)
    expect(result?.items[0]).toEqual({ kind: 'scatter', reason: 'correlation' })
  })

  it('prefers bar over column for long category labels', () => {
    const result = recommendCharts([
      ['Department', 'Headcount'],
      ['Customer Success and Support', 12],
      ['Engineering Infrastructure Team', 30],
      ['Product Design and Research Org', 9],
      ['International Sales Operations', 14],
      ['Legal and Regulatory Compliance', 4],
      ['Enterprise Marketing Division', 7],
      ['Talent Acquisition Partners', 5],
      ['Financial Planning and Analysis', 6],
      ['Corporate Development Office', 3],
    ])
    const kinds = result?.items.map((item) => item.kind) ?? []
    expect(kinds.indexOf('bar')).toBeGreaterThanOrEqual(0)
    expect(kinds.indexOf('bar')).toBeLessThan(kinds.indexOf('column'))
  })

  it('recommends combo when the two series scales differ wildly', () => {
    const result = recommendCharts([
      ['Month', 'Revenue', 'Margin %'],
      ['Jan', 120000, 12],
      ['Feb', 135000, 14],
      ['Mar', 110000, 9],
      ['Apr', 150000, 16],
    ])
    expect(result?.items.map((item) => item.kind)).toContain('combo')
  })

  it('caps recommendations at five and always offers a fallback', () => {
    const result = recommendCharts([
      ['Cat', 'V'],
      ['a', 1],
      ['b', 2],
    ])
    expect(result).not.toBeNull()
    expect(result!.items.length).toBeGreaterThan(0)
    expect(result!.items.length).toBeLessThanOrEqual(5)
    expect(result!.items.map((item) => item.kind)).toContain('column')
  })

  it('returns null for text-only selections', () => {
    expect(
      recommendCharts([
        ['a', 'b'],
        ['c', 'd'],
      ]),
    ).toBeNull()
  })
})
