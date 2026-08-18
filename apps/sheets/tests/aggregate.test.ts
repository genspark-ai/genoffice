import { describe, expect, it } from 'vitest'
import { createRangeAggregator, formatRangeAggregate } from '../src/renderer/ai/aggregate'

describe('createRangeAggregator', () => {
  it('computes counts, distinct values, and numeric stats', () => {
    const aggregator = createRangeAggregator()
    for (const value of ['供应商A', '供应商B', '供应商A', 10, 20, 10, null, '']) {
      aggregator.add(value)
    }
    aggregator.addEmpty(2)
    const result = aggregator.finish(10)
    expect(result.cells).toBe(10)
    expect(result.nonEmpty).toBe(6)
    expect(result.distinct).toBe(4) // two CJK supplier strings plus 10 and 20
    expect(result.numericCount).toBe(3)
    expect(result.sum).toBe(40)
    expect(result.min).toBe(10)
    expect(result.max).toBe(20)
    expect(result.average).toBeCloseTo(40 / 3)
  })

  it('ranks top values by frequency', () => {
    const aggregator = createRangeAggregator()
    for (const value of ['a', 'b', 'a', 'c', 'a', 'b']) aggregator.add(value)
    const result = aggregator.finish(2)
    expect(result.topValues).toEqual([
      { value: 'a', count: 3 },
      { value: 'b', count: 2 },
    ])
  })

  it('keeps exact counts after distinct tracking overflows', () => {
    const aggregator = createRangeAggregator()
    for (let index = 0; index < 200_001; index += 1) aggregator.add(`v${index}`)
    const result = aggregator.finish(10)
    expect(result.distinct).toBeNull()
    expect(result.topValues).toEqual([])
    expect(result.nonEmpty).toBe(200_001)
  })
})

describe('formatRangeAggregate', () => {
  it('renders a compact report and slices top values', () => {
    const aggregator = createRangeAggregator()
    for (const value of ['a', 'a', 'b', 5]) aggregator.add(value)
    const text = formatRangeAggregate('D2:D5', aggregator.finish(10), 1)
    expect(text).toContain('Statistics for D2:D5')
    expect(text).toContain('non-empty: 4')
    expect(text).toContain('distinct values: 3')
    expect(text).toContain('sum: 5')
    expect(text).toContain('a: 2')
    expect(text).not.toContain('b: 1')
  })
})
