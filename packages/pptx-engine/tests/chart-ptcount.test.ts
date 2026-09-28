import { performance } from 'node:perf_hooks'
import { describe, expect, it } from 'vitest'
import { parseChartXml } from '../src/chart'

const MB = 1024 * 1024

function chartXml(sers: number, declared: number, realPts = 1): string {
  const ser = (i: number) =>
    `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` +
    `<c:val><c:numRef><c:f>Sheet1!$A$1</c:f><c:numCache><c:ptCount val="${declared}"/>` +
    Array.from({ length: realPts }, (_, k) => `<c:pt idx="${k}"><c:v>${k + 1}</c:v></c:pt>`).join(
      '',
    ) +
    `</c:numCache></c:numRef></c:val></c:ser>`
  return (
    `<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
    `<c:chart><c:plotArea><c:barChart><c:barDir val="col"/>` +
    Array.from({ length: sers }, (_, i) => ser(i)).join('') +
    `</c:barChart></c:plotArea></c:chart></c:chartSpace>`
  )
}

describe('chart point allocation follows the data, not ptCount', () => {
  it('allocates from the actual c:pt count even when ptCount claims a million', () => {
    // The #759 shape: one number used to buy a MAX_CHART_POINTS-sized array.
    const model = parseChartXml(chartXml(1, 1_048_576, 1), undefined) as any
    const series = model.plots?.[0]?.series ?? model.series
    // The declaration may only buy MAX_TAIL_PADDING slots past the real data.
    expect(series[0].values.length).toBeLessThanOrEqual(1025)
    expect(series[0].values[0]).toBe(1)
  })

  it('keeps honest files exactly as they were', () => {
    const model = parseChartXml(chartXml(2, 3, 3), undefined) as any
    const series = model.plots?.[0]?.series ?? model.series
    expect(series[0].values).toEqual([1, 2, 3])
    expect(series[1].values).toEqual([1, 2, 3])
  })

  it('caps a real 400-series hostile chart at negligible memory', () => {
    // 400 × 1,048,576 declared points held 2.3 GB of RSS on main; with
    // allocation following the data, one real point per series stays tiny.
    const before = process.memoryUsage().rss
    const t0 = performance.now()
    const model = parseChartXml(chartXml(400, 1_048_576, 1), undefined) as any
    const series = model.plots?.[0]?.series ?? model.series
    expect(series).toHaveLength(256) // MAX_CHART_SERIES still applies
    expect(series.every((s: { values: unknown[] }) => s.values.length <= 1025)).toBe(true)
    expect(performance.now() - t0).toBeLessThan(2_000)
    expect(process.memoryUsage().rss - before).toBeLessThan(64 * MB)
  })
})
