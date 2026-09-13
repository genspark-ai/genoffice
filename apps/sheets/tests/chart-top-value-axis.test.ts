import { describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'

import { BarChart } from '../src/renderer/WorkbookVisuals'

const series = [
  { name: 'Days', categories: ['Okanogan', 'Thurston', 'Lewis'], values: [70.5, 67.9, 64.4] },
]
const axisLabelYs = (markup: string): number[] =>
  [...markup.matchAll(/<text x="[^"]*" y="([^"]*)" text-anchor="middle" class="axis-label"/g)].map(
    (match) => Number(match[1]),
  )
const barYs = (markup: string): number[] =>
  [...markup.matchAll(/<rect [^>]*y="([^"]*)" width="[^"]*" height="[^"]*" fill="#/g)].map(
    (match) => Number(match[1]),
  )

describe('horizontal bar chart value axis side (c:valAx/c:axPos)', () => {
  it('draws the scale along the bottom by default', () => {
    const markup = renderToStaticMarkup(
      createElement(BarChart, { seriesList: series, isHorizontal: true }),
    )
    const ticks = axisLabelYs(markup)
    expect(ticks.length).toBeGreaterThan(1)
    expect(Math.min(...ticks)).toBeGreaterThan(Math.max(...barYs(markup)))
    expect(markup).toContain('y="298"')
  })

  it('moves the scale, gridline start and plot down for axPos="t"', () => {
    const markup = renderToStaticMarkup(
      createElement(BarChart, {
        seriesList: series,
        isHorizontal: true,
        valueAxis: { position: 't' },
        axisTitles: { category: null, value: 'Days per patient' },
      }),
    )
    const ticks = axisLabelYs(markup)
    const bars = barYs(markup)
    expect(ticks.length).toBeGreaterThan(1)
    expect(Math.max(...ticks)).toBeLessThan(Math.min(...bars))
    expect(markup).not.toContain('y="298"')
    // The value title rides above the ticks instead of the bottom edge.
    expect(markup).toMatch(/y="12" text-anchor="middle" class="axis-title">Days per patient</)
    expect(markup).not.toContain('y="317"')
    expect(markup).toContain('Okanogan')
  })
})
