import { describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'

import { BarChart } from '../src/renderer/WorkbookVisuals'

const barRects = (markup: string): Array<{ x: number; y: number; width: number; height: number }> =>
  [...markup.matchAll(/<rect [^>]*?x="([^"]+)" y="([^"]+)" width="([^"]+)" height="([^"]+)"/g)]
    .map(([, x, y, width, height]) => ({
      x: Number(x),
      y: Number(y),
      width: Number(width),
      height: Number(height),
    }))
    .filter((rect) => rect.height > 0 && rect.height < 240)

describe('bars with a below-zero value axis', () => {
  it('grows columns from the zero line in both directions', () => {
    const seriesList = [{ name: 'Delta', categories: ['a', 'b'], values: [300, -500] }]
    const markup = renderToStaticMarkup(
      createElement(BarChart, { seriesList, isHorizontal: false, dataLabels: 'none' }),
    )
    // Axis -600..400: zero sits at 60% of the plot height.
    const zeroY = 280 - 0.6 * 240
    const [positive, negative] = barRects(markup)
    expect(positive).toBeDefined()
    expect(negative).toBeDefined()
    expect(positive!.y + positive!.height).toBeCloseTo(zeroY, 5)
    expect(positive!.height).toBeCloseTo(0.3 * 240, 5)
    expect(negative!.y).toBeCloseTo(zeroY, 5)
    expect(negative!.height).toBeCloseTo(0.5 * 240, 5)
  })

  it('grows horizontal bars from the zero line', () => {
    const seriesList = [{ name: 'Delta', categories: ['a', 'b'], values: [300, -500] }]
    const markup = renderToStaticMarkup(
      createElement(BarChart, { seriesList, isHorizontal: true, dataLabels: 'none' }),
    )
    const [positive, negative] = barRects(markup)
    expect(positive).toBeDefined()
    expect(negative).toBeDefined()
    expect(negative!.x + negative!.width).toBeCloseTo(positive!.x, 5)
    expect(negative!.width / positive!.width).toBeCloseTo(5 / 3, 5)
  })

  it('draws a missing value as a zero-height bar, not a full-depth negative one', () => {
    const seriesList = [
      { name: 'A', categories: ['a', 'b'], values: [300, -500] },
      { name: 'B', categories: ['a', 'b'], values: [100] },
    ]
    const markup = renderToStaticMarkup(
      createElement(BarChart, { seriesList, isHorizontal: false, dataLabels: 'none' }),
    )
    const heights = [...markup.matchAll(/<rect [^>]*?height="([^"]+)"/g)].map(([, h]) => Number(h))
    expect(heights.filter((h) => h === 0.5 * 240)).toHaveLength(1)
  })

  it('keeps all-positive columns on the plot floor', () => {
    const seriesList = [{ name: 'Sales', categories: ['a'], values: [100] }]
    const markup = renderToStaticMarkup(
      createElement(BarChart, { seriesList, isHorizontal: false, dataLabels: 'none' }),
    )
    const [bar] = barRects(markup)
    expect(bar!.y + bar!.height).toBeCloseTo(280, 5)
  })
})
