import { describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'

import { BarChart, LineChart } from '../src/renderer/WorkbookVisuals'

const series = [
  { name: 'Actual', categories: ['Harvest', 'Sales'], values: [286.08, 243] },
  { name: 'Balance', categories: ['Harvest', 'Sales'], values: [1663.9, 1507] },
]
const axisLabels = (markup: string): string[] =>
  [...markup.matchAll(/class="axis-label"[^>]*>([^<]*)</g)].map((match) => match[1] ?? '')

describe('c:delete axes', () => {
  it('draws the column value axis unless the axis is deleted', () => {
    const shown = renderToStaticMarkup(
      createElement(BarChart, { seriesList: series, isHorizontal: false }),
    )
    expect(axisLabels(shown)).toContain('1,800')
    const hidden = renderToStaticMarkup(
      createElement(BarChart, {
        seriesList: series,
        isHorizontal: false,
        valueAxis: { hidden: true },
      }),
    )
    expect(axisLabels(hidden)).toEqual([])
    // Gridlines belong to the plot, not the scale labels — they survive.
    expect(hidden.match(/stroke="#e3e3e3"/g)?.length).toBe(shown.match(/stroke="#e3e3e3"/g)?.length)
    expect(hidden).toContain('Harvest')
  })

  it('drops the bottom scale of a horizontal bar chart when deleted', () => {
    const hidden = renderToStaticMarkup(
      createElement(BarChart, {
        seriesList: series,
        isHorizontal: true,
        valueAxis: { hidden: true },
      }),
    )
    expect(axisLabels(hidden)).toEqual([])
    expect(hidden).toContain('Harvest')
  })

  it('drops category labels when the category axis is deleted', () => {
    const hidden = renderToStaticMarkup(
      createElement(LineChart, { seriesList: series, categoryHidden: true }),
    )
    expect(hidden).not.toContain('Harvest')
    expect(axisLabels(hidden)).toContain('1,800')
  })
})

describe('c:delete category axis with multiLvlStrCache groups', () => {
  const grouped = [
    {
      name: 'Actual',
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      values: [10, 20, 30, 40],
      categoryGroups: [
        { label: 'H1', start: 0, end: 2 },
        { label: 'H2', start: 2, end: 4 },
      ],
    },
  ]
  const groupLabels = (markup: string): string[] =>
    ['H1', 'H2'].filter((label) => markup.includes(`>${label}<`))

  it.each([
    ['column', { isHorizontal: false }],
    ['horizontal bar', { isHorizontal: true }],
  ])('drops the %s group band when the category axis is deleted', (_kind, props) => {
    const shown = renderToStaticMarkup(createElement(BarChart, { seriesList: grouped, ...props }))
    expect(groupLabels(shown)).toEqual(['H1', 'H2'])
    const hidden = renderToStaticMarkup(
      createElement(BarChart, { seriesList: grouped, ...props, categoryHidden: true }),
    )
    expect(groupLabels(hidden)).toEqual([])
    expect(hidden).not.toContain('Q1')
    expect(hidden.match(/stroke="#e3e3e3"/g)?.length).toBe(shown.match(/stroke="#e3e3e3"/g)?.length)
  })

  it('drops the line chart group band when the category axis is deleted', () => {
    const shown = renderToStaticMarkup(createElement(LineChart, { seriesList: grouped }))
    expect(groupLabels(shown)).toEqual(['H1', 'H2'])
    const hidden = renderToStaticMarkup(
      createElement(LineChart, { seriesList: grouped, categoryHidden: true }),
    )
    expect(groupLabels(hidden)).toEqual([])
    expect(hidden).not.toContain('Q1')
  })
})
