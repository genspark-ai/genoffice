import { describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'

import {
  bottomAxisLayout,
  chartAreaInk,
  chartTextStyle,
  chartTextUnit,
  valueAxisLayout,
} from '../src/renderer/chart-text-scale'
import { BarChart, LineChart, ScatterChart } from '../src/renderer/WorkbookVisuals'

const axisLabels = (markup: string): string[] =>
  [...markup.matchAll(/class="axis-label"[^>]*>([^<]*)</g)].map((match) => match[1] ?? '')

describe('chart text sizing in points', () => {
  it('counter-scales the viewBox so 9pt stays 12px at zoom 1', () => {
    // A 600x320 frame draws 1:1 — one point is 4/3 units.
    expect(chartTextUnit({ width: 600, height: 320, zoom: 1 }, 600, 320)).toBeCloseTo(4 / 3, 3)
    // A wide-short dashboard frame (700x245) scales the viewBox by 245/320;
    // text units grow by the inverse so labels keep their point size.
    expect(chartTextUnit({ width: 700, height: 245, zoom: 1 }, 600, 320)).toBeCloseTo(
      4 / 3 / (245 / 320),
      3,
    )
    // Sheet zoom scales the frame and the text alike: the unit is unchanged.
    expect(chartTextUnit({ width: 1400, height: 490, zoom: 2 }, 600, 320)).toBeCloseTo(
      4 / 3 / (245 / 320),
      3,
    )
    expect(chartTextUnit(undefined, 600, 320)).toBe(1)
  })

  it('widens the viewBox to the left only when value labels overflow their gutter', () => {
    const box = { width: 600, height: 320, zoom: 1 }
    // "1,800" at 9pt: 5 chars · 0.54 em · 12 units ≈ 32 units — fits in 48.
    expect(valueAxisLayout(box, 320, 5 * 0.54, 9)).toEqual({ unit: 1.333, extra: 0 })
    // "50,000,000": 10 chars ≈ 65 units at 9pt — needs about 15 more.
    const wide = valueAxisLayout(box, 320, 10 * 0.54, 9)
    expect(wide.extra).toBeGreaterThan(10)
    expect(wide.extra).toBeLessThan(20)
    // Width-bound: the extension lowers the scale, so the unit grows a bit
    // and the extension is the fixed point of both.
    expect(wide.unit).toBeGreaterThan(1.333)
    expect(wide.unit * 10 * 0.54 * 9 - 48).toBeCloseTo(wide.extra, 0)
    // A rotated axis title plus a unit label reserve their own columns.
    const reserved = valueAxisLayout(box, 320, 4 * 0.54, 9, 1.15 * 10 + 1.1 * 9)
    expect(reserved.extra).toBeGreaterThan(0)
  })

  it('moves the bottom axis title below two-line ticks only when they would meet', () => {
    expect(bottomAxisLayout(false, true, false, 12, 13.3)).toEqual({
      bottomY: 317,
      viewBoxHeight: 320,
    })
    expect(bottomAxisLayout(true, false, false, 17, 19)).toEqual({
      bottomY: 317,
      viewBoxHeight: 320,
    })
    // 9pt ticks at scale 1 still clear the title.
    expect(bottomAxisLayout(true, true, false, 9, 10)).toEqual({ bottomY: 317, viewBoxHeight: 320 })
    const shifted = bottomAxisLayout(true, true, false, 17, 19)
    expect(shifted.bottomY).toBeGreaterThan(317)
    expect(shifted.viewBoxHeight).toBeGreaterThanOrEqual(Math.ceil(shifted.bottomY + 4))
  })

  it('emits the file font sizes and colors as custom properties', () => {
    const style = chartTextStyle(
      1.5,
      { labelSize: 9, labelColor: '#BFBFBF' },
      { labelSize: 11, titleSize: 10 },
    ) as Record<string, unknown>
    expect(style['--chart-pt']).toBe('1.5px')
    expect(style['--chart-cat-pt']).toBe(9)
    expect(style['--chart-cat-color']).toBe('#BFBFBF')
    expect(style['--chart-val-pt']).toBe(11)
    expect(style['--chart-title-pt']).toBe(10)
  })
})

const series = [{ name: 'Revenue', categories: ['Jun', 'Jul', 'Aug'], values: [12e6, 45e6, 30e6] }]

describe('c:dispUnits', () => {
  it('divides the value ticks and draws the unit label beside the axis title', () => {
    const markup = renderToStaticMarkup(
      createElement(BarChart, {
        seriesList: series,
        isHorizontal: false,
        axisTitles: { value: 'Amount' },
        valueAxis: { displayUnit: 1e6, displayUnitLabel: 'Millions' },
      }),
    )
    const labels = axisLabels(markup)
    expect(labels).toContain('50')
    expect(labels).not.toContain('50.0M')
    expect(labels).toContain('Millions')
    // The unit label sits to the right of the rotated axis title.
    expect(markup.indexOf('Amount')).toBeLessThan(markup.indexOf('Millions'))
  })

  it('scales the bottom axis of a horizontal bar chart too', () => {
    const markup = renderToStaticMarkup(
      createElement(BarChart, {
        seriesList: series,
        isHorizontal: true,
        valueAxis: { displayUnit: 1e6, displayUnitLabel: 'Millions' },
      }),
    )
    expect(axisLabels(markup)).toContain('50')
    expect(markup).toContain('Millions')
    const line = renderToStaticMarkup(
      createElement(LineChart, { seriesList: series, valueAxis: { displayUnit: 1e6 } }),
    )
    expect(axisLabels(line)).toContain('50')
    expect(line).not.toContain('Millions')
    // Combo right-hand scale.
    const combo = renderToStaticMarkup(
      createElement(BarChart, {
        seriesList: [{ name: 'Units', categories: ['Jun', 'Jul', 'Aug'], values: [1, 2, 3] }],
        lineSeriesList: series,
        isHorizontal: false,
        secondaryAxis: { displayUnit: 1e6 },
      }),
    )
    expect(axisLabels(combo)).toContain('50')
    expect(combo).not.toContain('50,000,000')
  })
})

describe('horizontal bar category labels', () => {
  const long = [
    {
      name: 'Total',
      categories: [
        'General & Admin Expenses',
        'Foreign Travel / Material Purchase',
        'Business and Promotion',
        'Staff Welfare',
      ],
      values: [4, 3, 2, 1],
    },
  ]

  it('keeps long category labels in full and shifts the plot right', () => {
    const markup = renderToStaticMarkup(
      createElement(BarChart, { seriesList: long, isHorizontal: true }),
    )
    expect(markup).toContain('General &amp; Admin Expenses')
    expect(markup).toContain('Business and Promotion')
    expect(markup).not.toContain('…')
    // The label column ends where the widest label needs it to, the bars
    // start 10 units further right.
    const labelX = Number(/<text x="([\d.]+)"[^>]*text-anchor="end"/.exec(markup)?.[1])
    expect(labelX).toBeGreaterThan(150)
    const barX = Number(/<rect[^>]*x="([\d.]+)"/.exec(markup)?.[1])
    expect(barX).toBeCloseTo(labelX + 10, 5)
  })

  it('wraps past a third of the chart before truncating', () => {
    const verbose = [
      {
        name: 'Total',
        categories: ['A very long category label that Excel would wrap onto two lines', 'Short'],
        values: [2, 1],
      },
    ]
    const markup = renderToStaticMarkup(
      createElement(BarChart, { seriesList: verbose, isHorizontal: true }),
    )
    const spans = markup.match(/<tspan/g)?.length ?? 0
    expect(spans).toBeGreaterThanOrEqual(3)
    expect(markup).toContain('A very long category label')
    // Short labels shrink the column so the plot gets the room.
    const compact = renderToStaticMarkup(
      createElement(BarChart, {
        seriesList: [{ name: 'T', categories: ['A', 'B'], values: [2, 1] }],
        isHorizontal: true,
      }),
    )
    const barX = Number(/<rect[^>]*x="([\d.]+)"/.exec(compact)?.[1])
    expect(barX).toBeLessThan(60)
  })
})

describe('chart and plot area fills', () => {
  it('paints the plot area fill behind the plot', () => {
    const markup = renderToStaticMarkup(
      createElement(BarChart, { seriesList: series, isHorizontal: false, plotAreaFill: '#ECECEC' }),
    )
    expect(markup).toContain('fill="#ECECEC"')
    const plain = renderToStaticMarkup(
      createElement(BarChart, { seriesList: series, isHorizontal: false }),
    )
    expect(plain).not.toContain('#ECECEC')
  })
})

describe('point-sized text and the viewBox', () => {
  // A wide-short dashboard frame: one point is 1.741 units, so 9pt text is
  // about 15.7 units tall and "50,000,000" no longer fits the 48-unit gutter.
  const wideShort = { width: 700, height: 245, zoom: 1 }
  const wide = series.map((entry) => ({ ...entry, numberFormat: '#,##0' }))
  const viewBoxLeft = (markup: string): number => Number(/viewBox="([-\d.]+) /.exec(markup)?.[1])

  it('reserves no gutter for a deleted value axis', () => {
    const shown = renderToStaticMarkup(
      createElement(BarChart, { seriesList: wide, isHorizontal: false, textBox: wideShort }),
    )
    expect(viewBoxLeft(shown)).toBeLessThan(0)
    expect(axisLabels(shown)).toContain('50,000,000')
    const hidden = renderToStaticMarkup(
      createElement(BarChart, {
        seriesList: wide,
        isHorizontal: false,
        textBox: wideShort,
        valueAxis: { hidden: true },
      }),
    )
    expect(viewBoxLeft(hidden)).toBe(0)
  })

  it('widens the scatter viewBox for wide point-sized Y ticks', () => {
    const scatter = wide.map((entry) => ({ ...entry, categories: ['1', '2', '3'] }))
    const fitted = renderToStaticMarkup(createElement(ScatterChart, { seriesList: scatter }))
    // Unmeasured: "45,000,000" at 1 unit/pt just about fits the gutter.
    expect(viewBoxLeft(fitted)).toBeGreaterThan(-1)
    const grown = renderToStaticMarkup(
      createElement(ScatterChart, { seriesList: scatter, textBox: wideShort }),
    )
    expect(viewBoxLeft(grown)).toBeLessThan(-30)
    expect(axisLabels(grown)).toContain('45,000,000')
    // Deleted axis: labels gone, no extension.
    const hidden = renderToStaticMarkup(
      createElement(ScatterChart, {
        seriesList: scatter,
        textBox: wideShort,
        valueAxis: { hidden: true },
      }),
    )
    expect(viewBoxLeft(hidden)).toBe(0)
  })

  it('budgets group labels with the point-sized glyph width', () => {
    const grouped = [
      {
        name: 'Actual',
        categories: ['Q1', 'Q2', 'Q3', 'Q4'],
        values: [10, 20, 30, 40],
        categoryGroups: [
          { label: 'First half of the fiscal year two thousand', start: 0, end: 2 },
          { label: 'H2', start: 2, end: 4 },
        ],
      },
    ]
    const roomy = renderToStaticMarkup(
      createElement(BarChart, { seriesList: grouped, isHorizontal: false }),
    )
    expect(roomy).toContain('>First half of the fiscal year two thousand<')
    const tight = renderToStaticMarkup(
      createElement(BarChart, { seriesList: grouped, isHorizontal: false, textBox: wideShort }),
    )
    // 240 units / (0.54 em · 9pt · 1.741) ≈ 28 characters.
    expect(tight).toMatch(/>First half of the fiscal yea[^<]*…</)
  })
})

describe('ink over the chart area fill', () => {
  it('uses the file text color, else light ink on a dark area', () => {
    expect(chartAreaInk('#3F3F3F', undefined)).toBe('#F2F2F2')
    expect(chartAreaInk('#303030', '#BFBFBF')).toBe('#BFBFBF')
    expect(chartAreaInk('#FFFFFF', undefined)).toBeUndefined()
    expect(chartAreaInk('#ECECEC', undefined)).toBeUndefined()
    expect(chartAreaInk(undefined, undefined)).toBeUndefined()
    expect(chartAreaInk('#1F4E79', undefined)).toBe('#F2F2F2')
  })

  it('separates data labels from the value-axis tick color', () => {
    const markup = renderToStaticMarkup(
      createElement(BarChart, {
        seriesList: series.map((entry) => ({ ...entry, numberFormat: '#,##0' })),
        isHorizontal: false,
        dataLabels: 'value',
        valueAxis: { labelColor: '#BFBFBF' },
      }),
    )
    // Tick labels keep the axis class; data labels never share it.
    const dataLabels = [...markup.matchAll(/class="data-label"[^>]*>([^<]*)</g)].map((m) => m[1])
    expect(dataLabels).toEqual(['12,000,000', '45,000,000', '30,000,000'])
    expect(axisLabels(markup)).toHaveLength(11)
  })
})
