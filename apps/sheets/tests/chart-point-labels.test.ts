import { describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'

import { BarChart } from '../src/renderer/WorkbookVisuals'

// Stacked column, first series labelled per point with manualLayout
// offsets, second series inheriting the plot element's showVal=0.
const actual = {
  name: 'Actual',
  categories: ['Harvest', 'Sales'],
  values: [286.08, 243],
  numberFormat: '#,##0.0',
  dataLabels: 'value' as const,
  pointLabels: [
    { index: 0, showVal: true, offsetX: -0.0035, offsetY: -0.3162 },
    { index: 1, showVal: true, offsetX: 0.0052, offsetY: -0.2544 },
  ],
}
const balance = {
  name: 'Balance',
  categories: ['Harvest', 'Sales'],
  values: [1663.9, 1507],
  numberFormat: '#,##0.0',
  dataLabels: 'none' as const,
}
const labels = (markup: string): Array<{ x: number; y: number; text: string }> =>
  [...markup.matchAll(/<text x="([-\d.]+)" y="([-\d.]+)"[^>]*class="axis-label"[^>]*>([^<]*)</g)]
    .map((match) => ({ x: Number(match[1]), y: Number(match[2]), text: match[3] ?? '' }))
    .filter((label) => label.text.includes('.'))

describe('per-point bar data labels', () => {
  it.each([
    ['column', false],
    ['horizontal bar', true],
  ])(
    'draws showVal points on a stacked %s chart and shifts them by manualLayout',
    (_kind, isHorizontal) => {
      const props = {
        seriesList: [actual, balance],
        isHorizontal,
        grouping: 'stacked' as const,
        dataLabels: 'value' as const,
        dataLabelPosition: 'center' as const,
        valueAxis: { hidden: true },
      }
      const shifted = labels(renderToStaticMarkup(createElement(BarChart, props)))
      expect(shifted.map((label) => label.text)).toEqual(['286.1', '243.0'])
      const unshifted = labels(
        renderToStaticMarkup(
          createElement(BarChart, {
            ...props,
            seriesList: [
              {
                ...actual,
                pointLabels: actual.pointLabels.map(({ index }) => ({ index, showVal: true })),
              },
              balance,
            ],
          }),
        ),
      )
      expect(unshifted.map((label) => label.text)).toEqual(['286.1', '243.0'])
      // Offsets are fractions of the 600x320 chart space.
      expect(shifted[0]!.x - unshifted[0]!.x).toBeCloseTo(-0.0035 * 600, 6)
      expect(shifted[0]!.y - unshifted[0]!.y).toBeCloseTo(-0.3162 * 320, 6)
      expect(shifted[1]!.y - unshifted[1]!.y).toBeCloseTo(-0.2544 * 320, 6)
      // Default stacked anchor: the middle of the segment, not over the bar.
      if (!isHorizontal) {
        const total = 286.08 + 1663.9
        expect(unshifted[0]!.y).toBeLessThan(283)
        expect(unshifted[0]!.y).toBeGreaterThan(283 - (286.08 / total) * 240)
      }
    },
  )

  it('hides a point whose dLbl says showVal=0 and honors the series mode', () => {
    const markup = renderToStaticMarkup(
      createElement(BarChart, {
        seriesList: [
          { ...actual, pointLabels: [{ index: 0, showVal: false }] },
          { ...balance, dataLabels: 'value' as const },
        ],
        isHorizontal: false,
        grouping: 'stacked' as const,
        dataLabels: 'value' as const,
        valueAxis: { hidden: true },
      }),
    )
    expect(labels(markup).map((label) => label.text)).toEqual(['1,663.9', '243.0', '1,507.0'])
  })

  it('lets a layout-only dLbl inherit the series mode and still shift the label', () => {
    const props = {
      seriesList: [
        { ...actual, pointLabels: [{ index: 0, offsetX: -0.0035, offsetY: -0.3162 }] },
        balance,
      ],
      isHorizontal: false,
      grouping: 'stacked' as const,
      dataLabels: 'value' as const,
      valueAxis: { hidden: true },
    }
    const shifted = labels(renderToStaticMarkup(createElement(BarChart, props)))
    expect(shifted.map((label) => label.text)).toEqual(['286.1', '243.0'])
    const unshifted = labels(
      renderToStaticMarkup(
        createElement(BarChart, {
          ...props,
          seriesList: [{ ...actual, pointLabels: [] }, balance],
        }),
      ),
    )
    expect(shifted[0]!.x - unshifted[0]!.x).toBeCloseTo(-0.0035 * 600, 6)
    expect(shifted[0]!.y - unshifted[0]!.y).toBeCloseTo(-0.3162 * 320, 6)
    // The series says none: the layout-only point stays hidden too.
    const hidden = renderToStaticMarkup(
      createElement(BarChart, {
        ...props,
        seriesList: [
          { ...actual, dataLabels: 'none' as const, pointLabels: [{ index: 0, offsetX: 0.01 }] },
          balance,
        ],
      }),
    )
    expect(labels(hidden)).toEqual([])
  })

  it('keeps clustered labels off when every series resolves to none', () => {
    const markup = renderToStaticMarkup(
      createElement(BarChart, {
        seriesList: [{ ...actual, dataLabels: 'none' as const, pointLabels: undefined }, balance],
        isHorizontal: false,
        dataLabels: 'value' as const,
        valueAxis: { hidden: true },
      }),
    )
    expect(labels(markup)).toEqual([])
  })
})
