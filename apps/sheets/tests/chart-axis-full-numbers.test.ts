import { describe, expect, it } from 'vitest'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'

import { BarChart, formatAxisValue } from '../src/renderer/WorkbookVisuals'

const axisLabels = (markup: string): string[] =>
  [...markup.matchAll(/class="axis-label"[^>]*>([^<]*)</g)].map((match) => match[1] ?? '')

describe('value axis numbers', () => {
  it('prints full numbers for large magnitudes like Excel and the labels do', () => {
    expect(formatAxisValue(1234567, undefined)).toBe('1,234,567')
    expect(formatAxisValue(1234567890, undefined)).toBe('1,234,567,890')
    // Small values keep their existing rendering.
    expect(formatAxisValue(1234, undefined)).toBe('1,234')
    expect(formatAxisValue(0, undefined)).toBe('0')
    expect(formatAxisValue(-1234567, undefined)).toBe('-1,234,567')
  })

  it('renders large axis ticks in full, not abbreviated', () => {
    const seriesList = [{ name: 'Users', categories: ['Q1', 'Q2'], values: [1234567, 1234567890] }]
    const markup = renderToStaticMarkup(
      createElement(BarChart, { seriesList, isHorizontal: false, dataLabels: 'value' }),
    )
    const ticks = axisLabels(markup)
    expect(ticks).toContain('200,000,000')
    expect(ticks.some((tick) => /[0-9]M$|[0-9]B$/.test(tick))).toBe(false)
    // The data labels of the same series already print the full numbers.
    const dataLabels = [...markup.matchAll(/class="data-label"[^>]*>([^<]*)</g)].map(
      (match) => match[1] ?? '',
    )
    expect(dataLabels).toEqual(expect.arrayContaining(['1,234,567', '1,234,567,890']))
  })
})
