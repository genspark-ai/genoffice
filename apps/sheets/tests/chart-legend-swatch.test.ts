import { describe, expect, it } from 'vitest'

import {
  legendSwatchColor,
  legendUsesLineSwatches,
  lineStroke,
} from '../src/renderer/WorkbookVisuals'

describe('lineStroke', () => {
  it('prefers the explicit series line color over the fill/accent default', () => {
    expect(lineStroke({ color: '#9BBB59', lineColor: '#10253F' }, 0)).toBe('#10253F')
    expect(lineStroke({ color: '#9BBB59' }, 0)).toBe('#9BBB59')
  })

  it('returns null for an explicit noFill line and falls back to the palette', () => {
    expect(lineStroke({ color: '#9BBB59', lineColor: 'none' }, 0)).toBeNull()
    expect(lineStroke({}, 0)).toBe('#4472c4')
  })
})

describe('legendSwatchColor', () => {
  const series = { color: '#9BBB59', lineColor: '#10253F' }

  it('mirrors the drawn stroke on line-family charts', () => {
    expect(legendSwatchColor(series, 0, true)).toBe('#10253F')
    // noFill line: the swatch keeps the series color instead of vanishing.
    expect(legendSwatchColor({ color: '#9BBB59', lineColor: 'none' }, 0, true)).toBe('#9BBB59')
    expect(legendSwatchColor({}, 1, true)).toBe('#ed7d31')
  })

  it('keeps the fill resolution for non-line charts', () => {
    expect(legendSwatchColor(series, 0, false)).toBe('#9BBB59')
    expect(legendSwatchColor({}, 1, false)).toBe('#ed7d31')
  })
})

describe('legendUsesLineSwatches', () => {
  it('follows the render dispatch: only a winning line/radar plot strokes', () => {
    expect(legendUsesLineSwatches(['lineChart'])).toBe(true)
    expect(legendUsesLineSwatches(['radarChart'])).toBe(true)
    // Area wins the cascade over line: the drawn series are fills, so the
    // swatches must keep the fill resolution.
    expect(legendUsesLineSwatches(['lineChart', 'areaChart'])).toBe(false)
    expect(legendUsesLineSwatches(['lineChart', 'scatterChart'])).toBe(false)
    // Bar combos render BarChart; pie keeps its own legend.
    expect(legendUsesLineSwatches(['barChart', 'lineChart'])).toBe(false)
    expect(legendUsesLineSwatches(['radarChart', 'barChart'])).toBe(false)
    expect(legendUsesLineSwatches(['pieChart', 'lineChart'])).toBe(false)
    expect(legendUsesLineSwatches(['barChart'])).toBe(false)
  })
})
