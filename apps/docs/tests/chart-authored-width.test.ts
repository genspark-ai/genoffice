import { describe, expect, it } from 'vitest'
import type { ChartDisplay } from '@genoffice/docx-engine'
import { CHART_MAX_WIDTH_PX, chartPlotWidth } from '../src/renderer/editor/protected-render'

describe('chart display width (issue #2805)', () => {
  it('an authored extent is honored instead of capped at the 660px render cap', () => {
    expect(chartPlotWidth({ widthPx: 768 } as ChartDisplay)).toBe(768)
    expect(chartPlotWidth({ widthPx: 300 } as ChartDisplay)).toBe(300)
  })

  it('charts without an explicit width keep the bounded default', () => {
    const width = chartPlotWidth({} as ChartDisplay)
    expect(width).toBeLessThanOrEqual(CHART_MAX_WIDTH_PX)
  })
})
