import type { ChartDisplay } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { drawChartSvg } from '../src/renderer/editor/protected-render'

function draw(chart: ChartDisplay): HTMLElement {
  const dom = document.createElement('div')
  const canvas = document.createElement('div')
  canvas.className = 'doc-chart-canvas'
  dom.appendChild(canvas)
  drawChartSvg(dom, chart)
  return dom
}

describe('chart svg drawing', () => {
  it('skips scatter points with a null cached x and keeps labels on their own points', () => {
    const dom = draw({
      partPath: 'word/charts/chart1.xml',
      kind: 'scatter',
      markers: true,
      categories: ['Alpha', 'Beta', 'Gamma'],
      series: [{ name: 'S1', values: [10, 20, 30], xValues: [1, null, 3] }],
    })
    // the Beta point (null x) is dropped, not plotted at index 2
    expect(dom.querySelectorAll('circle').length).toBe(2)
    const labels = [...dom.querySelectorAll('.doc-chart-axis-label')].map((t) => t.textContent)
    expect(labels).toContain('Alpha')
    expect(labels).toContain('Gamma')
    expect(labels).not.toContain('Beta')
    // Gamma anchors to the x=3 point (plot right edge), not to the second survivor slot
    const gamma = [...dom.querySelectorAll('.doc-chart-axis-label')].find(
      (t) => t.textContent === 'Gamma',
    )!
    const alpha = [...dom.querySelectorAll('.doc-chart-axis-label')].find(
      (t) => t.textContent === 'Alpha',
    )!
    const circles = [...dom.querySelectorAll('circle')]
    expect(Number(gamma.getAttribute('x'))).toBeCloseTo(Number(circles[1].getAttribute('cx')), 3)
    expect(Number(alpha.getAttribute('x'))).toBeCloseTo(Number(circles[0].getAttribute('cx')), 3)
  })

  it('stacks percentStacked area charts and normalizes the axis to 0-100%', () => {
    const dom = draw({
      partPath: 'word/charts/chart1.xml',
      kind: 'area',
      grouping: 'percentStacked',
      categories: ['A', 'B'],
      series: [
        { name: 'S1', values: [1, 3] },
        { name: 'S2', values: [3, 1] },
      ],
    })
    const labels = [...dom.querySelectorAll('.doc-chart-axis-label')].map((t) => t.textContent)
    expect(labels).toContain('100%')
    expect(labels).toContain('0%')
    const polygons = [...dom.querySelectorAll('polygon')]
    expect(polygons.length).toBe(2)
    // the top series' upper edge sits on the 100% line at both categories
    const topLine = dom.querySelectorAll('polyline')[1]
    const ys = topLine
      .getAttribute('points')!
      .split(' ')
      .map((p) => Number(p.split(',')[1]))
    expect(ys[0]).toBeCloseTo(ys[1], 3)
    const hundred = [...dom.querySelectorAll('.doc-chart-axis-label')].find(
      (t) => t.textContent === '100%',
    )!
    expect(ys[0]).toBeCloseTo(Number(hundred.getAttribute('y')) - 3, 3)
  })

  it('stacks plain stacked area charts cumulatively', () => {
    const dom = draw({
      partPath: 'word/charts/chart1.xml',
      kind: 'area',
      grouping: 'stacked',
      categories: ['A', 'B'],
      series: [
        { name: 'S1', values: [2, 2] },
        { name: 'S2', values: [2, 2] },
      ],
    })
    // axis labels carry no % suffix and reach the stacked total (4) or above
    const labels = [...dom.querySelectorAll('.doc-chart-axis-label')].map((t) => t.textContent)
    expect(labels.some((l) => l?.includes('%'))).toBe(false)
    expect(labels).toContain('4')
    expect(dom.querySelectorAll('polygon').length).toBe(2)
  })
})
