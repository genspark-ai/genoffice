import type { ChartDisplay } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import {
  CHART_TITLE_ROW_PX,
  chartTitleRowPx,
  drawChartSvg,
  renderChartSpec,
} from '../src/renderer/editor/protected-render'

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
    // the Beta point (null x) is dropped, not plotted at index 2 (series 0 marks points with diamonds)
    const markers = [...dom.querySelectorAll('polygon')]
    expect(markers.length).toBe(2)
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
    const markerX = (p: Element) => Number(p.getAttribute('points')!.split(' ')[0].split(',')[0])
    expect(Number(gamma.getAttribute('x'))).toBeCloseTo(markerX(markers[1]), 3)
    expect(Number(alpha.getAttribute('x'))).toBeCloseTo(markerX(markers[0]), 3)
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

  it('draws doughnut slices as rings and stacks the right legend vertically', () => {
    const dom = draw({
      partPath: 'word/charts/chart1.xml',
      kind: 'pie',
      holePct: 50,
      legendPos: 'r',
      categories: ['JSC', 'MX', 'ETC'],
      series: [{ values: [67, 32, 1] }],
    })
    const paths = [...dom.querySelectorAll('path')]
    expect(paths.length).toBe(3)
    // each slice carries a counter-sweep inner arc (the hole edge)
    for (const p of paths) {
      expect(p.getAttribute('d')).toMatch(/A [\d.e-]+ [\d.e-]+ 0 \d 0 /)
    }
    const legendTexts = [...dom.querySelectorAll('text')].filter((t) =>
      ['JSC', 'MX', 'ETC'].includes(t.textContent ?? ''),
    )
    expect(legendTexts.length).toBe(3)
    const xs = legendTexts.map((t) => Number(t.getAttribute('x')))
    const ys = legendTexts.map((t) => Number(t.getAttribute('y')))
    // one column on the right half, entries flowing downward
    expect(new Set(xs).size).toBe(1)
    expect(xs[0]).toBeGreaterThan(560 / 2)
    expect(ys[1]).toBeGreaterThan(ys[0])
    expect(ys[2]).toBeGreaterThan(ys[1])
    // the pie centers left of the legend gutter
    const svg = dom.querySelector('svg')!
    expect(svg.querySelectorAll('rect').length).toBe(3)
  })

  it('keeps solid pies and bottom legends unchanged without holePct/legendPos', () => {
    const dom = draw({
      partPath: 'word/charts/chart1.xml',
      kind: 'pie',
      categories: ['A', 'B'],
      series: [{ values: [3, 1] }],
    })
    const paths = [...dom.querySelectorAll('path')]
    expect(paths.length).toBe(2)
    for (const p of paths) expect(p.getAttribute('d')!.startsWith('M 280 ')).toBe(true)
    const legendTexts = [...dom.querySelectorAll('text')].filter((t) =>
      ['A', 'B'].includes(t.textContent ?? ''),
    )
    // bottom row: same y, different x
    expect(new Set(legendTexts.map((t) => t.getAttribute('y'))).size).toBe(1)
    expect(new Set(legendTexts.map((t) => t.getAttribute('x'))).size).toBe(2)
  })

  it('draws radar charts as rings, spokes and closed series outlines', () => {
    const dom = draw({
      partPath: 'word/charts/chart1.xml',
      kind: 'radar',
      radarStyle: 'standard',
      categories: ['A', 'B', 'C', 'D', 'E'],
      series: [
        { name: 'S1', values: [1, 2, 3, 4, 5] },
        { name: 'S2', values: [5, 4, 3, 2, 1] },
      ],
    })
    expect(dom.querySelectorAll('polyline').length).toBe(0)
    const polygons = [...dom.querySelectorAll('polygon')]
    const series = polygons.filter(
      (p) => p.getAttribute('fill') === 'none' && !p.classList.contains('doc-chart-grid'),
    )
    expect(series.length).toBe(2)
    for (const p of series) expect(p.getAttribute('points')!.split(' ').length).toBe(5)
    // value rings are polygons, one spoke per category
    expect(polygons.filter((p) => p.classList.contains('doc-chart-grid')).length).toBeGreaterThan(1)
    expect(dom.querySelectorAll('line.doc-chart-grid').length).toBe(5)
    const labels = [...dom.querySelectorAll('.doc-chart-axis-label')].map((t) => t.textContent)
    for (const c of ['A', 'B', 'C', 'D', 'E']) expect(labels).toContain(c)
  })

  it('fills filled-style radar series opaquely', () => {
    const dom = draw({
      partPath: 'word/charts/chart1.xml',
      kind: 'radar',
      radarStyle: 'filled',
      categories: ['A', 'B', 'C'],
      series: [{ name: 'S1', values: [1, 2, 3], color: 'C0504D' }],
    })
    const filled = [...dom.querySelectorAll('polygon')].filter(
      (p) => p.getAttribute('fill') === '#C0504D',
    )
    expect(filled.length).toBe(1)
    expect(filled[0].getAttribute('fill-opacity')).toBeNull()
  })

  it('explodes pie slices away from the center and labels them with percentages', () => {
    const base = {
      partPath: 'word/charts/chart1.xml',
      kind: 'pie' as const,
      categories: ['A', 'B', 'C'],
      series: [{ values: [1, 1, 2] }],
    }
    const solid = draw(base)
    const exploded = draw({ ...base, explosionPct: 25, dataLabels: { pct: true } })
    const apex = (dom: HTMLElement) =>
      [...dom.querySelectorAll('path')].map((p) =>
        p.getAttribute('d')!.split(' ').slice(1, 3).join(','),
      )
    // every slice starts from its own shifted center
    expect(new Set(apex(solid)).size).toBe(1)
    expect(new Set(apex(exploded)).size).toBe(3)
    const texts = [...exploded.querySelectorAll('text')].map((t) => t.textContent)
    expect(texts).toContain('25%')
    expect(texts).toContain('50%')
    expect([...solid.querySelectorAll('text')].map((t) => t.textContent)).not.toContain('50%')
  })

  it('stacks stacked line series on the running total', () => {
    const dom = draw({
      partPath: 'word/charts/chart1.xml',
      kind: 'line',
      grouping: 'stacked',
      categories: ['A', 'B'],
      series: [
        { name: 'S1', values: [2, 2] },
        { name: 'S2', values: [2, 2] },
      ],
    })
    const ys = [...dom.querySelectorAll('polyline')].map((l) =>
      l
        .getAttribute('points')!
        .split(' ')
        .map((p) => Number(p.split(',')[1])),
    )
    expect(ys.length).toBe(2)
    // the second series sits above the first by the same value step
    expect(ys[1][0]).toBeLessThan(ys[0][0])
    const labels = [...dom.querySelectorAll('.doc-chart-axis-label')].map((t) => t.textContent)
    expect(labels).toContain('4')
  })

  it('draws the data table under the plot and drops the axis category labels', () => {
    const dom = draw({
      partPath: 'word/charts/chart1.xml',
      kind: 'bar',
      noLegend: true,
      categories: ['apple', 'pear'],
      series: [
        { name: 'S1', values: [4.6, 3.9] },
        { name: 'S2', values: [2.4, 2.2] },
      ],
      dataTable: { keys: true, horz: true, vert: true, outline: true, line: '404040' },
    })
    const texts = [...dom.querySelectorAll('text')].map((t) => t.textContent)
    expect(texts.filter((t) => t === 'apple').length).toBe(1)
    expect(texts).toContain('4.6')
    expect(texts).toContain('2.2')
    const borders = [...dom.querySelectorAll('line')].filter(
      (l) => l.getAttribute('stroke') === '#404040',
    )
    expect(borders.length).toBeGreaterThan(4)
    // two key swatches, no legend row
    expect(dom.querySelectorAll('rect[width="8"]').length).toBe(2)
  })

  it('draws axis lines, titles and a top legend with line keys', () => {
    const dom = draw({
      partPath: 'word/charts/chart1.xml',
      kind: 'line',
      markers: true,
      legendPos: 't',
      xAxis: { line: '404040' },
      yAxis: { line: '404040', title: 'Axis Title' },
      categories: ['A', 'B'],
      series: [
        { name: 'S1', values: [1, 2] },
        { name: 'S2', values: [2, 1] },
      ],
    })
    const axisLines = [...dom.querySelectorAll('line')].filter(
      (l) => l.getAttribute('stroke') === '#404040',
    )
    expect(axisLines.length).toBe(2)
    const title = dom.querySelector('.doc-chart-axis-title')!
    expect(title.textContent).toBe('Axis Title')
    expect(title.getAttribute('transform')).toMatch(/^rotate\(-90/)
    const legend = [...dom.querySelectorAll('text')].filter(
      (t) => t.textContent === 'S1' || t.textContent === 'S2',
    )
    expect(legend.map((t) => t.getAttribute('y'))).toEqual(['12', '12'])
    // line series keys are line segments, not swatches; markers use the shape cycle
    expect(dom.querySelectorAll('rect[width="8"]').length).toBe(0)
    expect(dom.querySelectorAll('polygon').length).toBeGreaterThan(0)
  })

  it('lists stacked series top-first in a side legend and hides absent legends', () => {
    const base = {
      partPath: 'word/charts/chart1.xml',
      kind: 'bar' as const,
      grouping: 'stacked' as const,
      categories: ['A'],
      series: [
        { name: 'S1', values: [1] },
        { name: 'S2', values: [1] },
      ],
    }
    const side = draw({ ...base, legendPos: 'r' })
    const names = [...side.querySelectorAll('text')]
      .filter((t) => t.textContent === 'S1' || t.textContent === 'S2')
      .sort((a, b) => Number(a.getAttribute('y')) - Number(b.getAttribute('y')))
      .map((t) => t.textContent)
    expect(names).toEqual(['S2', 'S1'])
    const none = draw({ ...base, noLegend: true })
    expect([...none.querySelectorAll('text')].map((t) => t.textContent)).not.toContain('S1')
  })

  it('breaks radar outlines at gaps and pulls filled outlines to the center', () => {
    const base = {
      partPath: 'word/charts/chart1.xml',
      kind: 'radar' as const,
      categories: ['A', 'B', 'C', 'D'],
      series: [{ name: 'S1', values: [1, null, 3, 4] }],
    }
    const open = draw({ ...base, radarStyle: 'standard' })
    expect(open.querySelectorAll('polygon:not(.doc-chart-grid)').length).toBe(0)
    const run = open.querySelector('polyline')!
    expect(run.getAttribute('points')!.split(' ').length).toBe(3)
    const filled = draw({ ...base, radarStyle: 'filled' })
    const poly = filled.querySelector('polygon:not(.doc-chart-grid)')!
    expect(poly.getAttribute('points')!.split(' ').length).toBe(4)
  })

  it('uses the marker shape cycle on scatter points and skips tables on horizontal bars', () => {
    const scatter = draw({
      partPath: 'word/charts/chart1.xml',
      kind: 'scatter',
      markers: true,
      categories: [],
      series: [{ name: 'S1', values: [1, 2], xValues: [1, 2], line: true }],
    })
    expect(scatter.querySelectorAll('circle').length).toBe(0)
    expect(scatter.querySelectorAll('polygon').length).toBe(3) // 2 points + legend key
    const bars = draw({
      partPath: 'word/charts/chart1.xml',
      kind: 'bar',
      horizontal: true,
      categories: ['A', 'B'],
      series: [{ name: 'S1', values: [1, 2] }],
      dataTable: { keys: true, horz: true, vert: true, outline: true, line: '404040' },
    })
    expect(
      [...bars.querySelectorAll('line')].some((l) => l.getAttribute('stroke') === '#404040'),
    ).toBe(false)
  })
})

describe('chart text sizes, value labels and legend rows', () => {
  const bar: ChartDisplay = {
    partPath: 'word/charts/chart2.xml',
    kind: 'bar',
    title: 'Revenue',
    titleFontPt: 18,
    noLegend: true,
    widthPx: 576,
    heightPx: 288,
    categories: ['Base', 'A', 'B', 'C'],
    series: [{ name: 'Revenue', values: [1997, 2213, 2501, 2717] }],
    dataLabels: { val: true, fontPt: 12, color: '000000', numFmt: '#,##0' },
    xAxis: { fontPt: 12, color: '000000', line: '888888' },
    yAxis: { fontPt: 12, color: '000000', line: '888888', gridLine: '888888' },
  }

  it('draws c:showVal bar labels above the bars in the c:numFmt with the label size and color', () => {
    const dom = draw(bar)
    const texts = [...dom.querySelectorAll('text')]
    const label = texts.find((t) => t.textContent === '2,717')!
    expect(label).toBeDefined()
    expect(label.getAttribute('style')).toBe('font-size:16px;fill:#000000')
    const rects = [...dom.querySelectorAll('rect')]
    const top = Math.min(...rects.map((r) => Number(r.getAttribute('y'))))
    expect(Number(label.getAttribute('y'))).toBeLessThan(top)
    expect(texts.every((t) => t.textContent !== '2717')).toBe(true)
  })

  it('sizes and colors axis labels from c:txPr and strokes gridlines with the authored color', () => {
    const dom = draw(bar)
    const cat = [...dom.querySelectorAll('text')].find((t) => t.textContent === 'Base')!
    expect(cat.getAttribute('style')).toBe('font-size:16px;fill:#000000')
    const grid = dom.querySelectorAll('.doc-chart-grid')
    expect(grid.length).toBeGreaterThan(0)
    expect(grid[0].getAttribute('style')).toBe('stroke:#888888')
  })

  it('draws no gridlines when the parsed value axis has no c:majorGridlines', () => {
    const dom = draw({ ...bar, yAxis: { line: '888888' } })
    expect(dom.querySelectorAll('.doc-chart-grid').length).toBe(0)
    // models without a parsed axis keep the default grid
    const legacy = draw({ ...bar, xAxis: undefined, yAxis: undefined })
    expect(legacy.querySelectorAll('.doc-chart-grid').length).toBeGreaterThan(0)
  })

  it('budgets the title row from the title size, wrapping long titles', () => {
    expect(chartTitleRowPx({ ...bar, title: undefined })).toBe(0)
    expect(chartTitleRowPx({ ...bar, titleFontPt: undefined })).toBe(CHART_TITLE_ROW_PX)
    const oneLine = chartTitleRowPx(bar)
    expect(oneLine).toBeGreaterThan(CHART_TITLE_ROW_PX)
    const long = {
      ...bar,
      title: String.fromCodePoint(
        0x5e74,
        0x9593,
        0x58f2,
        0x4e0a,
        0x69cb,
        0x6210,
        0x5185,
        0x8a33,
      ).repeat(3),
    }
    expect(chartTitleRowPx(long)).toBeGreaterThan(oneLine * 1.8)
    const spec = renderChartSpec(bar)
    expect(JSON.stringify(spec)).toContain(`font-size:18pt;height:${oneLine}px`)
    // title row + plot still fill heightPx
    const svg = draw(bar).querySelector('svg')!
    expect(Number(svg.style.height.replace('px', ''))).toBe(288 - oneLine)
  })

  it('wraps an overflowing bottom legend into centered rows', () => {
    const pie: ChartDisplay = {
      partPath: 'word/charts/chart1.xml',
      kind: 'pie',
      holePct: 50,
      legendPos: 'b',
      widthPx: 400,
      heightPx: 240,
      categories: [
        'Pyrolysis oil ($972k)',
        'Tipping fee ($720k)',
        'Carbon ($226k)',
        'Steel ($77k)',
      ],
      series: [{ values: [972, 720, 226, 77] }],
    }
    const dom = draw(pie)
    const entries = pie.categories.map((c) =>
      [...dom.querySelectorAll('text')].find((t) => t.textContent === c)!,
    )
    const rows = new Set(entries.map((t) => t.getAttribute('y')))
    expect(rows.size).toBe(2)
    expect(entries.every((t) => t.getAttribute('text-anchor') === null)).toBe(true)
    const wide = draw({ ...pie, widthPx: 660 })
    const wideRows = new Set(
      pie.categories.map((c) =>
        [...wide.querySelectorAll('text')].find((t) => t.textContent === c)!.getAttribute('y'),
      ),
    )
    expect(wideRows.size).toBe(1)
  })
})

describe('category label wrapping', () => {
  it('wraps overflowing column-chart category labels into slot-width lines', () => {
    const dom = draw({
      partPath: 'word/charts/chart2.xml',
      kind: 'bar',
      noLegend: true,
      widthPx: 480,
      heightPx: 200,
      categories: ['Base case (no credit)', 'Scenario A ($15/t)', 'B'],
      series: [{ values: [1, 2, 3] }],
      xAxis: { fontPt: 12 },
      yAxis: { gridLine: 'D9D9D9' },
    })
    const texts = [...dom.querySelectorAll('text')].map((t) => t.textContent)
    expect(texts).toContain('Base case')
    expect(texts).toContain('(no credit)')
    expect(texts).toContain('B')
    expect(texts).not.toContain('Base case (no credit)')
  })

  it('caps a wrapped category label at three lines', () => {
    const dom = draw({
      partPath: 'word/charts/chart2.xml',
      kind: 'bar',
      noLegend: true,
      widthPx: 240,
      heightPx: 200,
      categories: ['one two three four five six seven eight nine ten eleven twelve', 'B'],
      series: [{ values: [1, 2] }],
      xAxis: { fontPt: 12 },
    })
    const texts = [...dom.querySelectorAll('text')].map((t) => t.textContent ?? '')
    const labelLines = texts.filter((t) => /[a-z]/.test(t) && t !== 'B')
    expect(labelLines.length).toBeLessThanOrEqual(3)
    expect(labelLines.join(' ')).toContain('twelve')
  })
})
