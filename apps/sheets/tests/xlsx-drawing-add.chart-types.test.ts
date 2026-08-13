import { describe, expect, it } from 'vitest'

import { buildChartXml, type ChartAdd } from '../src/gateway/xlsx-drawing-add'

const chart = (chartType: ChartAdd['chartType'], categories: readonly string[]): ChartAdd => ({
  chartType,
  title: 'Test Chart',
  series: [
    {
      name: 'Series 1',
      categories,
      values: [10, 20, 30],
    },
  ],
})

describe('buildChartXml chart types', () => {
  it('scatter plots xVal/yVal against two value axes and no category axis', () => {
    const xml = buildChartXml(chart('scatter', ['1', '2', '3']))
    expect(xml).toContain('<c:scatterChart>')
    expect(xml).toContain('<c:scatterStyle val="lineMarker"/>')
    expect(xml).toContain('</c:scatterChart>')
    expect(xml.match(/<c:valAx>/g)).toHaveLength(2)
    expect(xml).not.toContain('<c:catAx>')
    expect(xml).toContain('<c:xVal>')
    expect(xml).toContain('<c:yVal>')
    expect(xml).not.toContain('<c:cat>')
    // Numeric categories ride along as the X values.
    expect(xml).toContain('<c:xVal><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt>')
  })

  it('scatter falls back to ordinal X values when categories are not numeric', () => {
    const xml = buildChartXml(chart('scatter', ['a', 'b', 'c']))
    expect(xml).toContain(
      '<c:xVal><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>0</c:v></c:pt>' +
        '<c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt></c:numLit></c:xVal>',
    )
  })

  it('combo plots the last series as a line on a secondary right axis', () => {
    const xml = buildChartXml({
      chartType: 'combo',
      title: 'Combo',
      series: [
        { name: 'Sales', categories: ['a', 'b', 'c'], values: [10, 20, 30] },
        { name: 'Growth', categories: ['a', 'b', 'c'], values: [0.1, 0.2, 0.3] },
      ],
    })
    expect(xml).toContain('<c:barChart>')
    expect(xml).toContain('<c:lineChart>')
    // Bar series 0 in the bar plot, line series 1 in the line plot.
    expect(/<c:barChart>[\s\S]*?<\/c:barChart>/.exec(xml)?.[0]).toContain('<c:idx val="0"/>')
    expect(/<c:lineChart>[\s\S]*?<\/c:lineChart>/.exec(xml)?.[0]).toContain('<c:idx val="1"/>')
    // Primary cat+val pair, plus a hidden secondary cat axis and a right
    // value axis crossing at max.
    expect(xml.match(/<c:catAx>/g)).toHaveLength(2)
    expect(xml.match(/<c:valAx>/g)).toHaveLength(2)
    expect(xml).toContain('<c:axPos val="r"/>')
    expect(xml).toContain('<c:crosses val="max"/>')
    expect(xml).toContain('<c:delete val="1"/>')
  })

  it('combo with one series degrades to plain columns on the primary axes', () => {
    const xml = buildChartXml(chart('combo', ['a', 'b', 'c']))
    expect(xml).toContain('<c:barChart>')
    expect(xml).not.toContain('<c:lineChart>')
    expect(xml.match(/<c:catAx>/g)).toHaveLength(1)
    expect(xml.match(/<c:valAx>/g)).toHaveLength(1)
  })

  it('radar keeps the category + value axis pair', () => {
    const xml = buildChartXml(chart('radar', ['a', 'b', 'c']))
    expect(xml).toContain('<c:radarChart>')
    expect(xml).toContain('<c:radarStyle val="marker"/>')
    expect(xml).toContain('</c:radarChart>')
    expect(xml.match(/<c:catAx>/g)).toHaveLength(1)
    expect(xml.match(/<c:valAx>/g)).toHaveLength(1)
    expect(xml).toContain('<c:cat>')
    expect(xml).toContain('<c:val>')
  })

  it('doughnut is axis-free with a hole, like pie', () => {
    const xml = buildChartXml(chart('doughnut', ['a', 'b', 'c']))
    expect(xml).toContain('<c:doughnutChart>')
    expect(xml).toContain('<c:varyColors val="1"/>')
    expect(xml).toContain('<c:holeSize val="50"/>')
    expect(xml).toContain('</c:doughnutChart>')
    expect(xml).not.toContain('<c:catAx>')
    expect(xml).not.toContain('<c:valAx>')
  })
})
