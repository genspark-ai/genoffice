/** A new chart of each type Insert > Chart offers reads back as that type. */
import { describe, expect, it } from 'vitest'
import { buildChartPartXml, parseChartPartXml, type NewChart } from '../src/index'

const DATA = {
  title: 'Sales',
  categories: ['Q1', 'Q2', 'Q3'],
  series: [
    { name: 'North', values: [4, 6, 5] },
    { name: 'South', values: [2, null, 7] },
  ],
}

const read = (chart: NewChart) => parseChartPartXml(buildChartPartXml(chart), 'word/charts/c.xml')

describe('buildChartPartXml chart types', () => {
  it.each(['bar', 'line', 'pie', 'area', 'radar'] as const)('%s reads back as itself', (kind) => {
    const display = read({ kind, ...DATA })
    expect(display?.kind).toBe(kind)
    expect(display?.categories).toEqual(DATA.categories)
    expect(display?.series.map((s) => s.name)).toEqual(['North', 'South'])
    expect(display?.series[1].values).toEqual([2, null, 7])
  })

  it('columns stand up and bars lie down, with the axes swapped', () => {
    expect(read({ kind: 'bar', ...DATA })?.horizontal).toBeFalsy()
    expect(read({ kind: 'bar', horizontal: true, ...DATA })?.horizontal).toBe(true)
    const sideways = buildChartPartXml({ kind: 'bar', horizontal: true, ...DATA })
    expect(sideways).toMatch(/<c:catAx>.*?<c:axPos val="l"\/>.*?<\/c:catAx>/)
    expect(sideways).toMatch(/<c:valAx>.*?<c:axPos val="b"\/>.*?<\/c:valAx>/)
    // the flag means nothing to the other types
    expect(buildChartPartXml({ kind: 'line', horizontal: true, ...DATA })).toContain(
      '<c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>',
    )
  })

  it('a scatter chart reads back with its x values, as points with no line between them', () => {
    const display = read({ kind: 'scatter', ...DATA, categories: ['1.5', '3', 'n/a'] })
    expect(display?.kind).toBe('scatter')
    expect(display?.series.map((s) => s.name)).toEqual(['North', 'South'])
    // a category that is not a number takes its position
    expect(display?.series[0].xValues).toEqual([1.5, 3, 3])
    expect(display?.series[1].values).toEqual([2, null, 7])
    expect(display?.series.some((s) => s.line)).toBe(false)
    const xml = buildChartPartXml({ kind: 'scatter', ...DATA })
    expect(xml.match(/<c:valAx>/g)).toHaveLength(2)
    expect(xml).not.toContain('<c:catAx>')
  })

  it('a stock chart reads back as one: no line along a series, a high-low line per category', () => {
    const stock: NewChart = {
      kind: 'stock',
      title: 'Price',
      categories: ['Mon', 'Tue'],
      series: [
        { name: 'High', values: [55, 57] },
        { name: 'Low', values: [11, 12] },
        { name: 'Close', values: [32, 35] },
      ],
    }
    const display = read(stock)
    expect(display?.kind).toBe('stock')
    expect(display?.categories).toEqual(['Mon', 'Tue'])
    expect(display?.series.map((s) => [s.name, s.values])).toEqual([
      ['High', [55, 57]],
      ['Low', [11, 12]],
      ['Close', [32, 35]],
    ])
    const xml = buildChartPartXml(stock)
    expect(xml).toContain('<c:hiLowLines/>')
    // every series has its line off; only the close carries a mark
    expect(xml.match(/<a:ln w="19050"><a:noFill\/><\/a:ln>/g)).toHaveLength(3)
    expect(xml.match(/<c:symbol val="none"\/>/g)).toHaveLength(2)
    expect(xml.match(/<c:symbol val="dash"\/>/g)).toHaveLength(1)
    // the schema's order inside a series: marker before the data, smooth last
    expect(xml).toMatch(/<\/c:marker><c:cat>.*?<\/c:cat><c:val>.*?<\/c:val><c:smooth val="0"\/>/)
  })

  it('a four-series stock chart marks its open as well as its close', () => {
    const xml = buildChartPartXml({
      kind: 'stock',
      categories: ['Mon'],
      series: [
        { name: 'Open', values: [20] },
        { name: 'High', values: [55] },
        { name: 'Low', values: [11] },
        { name: 'Close', values: [32] },
      ],
    })
    const symbols = [...xml.matchAll(/<c:symbol val="(\w+)"\/>/g)].map((m) => m[1])
    expect(symbols).toEqual(['dash', 'none', 'none', 'dash'])
  })

  it('a surface chart reads back as one: seen from above, with its three axes', () => {
    const display = read({ kind: 'surface', ...DATA })
    expect(display?.kind).toBe('surface')
    expect(display?.series.map((s) => s.name)).toEqual(['North', 'South'])
    const xml = buildChartPartXml({ kind: 'surface', ...DATA })
    expect(xml).toMatch(/<c:autoTitleDeleted val="0"\/><c:view3D><c:rotX val="90"\/>/)
    expect(xml).toContain('<c:surfaceChart><c:wireframe val="0"/>')
    expect(
      xml.match(/<c:axId val="\d+"\/>(?=.*<\/c:surfaceChart>)/gs)?.length,
    ).toBeGreaterThanOrEqual(3)
    for (const axis of ['<c:catAx>', '<c:valAx>', '<c:serAx>']) expect(xml).toContain(axis)
    // the other types are not turned
    expect(buildChartPartXml({ kind: 'bar', ...DATA })).not.toContain('<c:view3D>')
  })

  it('a radar chart carries its style and both axes', () => {
    const xml = buildChartPartXml({ kind: 'radar', ...DATA })
    expect(xml).toContain('<c:radarStyle val="marker"/>')
    expect(xml).toContain('<c:catAx>')
    expect(xml).toContain('<c:valAx>')
  })
})
