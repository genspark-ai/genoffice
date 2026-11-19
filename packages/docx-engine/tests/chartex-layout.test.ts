/** A chart of the 2014 extension reads with the shape it is drawn as, next to its nearest classic kind. */
import { describe, expect, it } from 'vitest'
import { parseChartPartXml } from '../src/index'

const part = (series: string) =>
  '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">' +
  '<cx:chartData><cx:data id="0">' +
  '<cx:strDim type="cat"><cx:lvl ptCount="3">' +
  '<cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt><cx:pt idx="2">C</cx:pt></cx:lvl></cx:strDim>' +
  '<cx:numDim type="val"><cx:lvl ptCount="3">' +
  '<cx:pt idx="0">5</cx:pt><cx:pt idx="1">3</cx:pt><cx:pt idx="2">8</cx:pt></cx:lvl></cx:numDim>' +
  '</cx:data></cx:chartData>' +
  `<cx:chart><cx:plotArea><cx:plotAreaRegion>${series}</cx:plotAreaRegion></cx:plotArea></cx:chart>` +
  '</cx:chartSpace>'

const ser = (layoutId: string, name = 'Series1') =>
  `<cx:series layoutId="${layoutId}"><cx:tx><cx:txData><cx:v>${name}</cx:v></cx:txData></cx:tx>` +
  '<cx:dataId val="0"/></cx:series>'

const read = (series: string) => parseChartPartXml(part(series), 'word/charts/chartEx1.xml')!

describe('chartex layouts', () => {
  it.each([
    ['waterfall', 'waterfall', 'bar'],
    ['funnel', 'funnel', 'bar'],
    ['boxWhisker', 'boxWhisker', 'bar'],
    ['treemap', 'treemap', 'pie'],
    ['sunburst', 'sunburst', 'pie'],
    ['clusteredColumn', 'histogram', 'bar'],
  ] as const)('%s reads as %s over a %s', (layoutId, layout, kind) => {
    const display = read(ser(layoutId))
    expect(display.layout).toBe(layout)
    expect(display.kind).toBe(kind)
    expect(display.categories).toEqual(['A', 'B', 'C'])
    expect(display.series).toEqual([{ name: 'Series1', values: [5, 3, 8] }])
  })

  it('a Pareto is its columns: the line series makes the shape and is not a series of its own', () => {
    const display = read(ser('clusteredColumn') + ser('paretoLine', 'Cumulative'))
    expect(display.layout).toBe('pareto')
    expect(display.kind).toBe('bar')
    expect(display.series.map((s) => s.name)).toEqual(['Series1'])
  })

  it('a layout it does not know keeps the generic fallback', () => {
    const display = read(ser('regionMap'))
    expect(display.layout).toBeUndefined()
    expect(display.kind).toBe('other')
  })
})
