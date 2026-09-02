import { describe, expect, it } from 'vitest'
import { parseChartPartXml } from '../src/index'
import type { ThemeColors } from '../src/types'

const THEME: ThemeColors = {
  dk1: '000000',
  lt1: 'FFFFFF',
  accent1: '4F81BD',
  accent2: 'C0504D',
  accent3: '9BBB59',
  accent4: '8064A2',
  accent5: '4BACC6',
  accent6: 'F79646',
}

const chartSpace = (inner: string, pre = '', post = '') =>
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  `${pre}<c:chart><c:plotArea><c:layout/>${inner}</c:plotArea>${post}</c:chart></c:chartSpace>`

const numCache = (tag: string, values: (number | string)[], fmt = 'General') =>
  `<c:${tag}><c:numRef><c:f>S!$A$1</c:f><c:numCache><c:formatCode>${fmt}</c:formatCode>` +
  `<c:ptCount val="${values.length}"/>` +
  values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('') +
  `</c:numCache></c:numRef></c:${tag}>`

const strCache = (tag: string, values: string[]) =>
  `<c:${tag}><c:strRef><c:f>S!$A$1</c:f><c:strCache><c:ptCount val="${values.length}"/>` +
  values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('') +
  `</c:strCache></c:strRef></c:${tag}>`

const barSer = (i: number, values: number[], extra = '') =>
  `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` +
  `${strCache('tx', [`Series ${i + 1}`])}${extra}` +
  `${strCache('cat', ['A', 'B'])}${numCache('val', values)}</c:ser>`

describe('parseChartPartXml grouping and colors', () => {
  const bar = (grouping: string, extra = '') =>
    chartSpace(
      `<c:barChart><c:barDir val="col"/><c:grouping val="${grouping}"/>` +
        `${barSer(0, [1, 2], extra)}${barSer(1, [3, 4])}</c:barChart>`,
    )

  it('keeps stacked / percentStacked grouping and drops clustered', () => {
    expect(parseChartPartXml(bar('stacked'), 'p')!.grouping).toBe('stacked')
    expect(parseChartPartXml(bar('percentStacked'), 'p')!.grouping).toBe('percentStacked')
    expect(parseChartPartXml(bar('clustered'), 'p')!.grouping).toBeUndefined()
  })

  it('reads explicit series solid fills, resolving schemeClr through the theme', () => {
    const srgb = '<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>'
    expect(parseChartPartXml(bar('clustered', srgb), 'p')!.series[0].color).toBe('FF0000')
    const scheme = '<c:spPr><a:solidFill><a:schemeClr val="accent3"/></a:solidFill></c:spPr>'
    expect(parseChartPartXml(bar('clustered', scheme), 'p', THEME)!.series[0].color).toBe('9BBB59')
    // without a theme the scheme reference cannot resolve
    expect(parseChartPartXml(bar('clustered', scheme), 'p')!.series[0].color).toBeUndefined()
  })

  it('darkens lumMod-modified scheme colors', () => {
    const mod =
      '<c:spPr><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="60000"/></a:schemeClr>' +
      '</a:solidFill></c:spPr>'
    const color = parseChartPartXml(bar('clustered', mod), 'p', THEME)!.series[0].color!
    const lum = (hex: string) =>
      [0, 2, 4].reduce((a, i) => a + parseInt(hex.slice(i, i + 2), 16), 0)
    expect(lum(color)).toBeLessThan(lum('4F81BD'))
  })

  it('collects c:dPt point fills sparsely by index', () => {
    const dPt =
      '<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>' +
      '</c:spPr></c:dPt>'
    const display = parseChartPartXml(bar('clustered', dPt), 'p')!
    expect(display.series[0].pointColors).toEqual([null, '00FF00'])
  })
})

describe('parseChartPartXml chart-style palette', () => {
  const clustered = `<c:barChart><c:barDir val="col"/>${barSer(0, [1, 2])}</c:barChart>`

  it('theme accents drive the default cycle', () => {
    expect(parseChartPartXml(chartSpace(clustered), 'p', THEME)!.palette).toEqual([
      '4F81BD',
      'C0504D',
      '9BBB59',
      '8064A2',
      '4BACC6',
      'F79646',
    ])
    expect(parseChartPartXml(chartSpace(clustered), 'p')!.palette).toBeUndefined()
  })

  it('style column 1 is grayscale, also via the c14 AlternateContent form', () => {
    const direct = parseChartPartXml(chartSpace(clustered, '<c:style val="1"/>'), 'p', THEME)!
    expect(direct.palette![0]).toBe('595959')
    const alt =
      '<mc:AlternateContent><mc:Choice Requires="c14" ' +
      'xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">' +
      '<c14:style val="101"/></mc:Choice><mc:Fallback><c:style val="1"/></mc:Fallback>' +
      '</mc:AlternateContent>'
    expect(parseChartPartXml(chartSpace(clustered, alt), 'p', THEME)!.palette![0]).toBe('595959')
  })

  it('single-accent columns lead with their accent', () => {
    const display = parseChartPartXml(chartSpace(clustered, '<c:style val="29"/>'), 'p', THEME)!
    expect(display.palette![0]).toBe('9BBB59')
  })
})

describe('parseChartPartXml scatter and bubble', () => {
  it('reads x/y pairs, date category texts and the noFill line suppression', () => {
    const ser =
      '<c:ser><c:idx val="0"/><c:order val="0"/>' +
      strCache('tx', ['S1']) +
      '<c:spPr><a:ln w="28575"><a:noFill/></a:ln></c:spPr>' +
      numCache('xVal', [37377, 37408], 'm/d/yyyy') +
      numCache('yVal', [55, 57]) +
      '</c:ser>'
    const xml = chartSpace(
      `<c:scatterChart><c:scatterStyle val="lineMarker"/>${ser}</c:scatterChart>`,
    )
    const display = parseChartPartXml(xml, 'p')!
    expect(display.kind).toBe('scatter')
    expect(display.markers).toBe(true)
    expect(display.categories).toEqual(['5/1/2002', '6/1/2002'])
    expect(display.series[0].values).toEqual([55, 57])
    expect(display.series[0].xValues).toEqual([37377, 37408])
    expect(display.series[0].line).toBeUndefined()
  })

  it('keeps the line for lineMarker series without noFill', () => {
    const ser =
      `<c:ser><c:idx val="0"/><c:order val="0"/>${numCache('xVal', [1, 2])}` +
      `${numCache('yVal', [3, 4])}</c:ser>`
    const xml = chartSpace(
      `<c:scatterChart><c:scatterStyle val="lineMarker"/>${ser}</c:scatterChart>`,
    )
    expect(parseChartPartXml(xml, 'p')!.series[0].line).toBe(true)
  })

  it('names a single-series auto title after the series', () => {
    const ser =
      `<c:ser><c:idx val="0"/><c:order val="0"/>${strCache('tx', ['Y-Values'])}` +
      `${numCache('xVal', [1, 2])}${numCache('yVal', [3, 4])}</c:ser>`
    const xml = chartSpace(`<c:bubbleChart>${ser}</c:bubbleChart>`).replace(
      '<c:chart>',
      '<c:chart><c:title/>',
    )
    expect(parseChartPartXml(xml, 'p')!.title).toBe('Y-Values')
  })

  it('reads bubble sizes and trims raw double x texts', () => {
    const ser =
      '<c:ser><c:idx val="0"/><c:order val="0"/>' +
      strCache('tx', ['Y-Values']) +
      numCache('xVal', ['0.70000000000000062', 1.8]) +
      numCache('yVal', [2.7, 3.2]) +
      numCache('bubbleSize', [10, 4]) +
      '</c:ser>'
    const display = parseChartPartXml(chartSpace(`<c:bubbleChart>${ser}</c:bubbleChart>`), 'p')!
    expect(display.kind).toBe('bubble')
    expect(display.categories).toEqual(['0.7', '1.8'])
    expect(display.series[0].sizes).toEqual([10, 4])
  })
})

describe('parseChartPartXml doughnut hole and legend position', () => {
  const pieSer =
    `<c:ser><c:idx val="0"/><c:order val="0"/>${strCache('tx', ['S1'])}` +
    `${strCache('cat', ['A', 'B'])}${numCache('val', [67, 33])}</c:ser>`
  const doughnut = (hole: string) =>
    chartSpace(`<c:doughnutChart><c:varyColors val="1"/>${pieSer}${hole}</c:doughnutChart>`)

  it('reads doughnut charts as pie carrying the hole size', () => {
    const display = parseChartPartXml(doughnut('<c:holeSize val="65"/>'), 'p')!
    expect(display.kind).toBe('pie')
    expect(display.holePct).toBe(65)
  })

  it('defaults the hole to 50% when c:holeSize is absent', () => {
    expect(parseChartPartXml(doughnut(''), 'p')!.holePct).toBe(50)
  })

  it('leaves plain pies without a hole', () => {
    const xml = chartSpace(`<c:pieChart><c:varyColors val="1"/>${pieSer}</c:pieChart>`)
    expect(parseChartPartXml(xml, 'p')!.holePct).toBeUndefined()
  })

  it('reads the legend position, defaulting an empty c:legend to right', () => {
    const pie = (legend: string) =>
      chartSpace(`<c:pieChart><c:varyColors val="1"/>${pieSer}</c:pieChart>`, '', legend)
    expect(
      parseChartPartXml(
        pie('<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>'),
        'p',
      )!.legendPos,
    ).toBe('r')
    expect(
      parseChartPartXml(pie('<c:legend><c:legendPos val="b"/></c:legend>'), 'p')!.legendPos,
    ).toBe('b')
    expect(parseChartPartXml(pie('<c:legend/>'), 'p')!.legendPos).toBe('r')
    expect(parseChartPartXml(pie(''), 'p')!.legendPos).toBeUndefined()
  })
})

describe('parseChartexPartXml leniency', () => {
  it('parses a waterfall whose cx:chartData was renamed (unknown-element corpus)', () => {
    const xml =
      '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<cx:chartDataIntentionallyChanged><cx:data id="0">' +
      '<cx:numDim type="val"><cx:f>S!$A$2:$A$4</cx:f><cx:lvl ptCount="3">' +
      '<cx:pt idx="0">100</cx:pt><cx:pt idx="1">-40</cx:pt><cx:pt idx="2">60</cx:pt>' +
      '</cx:lvl></cx:numDim></cx:data></cx:chartDataIntentionallyChanged>' +
      '<cx:chart><cx:plotArea><cx:plotAreaRegion>' +
      '<cx:series layoutId="waterfall"><cx:tx><cx:txData><cx:v>Series1</cx:v></cx:txData></cx:tx>' +
      '<cx:dataId val="0"/></cx:series>' +
      '</cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>'
    const display = parseChartPartXml(xml, 'word/charts/chartEx1.xml')!
    expect(display.kind).toBe('bar')
    expect(display.series[0]).toEqual({ name: 'Series1', values: [100, -40, 60] })
  })
})
