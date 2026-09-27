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

  it('bounds declared and sparse cache indexes', () => {
    const categories =
      '<c:cat><c:strRef><c:f>S!$A$1</c:f><c:strCache><c:ptCount val="1000000000"/>' +
      '<c:pt idx="999999999"><c:v>too far</c:v></c:pt></c:strCache></c:strRef></c:cat>'
    const ser =
      '<c:ser><c:idx val="0"/><c:order val="0"/>' +
      `${strCache('tx', ['S1'])}${categories}${numCache('val', [1])}</c:ser>`
    const display = parseChartPartXml(chartSpace(`<c:barChart>${ser}</c:barChart>`), 'p')!
    expect(display.categories).toHaveLength(1_048_576)
    expect(display.categories[1_048_575]).toBe('')
  })

  it('keeps a cache index at the last Excel row and drops the one past it', () => {
    const categories =
      '<c:cat><c:strRef><c:f>S!$A$1</c:f><c:strCache><c:ptCount val="2"/>' +
      '<c:pt idx="1048575"><c:v>last</c:v></c:pt>' +
      '<c:pt idx="1048576"><c:v>past</c:v></c:pt></c:strCache></c:strRef></c:cat>'
    const ser =
      '<c:ser><c:idx val="0"/><c:order val="0"/>' +
      `${strCache('tx', ['S1'])}${categories}${numCache('val', [1])}</c:ser>`
    const display = parseChartPartXml(chartSpace(`<c:barChart>${ser}</c:barChart>`), 'p')!
    expect(display.categories).toHaveLength(1_048_576)
    expect(display.categories[1_048_575]).toBe('last')
    expect(display.categories).not.toContain('past')
  })

  it('bounds sparse c:dPt point indexes', () => {
    const dPt =
      '<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr></c:dPt>' +
      '<c:dPt><c:idx val="999999999"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt>'
    const display = parseChartPartXml(bar('clustered', dPt), 'p')!
    expect(display.series[0].pointColors).toEqual([null, '00FF00'])
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

  it('defaults the hole to 50% when c:holeSize is unparseable', () => {
    expect(parseChartPartXml(doughnut('<c:holeSize val="large"/>'), 'p')!.holePct).toBe(50)
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

  it('bounds sparse chartEx level point indexes', () => {
    const xml =
      '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<cx:chartData><cx:data id="0">' +
      '<cx:strDim type="cat"><cx:f>S!$A$2:$A$4</cx:f><cx:lvl ptCount="1000000000">' +
      '<cx:pt idx="0">A</cx:pt><cx:pt idx="999999999">far</cx:pt></cx:lvl></cx:strDim>' +
      '<cx:numDim type="val"><cx:f>S!$B$2:$B$4</cx:f><cx:lvl ptCount="1000000000">' +
      '<cx:pt idx="0">100</cx:pt><cx:pt idx="999999999">7</cx:pt>' +
      '</cx:lvl></cx:numDim></cx:data></cx:chartData>' +
      '<cx:chart><cx:plotArea><cx:plotAreaRegion>' +
      '<cx:series layoutId="waterfall"><cx:tx><cx:txData><cx:v>Series1</cx:v></cx:txData></cx:tx>' +
      '<cx:dataId val="0"/></cx:series>' +
      '</cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>'
    const display = parseChartPartXml(xml, 'word/charts/chartEx1.xml')!
    expect(display.categories.length).toBeLessThanOrEqual(1_048_576)
    expect(display.categories[0]).toBe('A')
    expect(display.categories).not.toContain('far')
    expect(display.series[0].values.length).toBeLessThanOrEqual(1_048_576)
    expect(display.series[0].values[0]).toBe(100)
    expect(display.series[0].values).not.toContain(7)
  })
})

describe('parseChartPartXml presentation features', () => {
  const lineSer = (i: number, values: number[], extra = '') =>
    `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${strCache('tx', [`S${i + 1}`])}${extra}` +
    `${strCache('cat', ['A', 'B', 'C'])}${numCache('val', values)}</c:ser>`
  const axes =
    '<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
    '<c:valAx><c:axId val="2"/><c:axPos val="l"/><c:title/><c:crossAx val="1"/></c:valAx>'

  it('reads radar charts with their style and honors symbol="none" over the marker style', () => {
    const radar = (style: string, marker = '') =>
      chartSpace(
        `<c:radarChart><c:radarStyle val="${style}"/>${lineSer(0, [1, 2, 3], marker)}</c:radarChart>`,
      )
    const filled = parseChartPartXml(radar('filled'), 'p')!
    expect(filled.kind).toBe('radar')
    expect(filled.radarStyle).toBe('filled')
    expect(filled.markers).toBeUndefined()
    expect(parseChartPartXml(radar('marker'), 'p')!.markers).toBe(true)
    const off = '<c:marker><c:symbol val="none"/></c:marker>'
    expect(parseChartPartXml(radar('marker', off), 'p')!.markers).toBeUndefined()
    expect(parseChartPartXml(radar('standard'), 'p')!.radarStyle).toBe('standard')
  })

  it('keeps stacked grouping on line charts', () => {
    const xml = chartSpace(
      `<c:lineChart><c:grouping val="stacked"/>${lineSer(0, [1, 2, 3])}${lineSer(1, [1, 1, 1])}` +
        `<c:marker val="1"/></c:lineChart>`,
    )
    const display = parseChartPartXml(xml, 'p')!
    expect(display.grouping).toBe('stacked')
    expect(display.markers).toBe(true)
  })

  it('reads pie explosion and percent data labels', () => {
    const ser =
      `<c:ser><c:idx val="0"/><c:order val="0"/>${strCache('tx', ['S1'])}` +
      `<c:explosion val="25"/><c:dLbls><c:showPercent val="1"/></c:dLbls>` +
      `${strCache('cat', ['A', 'B'])}${numCache('val', [3, 1])}</c:ser>`
    const display = parseChartPartXml(chartSpace(`<c:pieChart>${ser}</c:pieChart>`), 'p')!
    expect(display.explosionPct).toBe(25)
    expect(display.dataLabels).toEqual({ pct: true })
  })

  it('flags parts without c:legend and resolves the chart-area border', () => {
    const line = (post = '', spPr = '') =>
      chartSpace(`<c:lineChart>${lineSer(0, [1, 2, 3])}</c:lineChart>`, '', post + spPr)
    const bare = parseChartPartXml(line(), 'p')!
    expect(bare.noLegend).toBe(true)
    expect(bare.legendPos).toBeUndefined()
    // Word's automatic border shows on parts without c:chartSpace/c:spPr
    expect(bare.frameLine).toBe('868686')
    const withLegend = parseChartPartXml(line('<c:legend><c:legendPos val="t"/></c:legend>'), 'p')!
    expect(withLegend.legendPos).toBe('t')
    expect(withLegend.noLegend).toBeUndefined()
    const noBorder =
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      `<c:chart><c:plotArea><c:lineChart>${lineSer(0, [1, 2, 3])}</c:lineChart></c:plotArea></c:chart>` +
      '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:chartSpace>'
    expect(parseChartPartXml(noBorder, 'p')!.frameLine).toBeUndefined()
    const gray = noBorder.replace(
      '<a:ln><a:noFill/></a:ln>',
      '<a:ln><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln>',
    )
    expect(parseChartPartXml(gray, 'p')!.frameLine).toBe('D9D9D9')
  })

  it('classifies axes by position with placeholder titles and automatic lines', () => {
    const xml = chartSpace(`<c:lineChart>${lineSer(0, [1, 2, 3])}</c:lineChart>${axes}`)
    const display = parseChartPartXml(xml, 'p', THEME)!
    expect(display.xAxis).toEqual({ line: '404040' })
    // an empty c:title shows Word's placeholder text
    expect(display.yAxis).toEqual({ title: 'Axis Title', line: '404040' })
    const titled = xml.replace(
      '<c:title/>',
      '<c:title><c:tx><c:rich><a:p><a:r><a:t>Units</a:t></a:r></a:p></c:rich></c:tx></c:title>',
    )
    expect(parseChartPartXml(titled, 'p')!.yAxis!.title).toBe('Units')
    const hidden = xml.replace(
      '<c:axPos val="b"/>',
      '<c:axPos val="b"/><c:delete val="1"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>',
    )
    expect(parseChartPartXml(hidden, 'p')!.xAxis).toEqual({ deleted: true })
  })

  it('reads the data table flags', () => {
    const xml = chartSpace(
      `<c:barChart><c:barDir val="col"/>${lineSer(0, [1, 2, 3])}</c:barChart>${axes}` +
        '<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="1"/></c:dTable>',
    )
    expect(parseChartPartXml(xml, 'p')!.dataTable).toEqual({
      keys: true,
      horz: true,
      vert: true,
      outline: true,
      line: '404040',
    })
    expect(
      parseChartPartXml(xml.replace(/<c:dTable>.*<\/c:dTable>/, ''), 'p')!.dataTable,
    ).toBeUndefined()
  })
})

describe('parseChartPartXml text sizes and label overrides', () => {
  const rPr = (sz: number, color = '000000') =>
    `<a:pPr><a:defRPr sz="${sz}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:defRPr></a:pPr>`
  const txPr = (sz: number, color?: string) =>
    `<c:txPr><a:bodyPr/><a:lstStyle/><a:p>${rPr(sz, color)}</a:p></c:txPr>`
  const flags = (val: number, cat: number, pct: number) =>
    `<c:showLegendKey val="0"/><c:showVal val="${val}"/><c:showCatName val="${cat}"/>` +
    `<c:showSerName val="0"/><c:showPercent val="${pct}"/><c:showBubbleSize val="0"/>`
  const title =
    '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1800"/></a:pPr>' +
    '<a:r><a:rPr sz="1800"/><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>'

  it('reads title, axis and legend sizes, axis label color and major gridline color', () => {
    const chart = parseChartPartXml(
      chartSpace(
        `<c:barChart><c:barDir val="col"/>${barSer(0, [1, 2])}<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
          `<c:catAx><c:axId val="1"/><c:axPos val="b"/>${txPr(1200, '000000')}<c:crossAx val="2"/></c:catAx>` +
          '<c:valAx><c:axId val="2"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln>' +
          '<a:solidFill><a:srgbClr val="888888"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>' +
          `${txPr(1200)}<c:crossAx val="1"/></c:valAx>`,
        '',
        `${title}<c:legend><c:legendPos val="b"/>${txPr(900)}</c:legend>`,
      ),
      'word/charts/chart1.xml',
      THEME,
    )!
    expect(chart.titleFontPt).toBe(18)
    expect(chart.legendFontPt).toBe(9)
    expect(chart.xAxis).toMatchObject({ fontPt: 12, color: '000000' })
    expect(chart.yAxis).toMatchObject({ fontPt: 12, gridLine: '888888' })
  })

  it('gives an automatic major gridline the Office tint and none without c:majorGridlines', () => {
    const axes = (grid: string) =>
      `<c:barChart><c:barDir val="col"/>${barSer(0, [1, 2])}<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
      '<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
      `<c:valAx><c:axId val="2"/><c:axPos val="l"/>${grid}<c:crossAx val="1"/></c:valAx>`
    const auto = parseChartPartXml(chartSpace(axes('<c:majorGridlines/>')), 'p', THEME)!
    expect(auto.yAxis?.gridLine).toBe('D9D9D9')
    const none = parseChartPartXml(chartSpace(axes('')), 'p', THEME)!
    expect(none.yAxis?.gridLine).toBeUndefined()
    expect(none.titleFontPt).toBeUndefined()
  })

  it('reads bar value labels with their size, color and number format', () => {
    const dLbls = `<c:dLbls><c:numFmt formatCode="#,##0" sourceLinked="0"/>${txPr(1200, '000000')}${flags(1, 0, 0)}</c:dLbls>`
    const chart = parseChartPartXml(
      chartSpace(`<c:barChart><c:barDir val="col"/>${barSer(0, [1997, 2213], dLbls)}</c:barChart>`),
      'p',
      THEME,
    )!
    expect(chart.dataLabels).toEqual({ val: true, fontPt: 12, color: '000000', numFmt: '#,##0' })
  })

  it('lets per-point c:dLbl blocks covering every point override the series flags', () => {
    const point = (i: number) =>
      `<c:dLbl><c:idx val="${i}"/><c:numFmt formatCode="0%" sourceLinked="0"/>${txPr(1200)}${flags(0, 0, 1)}</c:dLbl>`
    const dLbls =
      `<c:dLbls>${point(0)}${point(1)}<c:numFmt formatCode="0%" sourceLinked="0"/>${txPr(1800)}` +
      `${flags(0, 1, 1)}</c:dLbls>`
    const full = parseChartPartXml(
      chartSpace(
        `<c:doughnutChart>${barSer(0, [3, 7], dLbls)}<c:holeSize val="50"/></c:doughnutChart>`,
      ),
      'p',
      THEME,
    )!
    // Word shows only what the point blocks enable (percent, 12pt), not the series category names
    expect(full.dataLabels).toEqual({ pct: true, fontPt: 12, color: '000000', numFmt: '0%' })
    const partial = parseChartPartXml(
      chartSpace(
        `<c:doughnutChart>${barSer(0, [3, 7], `<c:dLbls>${point(0)}${txPr(1800)}${flags(0, 1, 1)}</c:dLbls>`)}</c:doughnutChart>`,
      ),
      'p',
      THEME,
    )!
    expect(partial.dataLabels).toEqual({ pct: true, cat: true, fontPt: 18, color: '000000' })
  })
})
