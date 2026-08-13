import { describe, expect, it } from 'vitest'

import { applyChartEdit } from '../src/gateway/xlsx-chart'
import { buildChartXml, type ChartAdd } from '../src/gateway/xlsx-drawing-add'

const chart = (overrides: Partial<ChartAdd> = {}): ChartAdd => ({
  chartType: 'column',
  title: 'Test Chart',
  series: [
    {
      name: 'Series 1',
      categories: ['a', 'b', 'c'],
      values: [10, 20, 30],
    },
  ],
  ...overrides,
})

describe('buildChartXml legend', () => {
  it.each([
    ['right', 'r'],
    ['bottom', 'b'],
    ['top', 't'],
    ['left', 'l'],
  ] as const)('%s writes legendPos %s', (legend, position) => {
    const xml = buildChartXml(chart({ legend }))
    expect(xml).toContain(
      `<c:legend><c:legendPos val="${position}"/><c:overlay val="0"/></c:legend>`,
    )
    expect(xml.match(/<c:legend>/g)).toHaveLength(1)
  })

  it('none omits the legend even on a pie', () => {
    const xml = buildChartXml(chart({ chartType: 'pie', legend: 'none' }))
    expect(xml).not.toContain('<c:legend>')
  })

  it('keeps the defaults when unset: none for one series, bottom for pie', () => {
    expect(buildChartXml(chart())).not.toContain('<c:legend>')
    expect(buildChartXml(chart({ chartType: 'pie' }))).toContain(
      '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>',
    )
  })
})

describe('buildChartXml data labels', () => {
  it('value labels sit after the last series, inside the plot', () => {
    const xml = buildChartXml(chart({ dataLabels: 'value' }))
    expect(xml).toContain(
      '</c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/>' +
        '<c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>' +
        '<c:showBubbleSize val="0"/></c:dLbls><c:axId',
    )
  })

  it('percent labels on a pie', () => {
    const xml = buildChartXml(chart({ chartType: 'pie', dataLabels: 'percent' }))
    expect(xml).toContain('<c:showVal val="0"/>')
    expect(xml).toContain('<c:showPercent val="1"/>')
    expect(xml).toContain('<c:showCatName val="0"/>')
  })

  it('category-percent labels show both flags', () => {
    const xml = buildChartXml(chart({ chartType: 'doughnut', dataLabels: 'category-percent' }))
    expect(xml).toContain('<c:showCatName val="1"/>')
    expect(xml).toContain('<c:showPercent val="1"/>')
    expect(xml).toContain('</c:dLbls><c:holeSize val="50"/>')
  })

  it('none and unset write no dLbls', () => {
    expect(buildChartXml(chart({ dataLabels: 'none' }))).not.toContain('<c:dLbls>')
    expect(buildChartXml(chart())).not.toContain('<c:dLbls>')
  })
})

describe('buildChartXml axis titles', () => {
  it('category and value titles follow axPos on their axes', () => {
    const xml = buildChartXml(chart({ axisTitles: { category: 'Month', value: 'Sales' } }))
    expect(xml).toContain(
      '<c:axPos val="b"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>' +
        '<a:t>Month</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:crossAx',
    )
    expect(xml).toContain(
      '<c:axPos val="l"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Sales</a:t>',
    )
  })

  it('scatter puts the category title on the bottom value axis', () => {
    const xml = buildChartXml(chart({ chartType: 'scatter', axisTitles: { category: 'X' } }))
    expect(xml).toContain(
      '<c:axPos val="b"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>X</a:t>',
    )
  })

  it('pie charts have no axes to title', () => {
    const xml = buildChartXml(
      chart({ chartType: 'pie', axisTitles: { category: 'X', value: 'Y' } }),
    )
    expect(xml).not.toContain('<c:catAx>')
    expect(xml).not.toContain('<a:t>X</a:t>')
    expect(xml).not.toContain('<a:t>Y</a:t>')
  })
})

describe('buildChartXml series color and point colors', () => {
  it('series color writes spPr right after tx', () => {
    const xml = buildChartXml(
      chart({
        series: [{ name: 'S', categories: ['a'], values: [1], color: '#ff8800' }],
      }),
    )
    expect(xml).toContain(
      '</c:tx><c:spPr><a:solidFill><a:srgbClr val="FF8800"/></a:solidFill></c:spPr><c:cat>',
    )
  })

  it('point colors write ascending dPt entries between spPr and cat, palette filling the gaps', () => {
    const xml = buildChartXml(
      chart({
        chartType: 'pie',
        series: [
          {
            name: 'S',
            categories: ['a', 'b', 'c'],
            values: [1, 2, 3],
            color: '#111111',
            pointColors: { 2: '#0000ff', 0: '#00ff00' },
          },
        ],
      }),
    )
    expect(xml).toContain(
      '</c:spPr>' +
        '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr></c:dPt>' +
        '<c:dPt><c:idx val="1"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></c:spPr></c:dPt>' +
        '<c:dPt><c:idx val="2"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></c:spPr></c:dPt>' +
        '<c:cat>',
    )
  })

  it('line series keep the marker between spPr and dPt', () => {
    const xml = buildChartXml(
      chart({
        chartType: 'line',
        series: [
          {
            name: 'S',
            categories: ['a'],
            values: [1],
            color: '#111111',
            pointColors: { 0: '#222222' },
          },
        ],
      }),
    )
    expect(xml).toContain('</c:spPr><c:marker><c:symbol val="none"/></c:marker><c:dPt>')
  })

  it('scatter series carry the color as a line stroke and dPt before xVal', () => {
    const xml = buildChartXml(
      chart({
        chartType: 'scatter',
        series: [
          {
            name: 'S',
            categories: ['1'],
            values: [1],
            color: '#111111',
            pointColors: { 0: '#222222' },
          },
        ],
      }),
    )
    expect(xml).toContain(
      '</c:tx><c:spPr><a:ln><a:solidFill><a:srgbClr val="111111"/></a:solidFill></a:ln></c:spPr>',
    )
    expect(xml).toContain('</c:dPt><c:xVal>')
  })
})

describe('buildChartXml default colors', () => {
  it('column series cycle the Office palette as solid fills', () => {
    const xml = buildChartXml(
      chart({
        series: [
          { name: 'A', categories: ['a'], values: [1] },
          { name: 'B', categories: ['a'], values: [2] },
        ],
      }),
    )
    expect(xml).toContain(
      '<c:tx><c:v>A</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>',
    )
    expect(xml).toContain(
      '<c:tx><c:v>B</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></c:spPr>',
    )
  })

  it('an explicit series color beats the palette', () => {
    const xml = buildChartXml(
      chart({
        series: [{ name: 'A', categories: ['a'], values: [1], color: '#123456' }],
      }),
    )
    expect(xml).toContain('<a:srgbClr val="123456"/>')
    expect(xml).not.toContain('4472C4')
  })

  it('line and scatter series color the stroke, not the fill', () => {
    const line = buildChartXml(chart({ chartType: 'line' }))
    expect(line).toContain(
      '</c:tx><c:spPr><a:ln><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:ln></c:spPr><c:marker>',
    )
    const scatter = buildChartXml(chart({ chartType: 'scatter' }))
    expect(scatter).toContain(
      '</c:tx><c:spPr><a:ln><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:ln></c:spPr><c:xVal>',
    )
  })

  it('pie slices cycle the palette per point instead of a series fill', () => {
    const xml = buildChartXml(chart({ chartType: 'pie' }))
    expect(xml).toContain(
      '<c:tx><c:v>Series 1</c:v></c:tx>' +
        '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr></c:dPt>' +
        '<c:dPt><c:idx val="1"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></c:spPr></c:dPt>' +
        '<c:dPt><c:idx val="2"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="A5A5A5"/></a:solidFill></c:spPr></c:dPt>' +
        '<c:cat>',
    )
  })

  it('doughnut slices wrap the palette past ten points', () => {
    const values = Array.from({ length: 11 }, (_, i) => i + 1)
    const xml = buildChartXml(
      chart({
        chartType: 'doughnut',
        series: [{ name: 'S', categories: values.map(String), values }],
      }),
    )
    expect(xml).toContain(
      '<c:dPt><c:idx val="10"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="4472C4"/>',
    )
  })
})

describe('buildChartXml data label position and format', () => {
  it('numFmt and dLblPos lead the dLbls in schema order', () => {
    const xml = buildChartXml(
      chart({
        dataLabels: 'value',
        dataLabelPosition: 'outside-end',
        dataLabelFormat: '0.0%',
      }),
    )
    expect(xml).toContain(
      '</c:ser><c:dLbls><c:numFmt formatCode="0.0%" sourceLinked="0"/>' +
        '<c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/>',
    )
  })

  it('detail without a mode implies value labels', () => {
    const xml = buildChartXml(chart({ dataLabelPosition: 'center' }))
    expect(xml).toContain(
      '<c:dLbls><c:dLblPos val="ctr"/><c:showLegendKey val="0"/><c:showVal val="1"/>',
    )
  })

  it('none suppresses labels even with a position or format', () => {
    const xml = buildChartXml(
      chart({
        dataLabels: 'none',
        dataLabelPosition: 'center',
        dataLabelFormat: '0%',
      }),
    )
    expect(xml).not.toContain('<c:dLbls>')
  })
})

describe('buildChartXml gridlines and value axis', () => {
  it('gridlines true writes majorGridlines after the value-axis axPos', () => {
    const xml = buildChartXml(chart({ gridlines: true }))
    expect(xml).toContain('<c:axPos val="l"/><c:majorGridlines/><c:crossAx')
    expect(xml.match(/<c:majorGridlines\/>/g)).toHaveLength(1)
  })

  it('gridlines false and unset write none', () => {
    expect(buildChartXml(chart({ gridlines: false }))).not.toContain('majorGridlines')
    expect(buildChartXml(chart())).not.toContain('majorGridlines')
  })

  it('valueAxis bounds land in the value-axis scaling, max before min', () => {
    const xml = buildChartXml(chart({ valueAxis: { min: 2, max: 10 } }))
    expect(xml).toContain(
      '<c:valAx><c:axId val="100000002"/><c:scaling><c:orientation val="minMax"/>' +
        '<c:max val="10"/><c:min val="2"/></c:scaling>',
    )
    expect(xml).toContain(
      '<c:catAx><c:axId val="100000001"/><c:scaling><c:orientation val="minMax"/></c:scaling>',
    )
  })

  it('scatter puts bounds and gridlines on the Y axis only', () => {
    const xml = buildChartXml(
      chart({ chartType: 'scatter', gridlines: true, valueAxis: { max: 5 } }),
    )
    expect(xml).toContain(
      '<c:valAx><c:axId val="100000001"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
        '<c:delete val="0"/><c:axPos val="b"/><c:crossAx',
    )
    expect(xml).toContain(
      '<c:valAx><c:axId val="100000002"/><c:scaling><c:orientation val="minMax"/>' +
        '<c:max val="5"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>',
    )
  })
})

describe('buildChartXml gap width and hole size', () => {
  it('gapWidthPct writes gapWidth between labels slot and axis ids', () => {
    const xml = buildChartXml(chart({ gapWidthPct: 80 }))
    expect(xml).toContain('</c:ser><c:gapWidth val="80"/><c:axId')
  })

  it('unset writes no gapWidth; pie-likes ignore it', () => {
    expect(buildChartXml(chart())).not.toContain('gapWidth')
    expect(buildChartXml(chart({ chartType: 'pie', gapWidthPct: 80 }))).not.toContain('gapWidth')
  })

  it('holeSizePct overrides the doughnut default of 50', () => {
    expect(buildChartXml(chart({ chartType: 'doughnut', holeSizePct: 70 }))).toContain(
      '<c:holeSize val="70"/>',
    )
    expect(buildChartXml(chart({ chartType: 'doughnut' }))).toContain('<c:holeSize val="50"/>')
  })
})

describe('buildChartXml explosion', () => {
  it('series explosion sits after spPr on a pie; other types ignore it', () => {
    const xml = buildChartXml(
      chart({
        chartType: 'pie',
        series: [{ name: 'S', categories: ['a'], values: [1], color: '#111111', explosionPct: 25 }],
      }),
    )
    expect(xml).toContain('</c:spPr><c:explosion val="25"/><c:dPt>')
    expect(
      buildChartXml(
        chart({
          series: [{ name: 'S', categories: ['a'], values: [1], explosionPct: 25 }],
        }),
      ),
    ).not.toContain('explosion')
  })

  it('same-idx color and explosion share one dPt, explosion before spPr', () => {
    const xml = buildChartXml(
      chart({
        chartType: 'pie',
        series: [
          {
            name: 'S',
            categories: ['a', 'b'],
            values: [1, 2],
            pointColors: { 1: '#00ff00' },
            pointExplosions: { 1: 12, 0: 6 },
          },
        ],
      }),
    )
    expect(xml).toContain(
      '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:explosion val="6"/>' +
        '<c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr></c:dPt>' +
        '<c:dPt><c:idx val="1"/><c:bubble3D val="0"/><c:explosion val="12"/>' +
        '<c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr></c:dPt>',
    )
  })
})

describe('buildChartXml output stays editable', () => {
  it('applyChartEdit reworks legend, titles, labels, and point colors', () => {
    const built = buildChartXml(
      chart({
        legend: 'right',
        dataLabels: 'value',
        axisTitles: { category: 'Cat', value: 'Val' },
        series: [
          {
            name: 'S',
            categories: ['a', 'b', 'c'],
            values: [1, 2, 3],
            color: '#111111',
            pointColors: { 1: '#abcdef' },
          },
        ],
      }),
    ).replace(/^<\?xml[^>]*\?>\n/, '')
    const edited = applyChartEdit(built, {
      chartPath: 'xl/charts/chart1.xml',
      title: 'New Title',
      legend: 'bottom',
      dataLabels: 'none',
      axisTitles: { value: 'Sales' },
      pointColors: { 0: { 0: '#123456', 1: '#654321' } },
    })
    expect(edited).toContain('<a:t>New Title</a:t>')
    expect(edited).toContain('<c:legendPos val="b"/>')
    expect(edited.match(/<c:legend>/g)).toHaveLength(1)
    expect(edited).not.toContain('<c:dLbls>')
    expect(edited).toContain('<a:t>Sales</a:t>')
    expect(edited).toContain('<a:t>Cat</a:t>')
    expect(edited).toContain('<a:srgbClr val="123456"/>')
    expect(edited).toContain('<a:srgbClr val="654321"/>')
    expect(edited).not.toContain('ABCDEF')
    const first = edited.indexOf('<c:dPt><c:idx val="0"/>')
    const second = edited.indexOf('<c:dPt><c:idx val="1"/>')
    expect(first).toBeGreaterThan(-1)
    expect(second).toBeGreaterThan(first)
  })

  it('a built pie converts to column and keeps its point colors', () => {
    const built = buildChartXml(
      chart({
        chartType: 'pie',
        series: [
          {
            name: 'S',
            categories: ['a', 'b'],
            values: [1, 2],
            pointColors: { 0: '#00ff00' },
          },
        ],
      }),
    ).replace(/^<\?xml[^>]*\?>\n/, '')
    const edited = applyChartEdit(built, {
      chartPath: 'xl/charts/chart1.xml',
      chartType: 'column',
    })
    expect(edited).toContain('<c:barChart>')
    expect(edited).toContain(
      '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr></c:dPt>',
    )
  })

  it('applyChartEdit reworks gridlines, bounds, and gap width', () => {
    const built = buildChartXml(
      chart({
        gridlines: true,
        valueAxis: { min: 0, max: 100 },
        gapWidthPct: 60,
      }),
    ).replace(/^<\?xml[^>]*\?>\n/, '')
    const edited = applyChartEdit(built, {
      chartPath: 'xl/charts/chart1.xml',
      gridlines: false,
      valueAxis: { min: null, max: 40 },
      gapWidthPct: 90,
    })
    expect(edited).not.toContain('majorGridlines')
    expect(edited).toContain('<c:max val="40"/>')
    expect(edited).not.toContain('<c:min ')
    expect(edited).toContain('<c:gapWidth val="90"/>')
  })

  it('a built doughnut keeps hole size and slice explosions editable', () => {
    const built = buildChartXml(
      chart({
        chartType: 'doughnut',
        holeSizePct: 60,
        series: [
          {
            name: 'S',
            categories: ['a', 'b'],
            values: [1, 2],
            pointExplosions: { 0: 10 },
          },
        ],
      }),
    ).replace(/^<\?xml[^>]*\?>\n/, '')
    const edited = applyChartEdit(built, {
      chartPath: 'xl/charts/chart1.xml',
      holeSizePct: 30,
      pointExplosions: { 0: 20, 1: 5 },
    })
    expect(edited).toContain('<c:holeSize val="30"/>')
    expect(edited).not.toContain('<c:holeSize val="60"/>')
    expect(edited).toContain(
      '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:explosion val="20"/>' +
        '<c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr></c:dPt>',
    )
    expect(edited).toContain(
      '<c:dPt><c:idx val="1"/><c:bubble3D val="0"/><c:explosion val="5"/>' +
        '<c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></c:spPr></c:dPt>',
    )
  })

  it('seriesSet replaces a built chart series and stays colorable', () => {
    const built = buildChartXml(chart({ gapWidthPct: 60 })).replace(/^<\?xml[^>]*\?>\n/, '')
    const edited = applyChartEdit(built, {
      chartPath: 'xl/charts/chart1.xml',
      seriesSet: [
        { name: 'A', values: [1, 2], categories: ['x', 'y'] },
        { name: 'B', values: [3, 4] },
      ],
      seriesColors: { 1: '#00AA55' },
    })
    expect(edited.match(/<c:ser>/g)).toHaveLength(2)
    expect(edited).toContain('<c:gapWidth val="60"/>')
    expect(edited).toContain(
      '<c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>x</c:v></c:pt>',
    )
    expect(edited).toContain(
      '<c:tx><c:v>B</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="00AA55"/></a:solidFill>',
    )
    expect(edited).not.toContain('Series 1')
  })
})
