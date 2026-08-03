import { describe, expect, it } from 'vitest'

import { applyChartEdit, ChartEditError } from '../src/gateway/xlsx-chart'

const BAR_CHART =
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<c:chart>' +
  '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr b="1"/><a:t>Old title</a:t></a:r>' +
  '<a:r><a:t> suffix</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
  '<c:autoTitleDeleted val="0"/>' +
  '<c:plotArea><c:layout/>' +
  '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
  '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>' +
  '<c:invertIfNegative val="0"/><c:cat/><c:val/></c:ser>' +
  '<c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:strRef><c:f>Data!$C$1</c:f></c:strRef></c:tx>' +
  '<c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></c:spPr><c:cat/><c:val/></c:ser>' +
  '<c:gapWidth val="150"/><c:overlap val="-27"/>' +
  '<c:axId val="111"/><c:axId val="222"/></c:barChart>' +
  '<c:catAx/><c:valAx/></c:plotArea>' +
  '</c:chart></c:chartSpace>'

const UNTITLED_PIE =
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<c:chart><c:autoTitleDeleted val="1"/>' +
  '<c:plotArea><c:pieChart><c:varyColors val="1"/>' +
  '<c:ser><c:idx val="0"/><c:order val="0"/><c:cat/><c:val/></c:ser>' +
  '</c:pieChart></c:plotArea></c:chart></c:chartSpace>'

describe('applyChartEdit title', () => {
  it('replaces the first run text, blanks the rest, keeps run styling', () => {
    const xml = applyChartEdit(BAR_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      title: 'New <Title>',
    })
    expect(xml).toContain('<a:rPr b="1"/><a:t>New &lt;Title&gt;</a:t>')
    expect(xml).toContain('<a:r><a:t></a:t></a:r>')
    expect(xml).not.toContain('Old title')
  })

  it('inserts a title and re-enables it on an auto-title-deleted chart', () => {
    const xml = applyChartEdit(UNTITLED_PIE, { chartPath: 'xl/charts/chart1.xml', title: 'Pie' })
    expect(xml).toContain(
      '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Pie</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>',
    )
    expect(xml.match(/<c:autoTitleDeleted /g)).toHaveLength(1)
    expect(xml).toContain('<c:autoTitleDeleted val="0"/>')
    expect(xml).not.toContain('<c:autoTitleDeleted val="1"/>')
  })

  it('does not clobber an axis title when the chart has none of its own', () => {
    const axisTitled = BAR_CHART.replace(/<c:title>[\s\S]*?<\/c:title>/, '').replace(
      '<c:catAx/>',
      '<c:catAx><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r>' +
        '<a:t>Month</a:t></a:r></a:p></c:rich></c:tx></c:title></c:catAx>',
    )
    const xml = applyChartEdit(axisTitled, { chartPath: 'xl/charts/chart1.xml', title: 'My Chart' })
    expect(xml).toContain('<a:t>Month</a:t>')
    const chartLevel = xml.slice(0, xml.indexOf('<c:plotArea>'))
    expect(chartLevel).toContain('<a:t>My Chart</a:t>')
    expect(xml.match(/<c:autoTitleDeleted /g)).toHaveLength(1)
  })
})

describe('applyChartEdit type conversion', () => {
  it('column to bar only flips barDir', () => {
    const xml = applyChartEdit(BAR_CHART, { chartPath: 'xl/charts/chart1.xml', chartType: 'bar' })
    expect(xml).toContain('<c:barDir val="bar"/>')
    expect(xml).toContain('<c:overlap val="-27"/>')
  })

  it('column to line rebuilds the plot keeping series and axes', () => {
    const xml = applyChartEdit(BAR_CHART, { chartPath: 'xl/charts/chart1.xml', chartType: 'line' })
    expect(xml).toContain('<c:lineChart>')
    expect(xml).not.toContain('<c:barChart>')
    expect(xml).not.toContain('barDir')
    expect(xml).not.toContain('invertIfNegative')
    expect(xml).toContain('<c:f>Data!$B$1</c:f>')
    expect(xml).toContain('<c:marker val="1"/>')
    expect(xml).toContain('<c:axId val="111"/><c:axId val="222"/>')
    // Existing series colors survive the conversion.
    expect(xml).toContain('<a:srgbClr val="112233"/>')
  })

  it('column to pie drops the axes and varies slice colors', () => {
    const xml = applyChartEdit(BAR_CHART, { chartPath: 'xl/charts/chart1.xml', chartType: 'pie' })
    expect(xml).toContain('<c:pieChart>')
    expect(xml).toContain('<c:varyColors val="1"/>')
    expect(xml).toContain('<c:firstSliceAng val="0"/>')
    expect(xml).not.toContain('<c:barChart>')
    expect(xml).not.toContain('axId')
    expect(xml).not.toContain('catAx')
    expect(xml).not.toContain('valAx')
    expect(xml).not.toContain('grouping')
    expect(xml).toContain('<c:f>Data!$B$1</c:f>')
  })

  it('pie to column synthesizes a category/value axis pair', () => {
    const xml = applyChartEdit(UNTITLED_PIE, {
      chartPath: 'xl/charts/chart1.xml',
      chartType: 'column',
    })
    expect(xml).toContain('<c:barChart>')
    expect(xml).toContain('<c:barDir val="col"/>')
    expect(xml).toContain('<c:grouping val="clustered"/>')
    expect(xml).not.toContain('<c:pieChart>')
    expect(xml).not.toContain('firstSliceAng')
    expect(xml).toContain('<c:axId val="900000001"/><c:axId val="900000002"/>')
    expect(xml).toContain('<c:catAx><c:axId val="900000001"/>')
    expect(xml).toContain('<c:valAx><c:axId val="900000002"/>')
    expect(xml).toContain('<c:crossAx val="900000001"/>')
  })

  it('pie to doughnut rebuilds the plot with a hole', () => {
    const xml = applyChartEdit(UNTITLED_PIE, {
      chartPath: 'xl/charts/chart1.xml',
      chartType: 'doughnut',
    })
    expect(xml).toContain('<c:doughnutChart>')
    expect(xml).toContain('<c:holeSize val="50"/>')
    expect(xml).not.toContain('<c:pieChart>')
  })

  it('refuses to convert a scatter chart', () => {
    const scatter = UNTITLED_PIE.replace(/pieChart/g, 'scatterChart')
    expect(() =>
      applyChartEdit(scatter, { chartPath: 'xl/charts/chart1.xml', chartType: 'line' }),
    ).toThrow(ChartEditError)
  })

  it('keeps plot-level data labels across a conversion, dropping the position', () => {
    const labeled = applyChartEdit(BAR_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      dataLabels: 'value',
      dataLabelPosition: 'outside-end',
    })
    const xml = applyChartEdit(labeled, { chartPath: 'xl/charts/chart1.xml', chartType: 'line' })
    expect(xml).toContain('<c:showVal val="1"/>')
    expect(xml).not.toContain('<c:dLblPos')
    // dLbls lands after the series, before the line marker.
    expect(xml.indexOf('<c:dLbls>')).toBeGreaterThan(xml.lastIndexOf('</c:ser>'))
    expect(xml.indexOf('<c:dLbls>')).toBeLessThan(xml.indexOf('<c:marker val="1"/>'))
  })
})

describe('applyChartEdit data labels', () => {
  it('inserts plot-level labels after the last series', () => {
    const xml = applyChartEdit(BAR_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      dataLabels: 'value',
    })
    expect(xml).toContain('</c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/>')
    expect(xml).toContain('<c:showPercent val="0"/>')
    // After the SECOND series, not the first.
    expect(/<\/c:ser><c:dLbls>/.exec(xml)?.index).toBeGreaterThan(xml.indexOf('Data!$C$1'))
  })

  it('writes category + percent labels on a pie and removes them on none', () => {
    const labeled = applyChartEdit(UNTITLED_PIE, {
      chartPath: 'xl/charts/chart1.xml',
      dataLabels: 'category-percent',
    })
    expect(labeled).toContain('<c:showCatName val="1"/>')
    expect(labeled).toContain('<c:showPercent val="1"/>')
    expect(labeled).toContain('<c:showVal val="0"/>')
    const cleared = applyChartEdit(labeled, {
      chartPath: 'xl/charts/chart1.xml',
      dataLabels: 'none',
    })
    expect(cleared).not.toContain('<c:dLbls>')
  })
})

describe('applyChartEdit series colors', () => {
  it('inserts spPr after tx when the series has none', () => {
    const xml = applyChartEdit(BAR_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      seriesColors: { 0: '#FF8800' },
    })
    expect(xml).toContain(
      '</c:tx><c:spPr><a:solidFill><a:srgbClr val="FF8800"/></a:solidFill>' +
        '<a:ln><a:solidFill><a:srgbClr val="FF8800"/></a:solidFill></a:ln></c:spPr><c:invertIfNegative',
    )
  })

  it('replaces an existing spPr', () => {
    const xml = applyChartEdit(BAR_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      seriesColors: { 1: '#00AA55' },
    })
    expect(xml).toContain('<a:srgbClr val="00AA55"/>')
    expect(xml).not.toContain('112233')
  })

  it('fails closed on a missing series index', () => {
    expect(() =>
      applyChartEdit(BAR_CHART, {
        chartPath: 'xl/charts/chart1.xml',
        seriesColors: { 7: '#00AA55' },
      }),
    ).toThrow(ChartEditError)
  })

  // The sidecar numbers series by document order; a file whose c:idx values
  // are shuffled must still edit the series the UI showed as N.
  it('targets series by document position, not c:idx', () => {
    const shuffled = BAR_CHART.replace(
      '<c:idx val="0"/><c:order val="0"/>',
      '<c:idx val="5"/><c:order val="0"/>',
    ).replace('<c:idx val="1"/><c:order val="1"/>', '<c:idx val="0"/><c:order val="1"/>')
    const xml = applyChartEdit(shuffled, {
      chartPath: 'xl/charts/chart1.xml',
      seriesColors: { 0: '#FF8800' },
    })
    const firstSer = /<c:ser>[\s\S]*?<\/c:ser>/.exec(xml)?.[0] ?? ''
    expect(firstSer).toContain('<c:idx val="5"/>')
    expect(firstSer).toContain('FF8800')
    expect(xml).toContain('112233')
  })
})

describe('applyChartEdit grouping', () => {
  it('stacks a bar chart and overlaps the bars fully', () => {
    const xml = applyChartEdit(BAR_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      grouping: 'stacked',
    })
    expect(xml).toContain('<c:grouping val="stacked"/>')
    expect(xml).toContain('<c:overlap val="100"/>')
    expect(xml).not.toContain('<c:overlap val="-27"/>')
  })

  it('returns a bar chart to clustered with the Excel default overlap', () => {
    const stacked = applyChartEdit(BAR_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      grouping: 'percentStacked',
    })
    const xml = applyChartEdit(stacked, {
      chartPath: 'xl/charts/chart1.xml',
      grouping: 'clustered',
    })
    expect(xml).toContain('<c:grouping val="clustered"/>')
    expect(xml).toContain('<c:overlap val="-27"/>')
  })

  it('writes standard (not clustered) on a line chart and inserts a missing overlap slot only for bars', () => {
    const line = applyChartEdit(BAR_CHART, { chartPath: 'xl/charts/chart1.xml', chartType: 'line' })
    const xml = applyChartEdit(line, { chartPath: 'xl/charts/chart1.xml', grouping: 'stacked' })
    expect(xml).toContain('<c:grouping val="stacked"/>')
    const back = applyChartEdit(xml, { chartPath: 'xl/charts/chart1.xml', grouping: 'clustered' })
    expect(back).toContain('<c:grouping val="standard"/>')
    expect(back).not.toContain('<c:overlap')
  })

  it('fails closed on a pie chart', () => {
    expect(() =>
      applyChartEdit(UNTITLED_PIE, { chartPath: 'xl/charts/chart1.xml', grouping: 'stacked' }),
    ).toThrow(ChartEditError)
  })
})

const PIE_WITH_DPT =
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<c:chart><c:plotArea><c:pieChart><c:varyColors val="1"/>' +
  '<c:ser><c:idx val="0"/><c:order val="0"/>' +
  '<c:dPt><c:idx val="1"/><c:bubble3D val="0"/><c:explosion val="5"/>' +
  '<c:spPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>' +
  '<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr></c:dPt>' +
  '<c:cat/><c:val/></c:ser>' +
  '</c:pieChart></c:plotArea></c:chart></c:chartSpace>'

describe('applyChartEdit point colors', () => {
  it('writes new dPt entries in ascending idx order before the data', () => {
    const xml = applyChartEdit(UNTITLED_PIE, {
      chartPath: 'xl/charts/chart1.xml',
      pointColors: { 0: { 1: '#ff0000', 0: '#00ff00' } },
    })
    expect(xml).toContain(
      '<c:order val="0"/>' +
        '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr></c:dPt>' +
        '<c:dPt><c:idx val="1"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt>' +
        '<c:cat/>',
    )
  })

  it('replaces only the fill of an existing dPt, keeping other children', () => {
    const xml = applyChartEdit(PIE_WITH_DPT, {
      chartPath: 'xl/charts/chart1.xml',
      pointColors: { 0: { 1: '#112233' } },
    })
    expect(xml).toContain('<a:solidFill><a:srgbClr val="112233"/></a:solidFill>')
    expect(xml).not.toContain('ABCDEF')
    expect(xml).toContain('<c:explosion val="5"/>')
    // the line fill inside a:ln stays untouched
    expect(xml).toContain('<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>')
    expect(xml.match(/<c:dPt>/g)).toHaveLength(1)
  })

  it('interleaves new dPt entries around an existing one by idx', () => {
    const xml = applyChartEdit(PIE_WITH_DPT, {
      chartPath: 'xl/charts/chart1.xml',
      pointColors: { 0: { 2: '#0000ff', 0: '#00ff00' } },
    })
    const first = xml.indexOf('<c:dPt><c:idx val="0"/>')
    const second = xml.indexOf('<c:dPt><c:idx val="1"/>')
    const third = xml.indexOf('<c:dPt><c:idx val="2"/>')
    expect(first).toBeGreaterThan(-1)
    expect(second).toBeGreaterThan(first)
    expect(third).toBeGreaterThan(second)
    expect(xml.indexOf('<c:cat/>')).toBeGreaterThan(third)
  })

  it('targets the series by idx and leaves its own spPr alone', () => {
    const xml = applyChartEdit(BAR_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      pointColors: { 1: { 0: '#FF8800' } },
    })
    expect(xml).toContain(
      '<c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></c:spPr>' +
        '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="FF8800"/></a:solidFill></c:spPr></c:dPt>' +
        '<c:cat/>',
    )
    // series 0 gains no dPt
    expect(xml.match(/<c:dPt>/g)).toHaveLength(1)
  })

  it('fails closed on a missing series index', () => {
    expect(() =>
      applyChartEdit(BAR_CHART, {
        chartPath: 'xl/charts/chart1.xml',
        pointColors: { 7: { 0: '#00AA55' } },
      }),
    ).toThrow(ChartEditError)
  })
})

describe('applyChartEdit series colors with dPt present', () => {
  it('inserts the series spPr without clobbering a point spPr', () => {
    const xml = applyChartEdit(PIE_WITH_DPT, {
      chartPath: 'xl/charts/chart1.xml',
      seriesColors: { 0: '#0055AA' },
    })
    expect(xml).toContain(
      '<c:order val="0"/><c:spPr><a:solidFill><a:srgbClr val="0055AA"/></a:solidFill>',
    )
    expect(xml).toContain('<a:srgbClr val="ABCDEF"/>')
  })
})

const AXIS_CHART =
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<c:chart><c:plotArea><c:layout/>' +
  '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>' +
  '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>' +
  '<c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/>' +
  '<c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt></c:strCache></c:strRef></c:cat>' +
  '<c:val><c:numRef><c:f>Data!$B$2:$B$3</c:f><c:numCache><c:ptCount val="2"/>' +
  '<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>' +
  '</c:ser>' +
  '<c:axId val="111"/><c:axId val="222"/></c:barChart>' +
  '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
  '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>' +
  '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
  '<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>' +
  '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>old axis</a:t></a:r></a:p></c:rich></c:tx></c:title>' +
  '<c:crossAx val="111"/></c:valAx>' +
  '</c:plotArea>' +
  '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>' +
  '<c:plotVisOnly val="1"/></c:chart></c:chartSpace>'

describe('applyChartEdit legend', () => {
  it('removes the legend on none', () => {
    const xml = applyChartEdit(AXIS_CHART, { chartPath: 'xl/charts/chart1.xml', legend: 'none' })
    expect(xml).not.toContain('<c:legend>')
  })

  it('moves the legend to the bottom', () => {
    const xml = applyChartEdit(AXIS_CHART, { chartPath: 'xl/charts/chart1.xml', legend: 'bottom' })
    expect(xml).toContain('<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>')
    expect(xml.match(/<c:legend>/g)).toHaveLength(1)
  })

  it('creates a legend on a chart without one', () => {
    const xml = applyChartEdit(BAR_CHART, { chartPath: 'xl/charts/chart1.xml', legend: 'top' })
    expect(xml).toContain('</c:plotArea><c:legend><c:legendPos val="t"/>')
  })
})

describe('applyChartEdit axis titles', () => {
  it('inserts a category-axis title after axPos', () => {
    const xml = applyChartEdit(AXIS_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      axisTitles: { category: 'Month' },
    })
    expect(xml).toContain(
      '<c:axPos val="b"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Month</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:crossAx val="222"/>',
    )
  })

  it('replaces the value-axis title after gridlines and clears it on null', () => {
    const renamed = applyChartEdit(AXIS_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      axisTitles: { value: 'Sales' },
    })
    expect(renamed).toContain(
      '<c:majorGridlines/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:crossAx val="111"/>',
    )
    expect(renamed).not.toContain('old axis')
    const cleared = applyChartEdit(AXIS_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      axisTitles: { value: null },
    })
    expect(cleared).not.toContain('old axis')
    expect(cleared).not.toContain(
      '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:title>',
    )
  })
})

describe('applyChartEdit series data', () => {
  it('rewrites name, values, and categories with refs and caches', () => {
    const xml = applyChartEdit(AXIS_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      series: [
        {
          index: 0,
          name: 'East',
          valuesRef: "'Data'!$C$2:$C$4",
          values: [3, 5, 8],
          categoriesRef: "'Data'!$A$2:$A$4",
          categories: ['Jan', 'Feb', 'Mar'],
        },
      ],
    })
    expect(xml).toContain('<c:tx><c:v>East</c:v></c:tx>')
    expect(xml).toContain("<c:f>'Data'!$C$2:$C$4</c:f>")
    expect(xml).toContain('<c:ptCount val="3"/><c:pt idx="0"><c:v>3</c:v></c:pt>')
    expect(xml).toContain('<c:pt idx="2"><c:v>Mar</c:v></c:pt>')
    expect(xml).not.toContain('Data!$B$2:$B$3')
  })

  it('handles self-closing cat/val placeholders with ref-less literals', () => {
    const xml = applyChartEdit(BAR_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      series: [{ index: 0, values: [1, 2], categories: ['x', 'y'] }],
    })
    // No reference given: strLit/numLit, never a strRef/numRef missing c:f.
    expect(xml).toContain('<c:cat><c:strLit><c:ptCount val="2"/>')
    expect(xml).toContain('<c:val><c:numLit>')
    expect(xml).not.toContain('<c:strRef><c:strCache>')
    expect(xml).not.toContain('<c:numRef><c:numCache>')
  })

  it('fails closed on a missing series index', () => {
    expect(() =>
      applyChartEdit(AXIS_CHART, {
        chartPath: 'xl/charts/chart1.xml',
        series: [{ index: 5, name: 'x' }],
      }),
    ).toThrow(ChartEditError)
  })

  it('keeps numRef categories numeric with their formatCode', () => {
    const dated = AXIS_CHART.replace(
      '<c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/>' +
        '<c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt></c:strCache></c:strRef></c:cat>',
      '<c:cat><c:numRef><c:f>Data!$A$2:$A$3</c:f><c:numCache><c:formatCode>mmm\\-yy</c:formatCode>' +
        '<c:ptCount val="2"/><c:pt idx="0"><c:v>44562</c:v></c:pt><c:pt idx="1"><c:v>44593</c:v></c:pt></c:numCache></c:numRef></c:cat>',
    )
    const xml = applyChartEdit(dated, {
      chartPath: 'xl/charts/chart1.xml',
      series: [
        {
          index: 0,
          categoriesRef: 'Data!$A$2:$A$3',
          categories: ['44562', '44593'],
          valuesRef: 'Data!$B$2:$B$3',
          values: [1, 2],
        },
      ],
    })
    expect(xml).toContain(
      '<c:cat><c:numRef><c:f>Data!$A$2:$A$3</c:f><c:numCache><c:formatCode>mmm\\-yy</c:formatCode>' +
        '<c:ptCount val="2"/><c:pt idx="0"><c:v>44562</c:v></c:pt>',
    )
    expect(xml).not.toContain('<c:strRef><c:f>Data!$A$2:$A$3')
    // Ref-less numeric categories become a numLit (c:f is mandatory in numRef).
    const literal = applyChartEdit(dated, {
      chartPath: 'xl/charts/chart1.xml',
      series: [{ index: 0, categories: ['44562', '44593'] }],
    })
    expect(literal).toContain('<c:cat><c:numLit><c:formatCode>mmm\\-yy</c:formatCode>')
    // Text categories still rewrite to the string shape.
    const texty = applyChartEdit(dated, {
      chartPath: 'xl/charts/chart1.xml',
      series: [{ index: 0, categories: ['Jan', 'Feb'], categoriesRef: 'Data!$A$2:$A$3' }],
    })
    expect(texty).toContain('<c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f><c:strCache>')
    expect(texty).not.toContain('<c:cat><c:numRef>')
  })

  it('keeps the original number format when rewriting values', () => {
    const formatted = AXIS_CHART.replace(
      '<c:numCache><c:ptCount val="2"/>',
      '<c:numCache><c:formatCode>#,##0.00</c:formatCode><c:ptCount val="2"/>',
    )
    const xml = applyChartEdit(formatted, {
      chartPath: 'xl/charts/chart1.xml',
      series: [{ index: 0, valuesRef: "'Data'!$C$2:$C$3", values: [7, 9] }],
    })
    expect(xml).toContain('<c:formatCode>#,##0.00</c:formatCode><c:ptCount val="2"/>')
    expect(xml).not.toContain('<c:formatCode>General</c:formatCode>')
  })
})

describe('applyChartEdit gridlines', () => {
  const NO_GRID = AXIS_CHART.replace('<c:majorGridlines/>', '')

  it('true inserts majorGridlines after axPos on the value axis only', () => {
    const xml = applyChartEdit(NO_GRID, { chartPath: 'xl/charts/chart1.xml', gridlines: true })
    expect(xml).toContain('<c:axPos val="l"/><c:majorGridlines/><c:title>')
    expect(xml.match(/<c:majorGridlines\/>/g)).toHaveLength(1)
    expect(xml).toContain('<c:axPos val="b"/><c:crossAx val="222"/>')
  })

  it('true keeps an already-present element untouched', () => {
    expect(applyChartEdit(AXIS_CHART, { chartPath: 'xl/charts/chart1.xml', gridlines: true })).toBe(
      AXIS_CHART,
    )
  })

  it('false removes the element, expanded form included', () => {
    const removed = applyChartEdit(AXIS_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      gridlines: false,
    })
    expect(removed).not.toContain('majorGridlines')
    expect(removed).toContain('<c:axPos val="l"/><c:title>')
    const expanded = AXIS_CHART.replace(
      '<c:majorGridlines/>',
      '<c:majorGridlines><c:spPr><a:ln/></c:spPr></c:majorGridlines>',
    )
    expect(
      applyChartEdit(expanded, { chartPath: 'xl/charts/chart1.xml', gridlines: false }),
    ).not.toContain('majorGridlines')
  })

  it('fails closed without a value axis', () => {
    expect(() =>
      applyChartEdit(UNTITLED_PIE, { chartPath: 'xl/charts/chart1.xml', gridlines: true }),
    ).toThrow(ChartEditError)
  })
})

describe('applyChartEdit value-axis bounds', () => {
  it('writes max then min into the scaling, in schema order', () => {
    const xml = applyChartEdit(AXIS_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      valueAxis: { min: 2, max: 10.5 },
    })
    expect(xml).toContain(
      '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/>' +
        '<c:max val="10.5"/><c:min val="2"/></c:scaling>',
    )
    // category-axis scaling untouched
    expect(xml).toContain(
      '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling>',
    )
  })

  it('replaces existing bounds and null resets a bound to auto', () => {
    const bounded = applyChartEdit(AXIS_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      valueAxis: { min: 0, max: 100 },
    })
    const changed = applyChartEdit(bounded, {
      chartPath: 'xl/charts/chart1.xml',
      valueAxis: { max: 50 },
    })
    expect(changed).toContain('<c:max val="50"/><c:min val="0"/>')
    const cleared = applyChartEdit(changed, {
      chartPath: 'xl/charts/chart1.xml',
      valueAxis: { min: null, max: null },
    })
    expect(cleared).not.toContain('<c:max')
    expect(cleared).not.toContain('<c:min ')
    expect(cleared).toContain(
      '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling>',
    )
  })

  it('fails closed without a value axis', () => {
    expect(() =>
      applyChartEdit(UNTITLED_PIE, { chartPath: 'xl/charts/chart1.xml', valueAxis: { max: 5 } }),
    ).toThrow(ChartEditError)
  })
})

describe('applyChartEdit gap width', () => {
  it('replaces an existing gapWidth in place', () => {
    const xml = applyChartEdit(BAR_CHART, { chartPath: 'xl/charts/chart1.xml', gapWidthPct: 80 })
    expect(xml).toContain('</c:ser><c:gapWidth val="80"/><c:overlap val="-27"/>')
    expect(xml).not.toContain('<c:gapWidth val="150"/>')
  })

  it('inserts after the last series when missing', () => {
    const xml = applyChartEdit(AXIS_CHART, { chartPath: 'xl/charts/chart1.xml', gapWidthPct: 40 })
    expect(xml).toContain('</c:ser><c:gapWidth val="40"/><c:axId val="111"/>')
  })

  it('lands after plot-level labels added in the same edit', () => {
    const xml = applyChartEdit(AXIS_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      dataLabels: 'value',
      gapWidthPct: 40,
    })
    expect(xml).toContain('</c:dLbls><c:gapWidth val="40"/><c:axId val="111"/>')
  })

  it('fails closed without a bar plot', () => {
    expect(() =>
      applyChartEdit(UNTITLED_PIE, { chartPath: 'xl/charts/chart1.xml', gapWidthPct: 80 }),
    ).toThrow(ChartEditError)
  })
})

describe('applyChartEdit hole size', () => {
  it('replaces the hole size on a doughnut (same edit as a conversion)', () => {
    const xml = applyChartEdit(UNTITLED_PIE, {
      chartPath: 'xl/charts/chart1.xml',
      chartType: 'doughnut',
      holeSizePct: 70,
    })
    expect(xml).toContain('<c:firstSliceAng val="0"/><c:holeSize val="70"/></c:doughnutChart>')
    expect(xml).not.toContain('<c:holeSize val="50"/>')
  })

  it('inserts before the plot close when missing', () => {
    const doughnut = applyChartEdit(UNTITLED_PIE, {
      chartPath: 'xl/charts/chart1.xml',
      chartType: 'doughnut',
    }).replace('<c:holeSize val="50"/>', '')
    const xml = applyChartEdit(doughnut, { chartPath: 'xl/charts/chart1.xml', holeSizePct: 30 })
    expect(xml).toContain('<c:holeSize val="30"/></c:doughnutChart>')
  })

  it('fails closed on a pie', () => {
    expect(() =>
      applyChartEdit(UNTITLED_PIE, { chartPath: 'xl/charts/chart1.xml', holeSizePct: 40 }),
    ).toThrow(ChartEditError)
  })
})

describe('applyChartEdit series explosion', () => {
  it('inserts before the data, and 0 is written too', () => {
    const xml = applyChartEdit(UNTITLED_PIE, {
      chartPath: 'xl/charts/chart1.xml',
      explosionPct: 25,
    })
    expect(xml).toContain('<c:order val="0"/><c:explosion val="25"/><c:cat/>')
    const zero = applyChartEdit(UNTITLED_PIE, {
      chartPath: 'xl/charts/chart1.xml',
      explosionPct: 0,
    })
    expect(zero).toContain('<c:explosion val="0"/>')
  })

  it('inserts before dPt entries, leaving their own explosion alone', () => {
    const xml = applyChartEdit(PIE_WITH_DPT, {
      chartPath: 'xl/charts/chart1.xml',
      explosionPct: 10,
    })
    expect(xml).toContain('<c:order val="0"/><c:explosion val="10"/><c:dPt>')
    expect(xml).toContain('<c:explosion val="5"/>')
  })

  it('replaces an existing series-level explosion', () => {
    const once = applyChartEdit(PIE_WITH_DPT, {
      chartPath: 'xl/charts/chart1.xml',
      explosionPct: 10,
    })
    const twice = applyChartEdit(once, { chartPath: 'xl/charts/chart1.xml', explosionPct: 20 })
    expect(twice).toContain('<c:order val="0"/><c:explosion val="20"/><c:dPt>')
    expect(twice).toContain('<c:explosion val="5"/>')
    expect(twice).not.toContain('<c:explosion val="10"/>')
  })

  it('fails closed on a bar chart', () => {
    expect(() =>
      applyChartEdit(BAR_CHART, { chartPath: 'xl/charts/chart1.xml', explosionPct: 25 }),
    ).toThrow(ChartEditError)
  })
})

describe('applyChartEdit point explosions', () => {
  it('updates the explosion of an existing dPt, keeping its fill', () => {
    const xml = applyChartEdit(PIE_WITH_DPT, {
      chartPath: 'xl/charts/chart1.xml',
      pointExplosions: { 1: 12 },
    })
    expect(xml).toContain('<c:idx val="1"/><c:bubble3D val="0"/><c:explosion val="12"/>')
    expect(xml).toContain('ABCDEF')
    expect(xml).not.toContain('<c:explosion val="5"/>')
    expect(xml.match(/<c:dPt>/g)).toHaveLength(1)
  })

  it('creates a minimal dPt for a new idx, in ascending order', () => {
    const xml = applyChartEdit(PIE_WITH_DPT, {
      chartPath: 'xl/charts/chart1.xml',
      pointExplosions: { 0: 8 },
    })
    expect(xml).toContain(
      '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:explosion val="8"/></c:dPt><c:dPt><c:idx val="1"/>',
    )
  })

  it('inserts the explosion before spPr on a dPt that has only a color', () => {
    const xml = applyChartEdit(UNTITLED_PIE, {
      chartPath: 'xl/charts/chart1.xml',
      pointColors: { 0: { 0: '#ff0000' } },
      pointExplosions: { 0: 15 },
    })
    expect(xml).toContain(
      '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:explosion val="15"/>' +
        '<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt>',
    )
  })

  it('fails closed on a bar chart', () => {
    expect(() =>
      applyChartEdit(BAR_CHART, { chartPath: 'xl/charts/chart1.xml', pointExplosions: { 0: 10 } }),
    ).toThrow(ChartEditError)
  })
})

describe('applyChartEdit seriesSet', () => {
  it('rebuilds the series in order with refs, caches, and colors', () => {
    const xml = applyChartEdit(AXIS_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      seriesSet: [
        {
          name: 'East',
          values: [1, 2],
          valuesRef: "'D'!$B$2:$B$3",
          categories: ['x', 'y'],
          categoriesRef: "'D'!$A$2:$A$3",
          color: '#112233',
        },
        { name: 'North', values: [3, 4] },
      ],
    })
    expect(xml.match(/<c:ser>/g)).toHaveLength(2)
    expect(xml).toContain(
      '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>East</c:v></c:tx>' +
        '<c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></c:spPr>' +
        '<c:cat><c:strRef><c:f>\'D\'!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/>',
    )
    expect(xml).toContain(
      "<c:val><c:numRef><c:f>'D'!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode>" +
        '<c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt>',
    )
    // ref-less series: literal values, no c:cat at all
    expect(xml).toContain(
      '<c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:v>North</c:v></c:tx>' +
        '<c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt>',
    )
    // old series (its tx ref and data) is gone
    expect(xml).not.toContain('Data!$B$1')
    expect(xml).not.toContain('Data!$B$2:$B$3')
  })

  it('keeps the other plot children in place', () => {
    const xml = applyChartEdit(BAR_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      seriesSet: [{ name: 'S', values: [1] }],
    })
    expect(xml.match(/<c:ser>/g)).toHaveLength(1)
    expect(xml).toContain(
      '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser>',
    )
    expect(xml).toContain(
      '</c:ser><c:gapWidth val="150"/><c:overlap val="-27"/><c:axId val="111"/>',
    )
    // old per-series leftovers dropped with their series
    expect(xml).not.toContain('invertIfNegative')
    expect(xml).not.toContain('112233')
  })

  it('later index-keyed edits target the new series', () => {
    const xml = applyChartEdit(AXIS_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      seriesSet: [
        { name: 'A', values: [1] },
        { name: 'B', values: [2] },
      ],
      seriesColors: { 1: '#00AA55' },
    })
    expect(xml).toContain(
      '<c:tx><c:v>B</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="00AA55"/></a:solidFill>',
    )
  })

  it('fails closed on combo and series-less plots', () => {
    const combo =
      '<c:chartSpace><c:chart><c:plotArea>' +
      '<c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:cat/><c:val/></c:ser></c:barChart>' +
      '<c:lineChart><c:ser><c:idx val="1"/><c:order val="1"/><c:cat/><c:val/></c:ser></c:lineChart>' +
      '</c:plotArea></c:chart></c:chartSpace>'
    expect(() =>
      applyChartEdit(combo, {
        chartPath: 'xl/charts/chart1.xml',
        seriesSet: [{ name: 'S', values: [1] }],
      }),
    ).toThrow(ChartEditError)
    const empty = UNTITLED_PIE.replace(/<c:ser>[\s\S]*?<\/c:ser>/, '')
    expect(() =>
      applyChartEdit(empty, {
        chartPath: 'xl/charts/chart1.xml',
        seriesSet: [{ name: 'S', values: [1] }],
      }),
    ).toThrow(ChartEditError)
  })
})

const SCATTER_CHART =
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<c:chart><c:plotArea><c:layout/>' +
  '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>' +
  '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>S1</c:v></c:tx>' +
  '<c:xVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt>' +
  '<c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:xVal>' +
  '<c:yVal><c:numRef><c:f>Data!$B$2:$B$3</c:f><c:numCache><c:formatCode>0.00</c:formatCode>' +
  '<c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt>' +
  '</c:numCache></c:numRef></c:yVal>' +
  '<c:smooth val="0"/></c:ser>' +
  '<c:axId val="111"/><c:axId val="222"/></c:scatterChart>' +
  '<c:valAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
  '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:valAx>' +
  '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
  '<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/></c:valAx>' +
  '</c:plotArea></c:chart></c:chartSpace>'

describe('applyChartEdit scatter series data', () => {
  it('rewrites name, X values, and Y values with refs and caches', () => {
    const xml = applyChartEdit(SCATTER_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      series: [
        {
          index: 0,
          name: 'Actual',
          categoriesRef: "'Data'!$A$2:$A$4",
          categories: ['1.5', '2.5', '4'],
          valuesRef: "'Data'!$C$2:$C$4",
          values: [7, 8, 9],
        },
      ],
    })
    expect(xml).toContain('<c:tx><c:v>Actual</c:v></c:tx>')
    expect(xml).toContain(
      "<c:xVal><c:numRef><c:f>'Data'!$A$2:$A$4</c:f><c:numCache>" +
        '<c:formatCode>General</c:formatCode><c:ptCount val="3"/>' +
        '<c:pt idx="0"><c:v>1.5</c:v></c:pt><c:pt idx="1"><c:v>2.5</c:v></c:pt>' +
        '<c:pt idx="2"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:xVal>',
    )
    // yVal keeps the original number format and no c:cat/c:val appears.
    expect(xml).toContain(
      "<c:yVal><c:numRef><c:f>'Data'!$C$2:$C$4</c:f><c:numCache><c:formatCode>0.00</c:formatCode>" +
        '<c:ptCount val="3"/><c:pt idx="0"><c:v>7</c:v></c:pt>',
    )
    expect(xml).not.toContain('<c:cat>')
    expect(xml).not.toContain('<c:val>')
    expect(xml).not.toContain('Data!$B$2:$B$3')
  })

  it('non-numeric X entries fall back to their own index', () => {
    const xml = applyChartEdit(SCATTER_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      series: [{ index: 0, categories: ['a', '5', ' '], values: [1, 2, 3] }],
    })
    expect(xml).toContain(
      '<c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt>' +
        '<c:pt idx="2"><c:v>2</c:v></c:pt></c:numLit></c:xVal>',
    )
  })

  it('values alone rewrite only yVal', () => {
    const xml = applyChartEdit(SCATTER_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      series: [{ index: 0, values: [10, 20] }],
    })
    expect(xml).toContain('<c:xVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt>')
    expect(xml).toContain(
      '<c:yVal><c:numLit><c:formatCode>0.00</c:formatCode>' +
        '<c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt>',
    )
  })

  it('inserts xVal before yVal on a series that lacks one', () => {
    const bare = SCATTER_CHART.replace(/<c:xVal>[\s\S]*?<\/c:xVal>/, '')
    const xml = applyChartEdit(bare, {
      chartPath: 'xl/charts/chart1.xml',
      series: [{ index: 0, categories: ['1', '2'], values: [1, 2] }],
    })
    expect(/<c:xVal>[\s\S]*?<\/c:xVal><c:yVal>/.test(xml)).toBe(true)
  })
})

describe('applyChartEdit scatter seriesSet', () => {
  it('rebuilds scatter series with xVal/yVal, smooth 0, and line color', () => {
    const xml = applyChartEdit(SCATTER_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      seriesSet: [
        {
          name: 'A',
          values: [3, 4],
          valuesRef: "'D'!$B$2:$B$3",
          categories: ['1', '2'],
          categoriesRef: "'D'!$A$2:$A$3",
          color: '#112233',
        },
        { name: 'B', values: [5, 6], categories: ['x', 'y'] },
      ],
    })
    expect(xml.match(/<c:ser>/g)).toHaveLength(2)
    expect(xml).toContain(
      '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>A</c:v></c:tx>' +
        '<c:spPr><a:ln><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></c:spPr>' +
        "<c:xVal><c:numRef><c:f>'D'!$A$2:$A$3</c:f><c:numCache><c:formatCode>General</c:formatCode>" +
        '<c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt>',
    )
    expect(xml).toContain(
      "<c:yVal><c:numRef><c:f>'D'!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode>" +
        '<c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt>',
    )
    // Non-numeric categories degrade to ordinal X values without a ref.
    expect(xml).toContain(
      '<c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:v>B</c:v></c:tx>' +
        '<c:xVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>0</c:v></c:pt>' +
        '<c:pt idx="1"><c:v>1</c:v></c:pt></c:numLit></c:xVal>' +
        '<c:yVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt>',
    )
    expect(xml.match(/<c:smooth val="0"\/>/g)).toHaveLength(2)
    expect(xml).toContain('<c:axId val="111"/><c:axId val="222"/></c:scatterChart>')
    expect(xml).not.toContain('S1')
  })

  it('the rebuilt series stay editable (roundtrip)', () => {
    const replaced = applyChartEdit(SCATTER_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      seriesSet: [{ name: 'A', values: [3, 4], categories: ['1', '2'] }],
    })
    const edited = applyChartEdit(replaced, {
      chartPath: 'xl/charts/chart1.xml',
      series: [{ index: 0, name: 'A2', values: [9, 9], categories: ['7', '8'] }],
      seriesColors: { 0: '#00AA55' },
    })
    expect(edited).toContain('<c:tx><c:v>A2</c:v></c:tx>')
    expect(edited).toContain('<c:pt idx="0"><c:v>7</c:v></c:pt>')
    expect(edited).toContain('<c:pt idx="0"><c:v>9</c:v></c:pt>')
    expect(edited).toContain('<a:srgbClr val="00AA55"/>')
    expect(edited).toContain('<c:smooth val="0"/>')
  })

  it('scatter still refuses type conversion', () => {
    expect(() =>
      applyChartEdit(SCATTER_CHART, { chartPath: 'xl/charts/chart1.xml', chartType: 'line' }),
    ).toThrow(ChartEditError)
  })
})

describe('applyChartEdit data label position and format', () => {
  it('creates value labels carrying position and numFmt when none exist', () => {
    const xml = applyChartEdit(BAR_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      dataLabelPosition: 'outside-end',
      dataLabelFormat: '0.0%',
    })
    expect(xml).toContain(
      '</c:ser><c:dLbls><c:numFmt formatCode="0.0%" sourceLinked="0"/><c:dLblPos val="outEnd"/>' +
        '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>' +
        '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>',
    )
  })

  it('upgrades labels created in the same edit', () => {
    const xml = applyChartEdit(UNTITLED_PIE, {
      chartPath: 'xl/charts/chart1.xml',
      dataLabels: 'percent',
      dataLabelPosition: 'center',
      dataLabelFormat: '#,##0',
    })
    expect(xml).toContain(
      '<c:dLbls><c:numFmt formatCode="#,##0" sourceLinked="0"/><c:dLblPos val="ctr"/>' +
        '<c:showLegendKey val="0"/><c:showVal val="0"/>',
    )
    expect(xml).toContain('<c:showPercent val="1"/>')
  })

  it('replaces an existing position and inserts a numFmt into existing labels', () => {
    const once = applyChartEdit(BAR_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      dataLabels: 'value',
      dataLabelPosition: 'center',
    })
    const twice = applyChartEdit(once, {
      chartPath: 'xl/charts/chart1.xml',
      dataLabelPosition: 'inside-end',
      dataLabelFormat: '0.00',
    })
    expect(twice).toContain(
      '<c:dLbls><c:numFmt formatCode="0.00" sourceLinked="0"/><c:dLblPos val="inEnd"/><c:showLegendKey',
    )
    expect(twice).not.toContain('<c:dLblPos val="ctr"/>')
    expect(twice.match(/<c:dLbls>/g)).toHaveLength(1)
  })

  it('none wins over a position/format sent in the same batch', () => {
    const labeled = applyChartEdit(BAR_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      dataLabels: 'value',
      dataLabelPosition: 'center',
    })
    const cleared = applyChartEdit(labeled, {
      chartPath: 'xl/charts/chart1.xml',
      dataLabels: 'none',
      dataLabelPosition: 'outside-end',
      dataLabelFormat: '0%',
    })
    expect(cleared).not.toContain('<c:dLbls>')
    expect(cleared).not.toContain('dLblPos')
  })

  it('escapes quotes in the format code attribute', () => {
    const xml = applyChartEdit(BAR_CHART, {
      chartPath: 'xl/charts/chart1.xml',
      dataLabelFormat: '0.0" kg"',
    })
    expect(xml).toContain('<c:numFmt formatCode="0.0&quot; kg&quot;" sourceLinked="0"/>')
  })
})
