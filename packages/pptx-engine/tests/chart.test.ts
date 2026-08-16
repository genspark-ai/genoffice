import { describe, it, expect } from 'vitest'
import { parseChartXml } from '../src/chart'
import { buildChartSpaceXml } from '../src/chart-insert'

const LINE_CHART = `<?xml version="1.0"?><c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:lineChart><c:ser>
  <c:idx val="0"/><c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
  <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="4CAF50"/></a:solidFill></a:ln></c:spPr>
  <c:marker><c:symbol val="circle"/></c:marker>
  <c:cat><c:strRef><c:f>S!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>
    <c:pt idx="0"><c:v>2016</c:v></c:pt><c:pt idx="1"><c:v>2017</c:v></c:pt><c:pt idx="2"><c:v>2018</c:v></c:pt></c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:f>S!$B$2:$B$4</c:f><c:numCache><c:ptCount val="3"/>
    <c:pt idx="0"><c:v>100</c:v></c:pt><c:pt idx="1"><c:v>120</c:v></c:pt><c:pt idx="2"><c:v>145</c:v></c:pt></c:numCache></c:numRef></c:val>
  <c:smooth val="1"/>
</c:ser></c:lineChart>
<c:valAx><c:scaling><c:orientation val="minMax"/></c:scaling><c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="E6E6E6"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr></c:majorGridlines>
  <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1200"><a:solidFill><a:srgbClr val="666666"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr></c:valAx>
</c:plotArea><c:legend><c:legendPos val="t"/></c:legend></c:chart></c:chartSpace>`

const BAR_CHART = `<?xml version="1.0"?><c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
<c:ser><c:idx val="0"/>
  <c:spPr><a:solidFill><a:srgbClr val="629BF8"/></a:solidFill></c:spPr>
  <c:cat><c:strRef><c:f>x</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>AAPL</c:v></c:pt><c:pt idx="1"><c:v>MSFT</c:v></c:pt></c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>30.5</c:v></c:pt><c:pt idx="1"><c:v>35</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser><c:gapWidth val="28"/></c:barChart>
</c:plotArea></c:chart></c:chartSpace>`

describe('parseChartXml', () => {
  it('parses line chart: series name/color/smooth/marker + categories + values + legend + axis', () => {
    const m = parseChartXml(LINE_CHART)!
    expect(m.kind).toBe('line')
    expect(m.series).toHaveLength(1)
    expect(m.series[0]!.name).toBe('Revenue')
    expect(m.series[0]!.color).toBe('#4CAF50')
    expect(m.series[0]!.smooth).toBe(true)
    expect(m.series[0]!.marker).toBe(true)
    expect(m.series[0]!.values).toEqual([100, 120, 145])
    expect(m.categories).toEqual(['2016', '2017', '2018'])
    expect(m.legendPos).toBe('t')
    expect(m.valAxis?.gridColor).toBe('#E6E6E6')
    expect(m.valAxis?.gridDash).toBe(true)
    expect(m.valAxis?.labelColor).toBe('#666666')
    expect(m.valAxis?.labelSizePt).toBe(12)
  })

  it('parses clustered bar chart with fill color and gapWidth', () => {
    const m = parseChartXml(BAR_CHART)!
    expect(m.kind).toBe('bar')
    expect(m.barDir).toBe('col')
    expect(m.grouping).toBe('clustered')
    expect(m.gapWidthPct).toBe(28)
    expect(m.series[0]!.color).toBe('#629BF8')
    expect(m.series[0]!.values).toEqual([30.5, 35])
    expect(m.legendPos).toBeUndefined()
  })

  it('returns null for chart space without recognizable plot', () => {
    expect(parseChartXml('<c:chartSpace><c:chart/></c:chartSpace>')).toBeNull()
  })

  it('parses doughnut hole size + per-point colors (c:dPt)', () => {
    const DOUGHNUT = `<?xml version="1.0"?><c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:doughnutChart><c:ser><c:idx val="0"/>
  <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="AA0000"/></a:solidFill></c:spPr></c:dPt>
  <c:dPt><c:idx val="2"/><c:spPr><a:solidFill><a:srgbClr val="0000AA"/></a:solidFill></c:spPr></c:dPt>
  <c:cat><c:strRef><c:f>x</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser><c:firstSliceAng val="90"/><c:holeSize val="60"/></c:doughnutChart>
</c:plotArea></c:chart></c:chartSpace>`
    const m = parseChartXml(DOUGHNUT)!
    expect(m.kind).toBe('pie')
    expect(m.holePct).toBe(60)
    expect(m.firstSliceAngDeg).toBe(90)
    expect(m.series[0]!.pointColors?.[0]).toBe('#AA0000')
    expect(m.series[0]!.pointColors?.[1]).toBeUndefined()
    expect(m.series[0]!.pointColors?.[2]).toBe('#0000AA')
  })

  it('parses pie explosion: series-level c:explosion and per-point c:dPt overrides', () => {
    const PIE = `<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea>
<c:pie3DChart><c:ser><c:idx val="0"/>
  <c:explosion val="25"/>
  <c:dPt><c:idx val="1"/><c:explosion val="10"/></c:dPt>
  <c:val><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser></c:pie3DChart>
</c:plotArea></c:chart></c:chartSpace>`
    const m = parseChartXml(PIE)!
    expect(m.series[0]!.explosionPct).toBe(25)
    expect(m.series[0]!.pointExplosionPct?.[0]).toBeUndefined()
    expect(m.series[0]!.pointExplosionPct?.[1]).toBe(10)
  })

  it('pie chart without holeSize defaults holePct 0', () => {
    const PIE = `<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea>
<c:pieChart><c:ser><c:idx val="0"/>
  <c:val><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser></c:pieChart>
</c:plotArea></c:chart></c:chartSpace>`
    const m = parseChartXml(PIE)!
    expect(m.kind).toBe('pie')
    expect(m.holePct).toBe(0)
  })

  it('parses stacked column grouping', () => {
    const STACKED = BAR_CHART.replace('val="clustered"', 'val="stacked"')
    const m = parseChartXml(STACKED)!
    expect(m.grouping).toBe('stacked')
  })

  it('parses horizontal bar (barDir=bar) with reversed category axis', () => {
    const HBAR = BAR_CHART.replace('val="col"', 'val="bar"').replace(
      '</c:plotArea>',
      '<c:catAx><c:scaling><c:orientation val="maxMin"/></c:scaling></c:catAx></c:plotArea>',
    )
    const m = parseChartXml(HBAR)!
    expect(m.barDir).toBe('bar')
    expect(m.catAxis?.reversed).toBe(true)
  })

  it('parses scatter chart: xVal/yVal + scatterStyle + dual value-axis dispatch', () => {
    const SCATTER = `<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea>
<c:scatterChart><c:scatterStyle val="lineMarker"/>
<c:ser><c:idx val="0"/>
  <c:tx><c:strRef><c:f>t</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Sample</c:v></c:pt></c:strCache></c:strRef></c:tx>
  <c:marker><c:symbol val="none"/></c:marker>
  <c:xVal><c:numRef><c:f>x</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>1.5</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:xVal>
  <c:yVal><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>15</c:v></c:pt></c:numCache></c:numRef></c:yVal>
</c:ser></c:scatterChart>
<c:valAx><c:axPos val="b"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling></c:valAx>
<c:valAx><c:axPos val="l"/><c:majorGridlines/></c:valAx>
</c:plotArea></c:chart></c:chartSpace>`
    const m = parseChartXml(SCATTER)!
    expect(m.kind).toBe('scatter')
    expect(m.scatterStyle).toBe('lineMarker')
    expect(m.series[0]!.name).toBe('Sample')
    expect(m.series[0]!.xValues).toEqual([1.5, 2, 4])
    expect(m.series[0]!.values).toEqual([10, 20, 15])
    expect(m.series[0]!.marker).toBe(false) // explicit symbol=none
    // The value axis with axPos=b goes into the catAxis slot (x axis); the other goes into valAxis
    expect(m.catAxis?.min).toBe(0)
    expect(m.catAxis?.max).toBe(5)
    expect(m.valAxis?.gridColor).toBe('#E6E6E6')
  })

  it('parses radar chart: radarStyle + categories', () => {
    const RADAR = `<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea>
<c:radarChart><c:radarStyle val="filled"/>
<c:ser><c:idx val="0"/>
  <c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></c:spPr>
  <c:cat><c:strRef><c:f>x</c:f><c:strCache><c:ptCount val="4"/><c:pt idx="0"><c:v>Listen</c:v></c:pt><c:pt idx="1"><c:v>Speak</c:v></c:pt><c:pt idx="2"><c:v>Read</c:v></c:pt><c:pt idx="3"><c:v>Write</c:v></c:pt></c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="4"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>5</c:v></c:pt><c:pt idx="3"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser></c:radarChart>
</c:plotArea></c:chart></c:chartSpace>`
    const m = parseChartXml(RADAR)!
    expect(m.kind).toBe('radar')
    expect(m.radarStyle).toBe('filled')
    expect(m.categories).toEqual(['Listen', 'Speak', 'Read', 'Write'])
    expect(m.series[0]!.color).toBe('#ED7D31')
    expect(m.series[0]!.values).toEqual([3, 4, 5, 2])
  })

  it('parses bar+line combo: both plots kept, series tagged with plotKind', () => {
    const COMBO = `<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
<c:ser><c:idx val="0"/>
  <c:tx><c:strRef><c:f>n</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx>
  <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>
  <c:cat><c:strRef><c:f>x</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser><c:gapWidth val="150"/></c:barChart>
<c:lineChart><c:ser><c:idx val="1"/>
  <c:tx><c:strRef><c:f>n2</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Growth Rate</c:v></c:pt></c:strCache></c:strRef></c:tx>
  <c:spPr><a:ln><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></a:ln></c:spPr>
  <c:val><c:numRef><c:f>y2</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>8</c:v></c:pt><c:pt idx="2"><c:v>12</c:v></c:pt></c:numCache></c:numRef></c:val>
  <c:smooth val="1"/>
</c:ser></c:lineChart>
</c:plotArea></c:chart></c:chartSpace>`
    const m = parseChartXml(COMBO)!
    expect(m.kind).toBe('bar') // primary type: bar > line
    expect(m.barDir).toBe('col')
    expect(m.series).toHaveLength(2)
    expect(m.series[0]!.name).toBe('Sales')
    expect(m.series[0]!.plotKind).toBe('bar')
    expect(m.series[0]!.values).toEqual([10, 20, 30])
    expect(m.series[1]!.name).toBe('Growth Rate')
    expect(m.series[1]!.plotKind).toBe('line')
    expect(m.series[1]!.values).toEqual([5, 8, 12])
    expect(m.series[1]!.smooth).toBe(true)
    expect(m.categories).toEqual(['Q1', 'Q2', 'Q3'])
  })

  it('single-plot chart keeps plotKind undefined (non-combo unchanged)', () => {
    const m = parseChartXml(BAR_CHART)!
    expect(m.series[0]!.plotKind).toBeUndefined()
  })

  it('combo dual axis: secondary value axis (axPos=r) matched via plot axId → series assignment + min/max of both axes', () => {
    const COMBO_DUAL = `<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
<c:ser><c:idx val="0"/>
  <c:tx><c:strRef><c:f>n</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
  <c:cat><c:strRef><c:f>x</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>120</c:v></c:pt><c:pt idx="1"><c:v>180</c:v></c:pt><c:pt idx="2"><c:v>240</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser><c:axId val="111111111"/><c:axId val="222222222"/></c:barChart>
<c:lineChart><c:ser><c:idx val="1"/>
  <c:tx><c:strRef><c:f>n2</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Growth</c:v></c:pt></c:strCache></c:strRef></c:tx>
  <c:val><c:numRef><c:f>y2</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser><c:axId val="333333333"/><c:axId val="444444444"/></c:lineChart>
<c:catAx><c:axId val="111111111"/><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>
<c:valAx><c:axId val="222222222"/><c:scaling><c:min val="0"/><c:max val="250"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>
<c:valAx><c:axId val="444444444"/><c:scaling><c:min val="0"/><c:max val="30"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:crossAx val="333333333"/><c:crosses val="max"/></c:valAx>
<c:catAx><c:axId val="333333333"/><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="444444444"/></c:catAx>
</c:plotArea></c:chart></c:chartSpace>`
    const m = parseChartXml(COMBO_DUAL)!
    // Series → axis assignment: bars on primary, line on secondary (lineChart's axId 444444444 matches the axPos=r value axis)
    expect(m.series.map((s) => !!s.secondaryAxis)).toEqual([false, true])
    // Primary/secondary value axes parse their min/max independently
    expect(m.valAxis?.min).toBe(0)
    expect(m.valAxis?.max).toBe(250)
    expect(m.valAxis2?.min).toBe(0)
    expect(m.valAxis2?.max).toBe(30)
    // Category axis is the non-deleted one (the secondary plot's hidden category axis is skipped)
    expect(m.categories).toEqual(['Q1', 'Q2', 'Q3'])
  })
})

describe('buildChartSpaceXml comboBarLine (generate → parse round-trip)', () => {
  const OPTS = {
    kind: 'comboBarLine' as const,
    title: 'Quarterly Results',
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
      { name: 'Sales Amount', values: [4.3, 2.5, 3.5, 4.5] },
      { name: 'Cost', values: [2.4, 1.8, 2.6, 2.8] },
      { name: 'Growth Rate', values: [2.0, 1.2, 2.6, 3.4] },
    ],
    offset: { x: 0, y: 0, cx: 100, cy: 100 },
  }

  it('first N-1 series go to barChart, the last to lineChart, line attached to the right secondary value axis', () => {
    const xml = buildChartSpaceXml(OPTS)
    // Both bars and a line inside the same plotArea
    expect(xml).toContain('<c:barChart>')
    expect(xml).toContain('<c:lineChart>')
    // Bars on the primary axes, line on the secondary (own axId + right value axis crosses=max + hidden category axis)
    expect(xml).toContain('<c:axId val="333333333"/><c:axId val="444444444"/>')
    expect(xml).toMatch(
      /<c:valAx><c:axId val="444444444"\/>.*?<c:axPos val="r"\/>.*?<c:crosses val="max"\/>/,
    )
    expect(xml).toMatch(/<c:catAx><c:axId val="333333333"\/>.*?<c:delete val="1"\/>/)
    // 2 bar series, 1 line series
    expect((xml.match(/<c:barChart>.*?<\/c:barChart>/)![0].match(/<c:ser>/g) ?? []).length).toBe(2)
    expect((xml.match(/<c:lineChart>.*?<\/c:lineChart>/)![0].match(/<c:ser>/g) ?? []).length).toBe(
      1,
    )
  })

  it('round-trip: parses back as combo (primary type bar + plotKind markers), data and categories intact', () => {
    const m = parseChartXml(buildChartSpaceXml(OPTS))!
    expect(m.kind).toBe('bar')
    expect(m.series).toHaveLength(3)
    expect(m.series.map((s) => s.plotKind)).toEqual(['bar', 'bar', 'line'])
    // Series → axis assignment round-trip: bars on primary, line on the right secondary value axis
    expect(m.series.map((s) => !!s.secondaryAxis)).toEqual([false, false, true])
    expect(m.series.map((s) => s.name)).toEqual(['Sales Amount', 'Cost', 'Growth Rate'])
    expect(m.series[0]!.values).toEqual([4.3, 2.5, 3.5, 4.5])
    expect(m.series[2]!.values).toEqual([2.0, 1.2, 2.6, 3.4])
    expect(m.categories).toEqual(['Q1', 'Q2', 'Q3', 'Q4'])
  })

  it('single series degrades to a plain bar chart (no lineChart / secondary axis)', () => {
    const xml = buildChartSpaceXml({ ...OPTS, series: [OPTS.series[0]!] })
    expect(xml).toContain('<c:barChart>')
    expect(xml).not.toContain('<c:lineChart>')
    expect(xml).not.toContain('444444444')
    const m = parseChartXml(xml)!
    expect(m.kind).toBe('bar')
    expect(m.series[0]!.plotKind).toBeUndefined()
  })
})

describe('buildChartSpaceXml style options (generate → parse round-trip)', () => {
  const BASE = {
    kind: 'bar' as const,
    categories: ['A', 'B', 'C'],
    series: [{ name: 'S1', values: [1, 2, 3] }],
    offset: { x: 0, y: 0, cx: 100, cy: 100 },
  }

  it('defaults: legend at bottom, no gridlines, no data labels, no axis titles', () => {
    const m = parseChartXml(buildChartSpaceXml(BASE))!
    expect(m.legendPos).toBe('b')
    expect(m.valAxis?.gridColor).toBeUndefined()
    expect(m.dataLabels).toBeUndefined()
    expect(m.valAxis?.title).toBeUndefined()
  })

  it('legendPos/dataLabels/gridlines/axis titles/gapWidth full round-trip', () => {
    const m = parseChartXml(
      buildChartSpaceXml({
        ...BASE,
        legendPos: 'r',
        dataLabels: true,
        gridlines: true,
        catAxisTitle: 'Quarter',
        valAxisTitle: 'Sales (100M)',
        gapWidthPct: 80,
      }),
    )!
    expect(m.legendPos).toBe('r')
    expect(m.dataLabels).toBe(true)
    expect(m.dataLabelsPct).toBeUndefined()
    expect(m.valAxis?.gridColor).toBeTruthy()
    expect(m.catAxis?.title).toBe('Quarter')
    expect(m.valAxis?.title).toBe('Sales (100M)')
    expect(m.gapWidthPct).toBe(80)
  })

  it('legendPos none writes no c:legend', () => {
    const xml = buildChartSpaceXml({ ...BASE, legendPos: 'none' })
    expect(xml).not.toContain('<c:legend>')
    expect(parseChartXml(xml)!.legendPos).toBeUndefined()
  })

  it('combo charts carry dLbls on both plots; pie/line/scatter supported too', () => {
    for (const kind of ['comboBarLine', 'pie', 'line', 'scatter'] as const) {
      const xml = buildChartSpaceXml({
        ...BASE,
        kind,
        series: [
          { name: 'S1', values: [1, 2, 3] },
          { name: 'S2', values: [4, 5, 6] },
        ],
        dataLabels: true,
      })
      expect(parseChartXml(xml)!.dataLabels, kind).toBe(true)
    }
  })

  it('empty series/category names round-trip as empty: c:tx and c:cat omitted, no placeholders injected', () => {
    const xml = buildChartSpaceXml({
      ...BASE,
      categories: ['', '', ''],
      series: [{ name: '', values: [1, 2, 3] }],
    })
    expect(xml).not.toContain('<c:tx>')
    expect(xml).not.toContain('<c:cat>')
    const m = parseChartXml(xml)!
    expect(m.series[0]!.name).toBeUndefined()
    expect(m.series[0]!.values).toEqual([1, 2, 3])
    expect(m.categories).toEqual(['', '', ''])
  })

  it('partially empty category names keep empty slots (no pt written for empty strings)', () => {
    const xml = buildChartSpaceXml({
      ...BASE,
      categories: ['Q1', '', 'Q3'],
      series: [{ name: 'S', values: [4, 5, 6] }],
    })
    const m = parseChartXml(xml)!
    expect(m.categories).toEqual(['Q1', '', 'Q3'])
    expect(m.series[0]!.name).toBe('S')
  })

  it('dLbls with only showPercent parses as dataLabelsPct (common in external pie charts)', () => {
    const PIE = `<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:pieChart><c:ser><c:idx val="0"/>
  <c:cat><c:strRef><c:f>x</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser><c:dLbls><c:showVal val="0"/><c:showPercent val="1"/></c:dLbls></c:pieChart>
</c:plotArea></c:chart></c:chartSpace>`
    const m = parseChartXml(PIE)!
    expect(m.dataLabels).toBe(true)
    expect(m.dataLabelsPct).toBe(true)
  })
})

describe('chart theme palette and default text size', () => {
  const PIE_V = `<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea>
    <c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>
      <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
      <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
    </c:ser></c:pieChart></c:plotArea></c:chart>
    <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1800"/></a:pPr></a:p></c:txPr>
  </c:chartSpace>`
  const theme = {
    colors: {
      accent1: '#4F81BD',
      accent2: '#C0504D',
      accent3: '#9BBB59',
      accent4: '#8064A2',
      accent5: '#4BACC6',
      accent6: '#F79646',
    },
    clrMap: {},
  } as unknown as import('../src/theme').Theme

  it('exposes theme accents as the default palette', () => {
    const m = parseChartXml(PIE_V, theme)!
    expect(m.themePalette).toEqual([
      '#4F81BD',
      '#C0504D',
      '#9BBB59',
      '#8064A2',
      '#4BACC6',
      '#F79646',
    ])
  })

  it('reads the chartSpace-level default text size', () => {
    const m = parseChartXml(PIE_V, theme)!
    expect(m.defaultTextPt).toBe(18)
  })

  it('leaves both unset without theme/txPr', () => {
    const m = parseChartXml(PIE_V.replace(/<c:txPr>[\s\S]*<\/c:txPr>/, ''))!
    expect(m.themePalette).toBeUndefined()
    expect(m.defaultTextPt).toBeUndefined()
  })
})

describe('3D chart variants flatten to their 2D projection', () => {
  it('pie3DChart parses as pie with data', () => {
    const m = parseChartXml(`<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea>
      <c:pie3DChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>
        <c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>7</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>
      </c:ser></c:pie3DChart></c:plotArea></c:chart></c:chartSpace>`)!
    expect(m.kind).toBe('pie')
    expect(m.holePct).toBe(0)
    expect(m.series[0]!.values).toEqual([7, 3])
    expect(m.categories).toEqual(['A', 'B'])
  })

  it('bar3DChart parses as bar (standard grouping renders clustered)', () => {
    const m = parseChartXml(`<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea>
      <c:bar3DChart><c:barDir val="col"/><c:grouping val="standard"/>
        <c:ser><c:idx val="0"/><c:order val="0"/>
          <c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser></c:bar3DChart></c:plotArea></c:chart></c:chartSpace>`)!
    expect(m.kind).toBe('bar')
    expect(m.barDir).toBe('col')
    expect(m.grouping).toBe('standard')
    expect(m.series[0]!.values).toEqual([1, 2])
  })
})

describe('buildChartSpaceXml 3D kinds (generate → parse round-trip)', () => {
  const data = {
    categories: ['A', 'B', 'C'],
    series: [{ name: 'S1', values: [3, 1, 2] }],
    offset: { x: 0, y: 0, cx: 100, cy: 100 },
  }
  it('pie3D writes c:pie3DChart + c:view3D and parses back as pseudo-3D pie', () => {
    const xml = buildChartSpaceXml({ kind: 'pie3D', ...data })
    expect(xml).toContain('<c:pie3DChart>')
    expect(xml).toContain('<c:view3D><c:rotX val="30"/>')
    const m = parseChartXml(xml)!
    expect(m.kind).toBe('pie')
    expect(m.pseudo3D).toBe(true)
    expect(m.rotXDeg).toBe(30)
    expect(m.series[0]!.values).toEqual([3, 1, 2])
  })
  it('bar3D writes c:bar3DChart with three axes + c:view3D and parses back as pseudo-3D bar', () => {
    const xml = buildChartSpaceXml({ kind: 'bar3D', ...data })
    expect(xml).toContain('<c:bar3DChart>')
    expect(xml).toContain('<c:serAx>')
    expect(xml).toContain('<c:view3D><c:rotX val="15"/>')
    const m = parseChartXml(xml)!
    expect(m.kind).toBe('bar')
    expect(m.barDir).toBe('col')
    expect(m.pseudo3D).toBe(true)
  })
})

describe('bubble / stock / varyColors / default background', () => {
  const BUBBLE_CHART = `<?xml version="1.0"?><c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:bubbleChart><c:varyColors val="0"/><c:ser><c:idx val="0"/>
  <c:xVal><c:numRef><c:f>x</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:xVal>
  <c:yVal><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>15</c:v></c:pt></c:numCache></c:numRef></c:yVal>
  <c:bubbleSize><c:numRef><c:f>s</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>8</c:v></c:pt></c:numCache></c:numRef></c:bubbleSize>
</c:ser><c:bubbleScale val="75"/></c:bubbleChart>
</c:plotArea></c:chart></c:chartSpace>`

  it('bubble chart parses as scatter with per-point sizes and bubbleScale', () => {
    const m = parseChartXml(BUBBLE_CHART)!
    expect(m.kind).toBe('scatter')
    expect(m.series[0]!.xValues).toEqual([1, 2, 3])
    expect(m.series[0]!.values).toEqual([10, 20, 15])
    expect(m.series[0]!.bubbleSizes).toEqual([10, 4, 8])
    expect(m.bubbleScale).toBe(75)
    expect(m.bubbleSizeIsWidth).toBeUndefined()
  })

  it('bubble sizeRepresents=w maps size to diameter', () => {
    const m = parseChartXml(
      BUBBLE_CHART.replace(
        '<c:bubbleScale val="75"/>',
        '<c:bubbleScale val="75"/><c:sizeRepresents val="w"/>',
      ),
    )!
    expect(m.bubbleSizeIsWidth).toBe(true)
  })

  const STOCK_CHART = `<?xml version="1.0"?><c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:stockChart>
<c:ser><c:idx val="0"/><c:tx><c:strRef><c:f>o</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Open</c:v></c:pt></c:strCache></c:strRef></c:tx>
  <c:val><c:numRef><c:f>a</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>30</c:v></c:pt><c:pt idx="1"><c:v>25</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
<c:ser><c:idx val="1"/><c:val><c:numRef><c:f>b</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>40</c:v></c:pt><c:pt idx="1"><c:v>35</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
<c:ser><c:idx val="2"/><c:val><c:numRef><c:f>c</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>20</c:v></c:pt><c:pt idx="1"><c:v>15</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
<c:ser><c:idx val="3"/><c:val><c:numRef><c:f>d</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>35</c:v></c:pt><c:pt idx="1"><c:v>18</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
<c:hiLowLines/><c:upDownBars><c:gapWidth val="120"/><c:upBars/><c:downBars/></c:upDownBars>
</c:stockChart>
</c:plotArea></c:chart></c:chartSpace>`

  it('stock chart parses as line with stock flags and gap width', () => {
    const m = parseChartXml(STOCK_CHART)!
    expect(m.kind).toBe('line')
    expect(m.series).toHaveLength(4)
    expect(m.stock).toEqual({ hiLowLines: true, upDownBars: true, gapWidthPct: 120 })
    expect(m.series.every((s) => s.fromStock)).toBe(true)
  })

  it('bubble series color prefers the fill over the outline', () => {
    const withBorder = BUBBLE_CHART.replace(
      '<c:ser><c:idx val="0"/>',
      '<c:ser><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="4CAF50"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="111111"/></a:solidFill></a:ln></c:spPr>',
    )
    const m = parseChartXml(withBorder)!
    expect(m.series[0]!.color).toBe('#4CAF50')
  })

  const PLAIN_PIE = `<?xml version="1.0"?><c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:pieChart><c:varyColors val="0"/><c:ser><c:idx val="0"/>
  <c:cat><c:strRef><c:f>c</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>B</c:v></c:pt><c:pt idx="1"><c:v>C</c:v></c:pt></c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:f>v</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser></c:pieChart>
</c:plotArea></c:chart></c:chartSpace>`

  it('pie varyColors=0 is exposed (all wedges use the series color)', () => {
    const m = parseChartXml(PLAIN_PIE)!
    expect(m.varyColors).toBe(false)
  })

  it('pie without varyColors stays multi-color (flag unset)', () => {
    const m = parseChartXml(PLAIN_PIE.replace('<c:varyColors val="0"/>', ''))!
    expect(m.varyColors).toBeUndefined()
  })
})

describe('date axis and stacked area', () => {
  const AREA = `<?xml version="1.0"?><c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:areaChart><c:grouping val="stacked"/><c:ser><c:idx val="0"/>
  <c:cat><c:numRef><c:f>x</c:f><c:numCache><c:formatCode>m/d/yyyy</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>37261</c:v></c:pt><c:pt idx="1"><c:v>37262</c:v></c:pt></c:numCache></c:numRef></c:cat>
  <c:val><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser></c:areaChart>
<c:dateAx><c:delete val="0"/><c:tickLblPos val="nextTo"/></c:dateAx>
</c:plotArea></c:chart></c:chartSpace>`

  it('date-formatted numeric categories render as dates', () => {
    const m = parseChartXml(AREA)!
    expect(m.categories).toEqual(['1/5/2002', '1/6/2002'])
  })

  it('area charts carry their grouping; c:dateAx parses as the category axis', () => {
    const m = parseChartXml(AREA)!
    expect(m.grouping).toBe('stacked')
  })
})

const COMBO_LABELED = `<?xml version="1.0"?><c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
<c:ser><c:idx val="0"/>
  <c:dLbls><c:showVal val="1"/></c:dLbls>
  <c:val><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>94</c:v></c:pt><c:pt idx="1"><c:v>97</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser>
<c:axId val="1"/><c:axId val="2"/></c:barChart>
<c:lineChart><c:ser><c:idx val="1"/>
  <c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="404040"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></c:spPr>
  <c:val><c:numRef><c:f>z</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>100</c:v></c:pt><c:pt idx="1"><c:v>100</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser><c:axId val="1"/><c:axId val="2"/></c:lineChart>
<c:catAx><c:axId val="1"/></c:catAx><c:valAx><c:axId val="2"/></c:valAx>
</c:plotArea></c:chart></c:chartSpace>`

describe('per-series data labels + line dash (acsa combo chart)', () => {
  it('only the series carrying c:dLbls shows labels; the index line stays label-free', () => {
    const m = parseChartXml(COMBO_LABELED)!
    const bar = m.series.find((s) => s.plotKind !== 'line') ?? m.series[0]!
    const line = m.series.find((s) => s.plotKind === 'line')!
    expect(bar.dataLabels).toBe(true)
    expect(line.dataLabels).toBe(false)
  })

  it('parses the line series prstDash and stroke width', () => {
    const m = parseChartXml(COMBO_LABELED)!
    const line = m.series.find((s) => s.plotKind === 'line')!
    expect(line.dash).toBe('sysDash')
    expect(line.lineWidthPt).toBeCloseTo(1.5, 3)
  })
})

describe('showSerName/showCatName data labels', () => {
  it('parses name flags and value-less labels', () => {
    const xml = COMBO_LABELED.replace(
      '<c:dLbls><c:showVal val="1"/></c:dLbls>',
      '<c:dLbls><c:showSerName val="1"/><c:showCatName val="1"/></c:dLbls>',
    )
    const m = parseChartXml(xml)!
    expect(m.dataLabels).toBe(true)
    expect(m.dataLabelSerName).toBe(true)
    expect(m.dataLabelCatName).toBe(true)
    expect(m.dataLabelNoValue).toBe(true)
  })
})

const MULTILVL_CHART = `<?xml version="1.0"?><c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:layout/>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
<c:ser><c:idx val="2"/><c:order val="0"/>
  <c:cat><c:multiLvlStrRef><c:f>S!$A$2:$B$5</c:f><c:multiLvlStrCache><c:ptCount val="4"/>
    <c:lvl><c:pt idx="0"><c:v>SF</c:v></c:pt><c:pt idx="1"><c:v>LA</c:v></c:pt><c:pt idx="2"><c:v>NY</c:v></c:pt><c:pt idx="3"><c:v>Albany</c:v></c:pt></c:lvl>
    <c:lvl><c:pt idx="0"><c:v>CA</c:v></c:pt><c:pt idx="2"><c:v>NY</c:v></c:pt></c:lvl>
  </c:multiLvlStrCache></c:multiLvlStrRef></c:cat>
  <c:val><c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="4"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser></c:barChart>
<c:catAx><c:tickLblPos val="none"/></c:catAx>
</c:plotArea></c:chart></c:chartSpace>`

describe('multi-level categories + palette idx + hidden tick labels', () => {
  it('parses leaf categories, outer group labels, and the series palette index', () => {
    const m = parseChartXml(MULTILVL_CHART)!
    expect(m.categories).toEqual(['SF', 'LA', 'NY', 'Albany'])
    expect(m.categoryGroups).toEqual([
      { label: 'CA', start: 0 },
      { label: 'NY', start: 2 },
    ])
    expect(m.series[0]!.paletteIdx).toBe(2)
    expect(m.catAxis?.tickLblHidden).toBe(true)
  })
})
