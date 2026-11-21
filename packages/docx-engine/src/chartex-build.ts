import { colLetter } from '@genoffice/pptx-engine/chart-workbook'
import type { ChartexLayout, NewChart } from './types'
import { escapeXmlText } from './xml-utils'

/** the relationship and content type of a chart part of the 2014 extension */
export const CHARTEX_REL_TYPE = 'http://schemas.microsoft.com/office/2014/relationships/chartEx'
export const CHARTEX_CONTENT_TYPE = 'application/vnd.ms-office.chartex+xml'
/** the namespace a w:drawing must require (mc:Choice) to carry such a chart */
export const CHARTEX_MC_NS = 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex'
export const CHARTEX_GRAPHIC_URI = 'http://schemas.microsoft.com/office/drawing/2014/chartex'

/** what each layout asks of its series beyond the data: Word's defaults for a new chart */
const LAYOUT_PR: Record<ChartexLayout, string> = {
  waterfall: '<cx:layoutPr><cx:subtotals/></cx:layoutPr>',
  funnel: '',
  treemap: '<cx:layoutPr><cx:parentLabelLayout val="overlapping"/></cx:layoutPr>',
  sunburst: '',
  boxWhisker:
    '<cx:layoutPr><cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1"/>' +
    '<cx:statistics quartileMethod="exclusive"/></cx:layoutPr>',
  histogram: '<cx:layoutPr><cx:binning intervalClosed="r"/></cx:layoutPr>',
  pareto: '<cx:layoutPr><cx:aggregation/></cx:layoutPr>',
}

/** the layoutId a layout's series carry: a histogram and a Pareto's columns are clustered columns */
const LAYOUT_ID: Record<ChartexLayout, string> = {
  waterfall: 'waterfall',
  funnel: 'funnel',
  treemap: 'treemap',
  sunburst: 'sunburst',
  boxWhisker: 'boxWhisker',
  histogram: 'clusteredColumn',
  pareto: 'clusteredColumn',
}

/** a treemap and a sunburst size their cells by the value; the rest plot it */
const sizedBy = (layout: ChartexLayout): 'size' | 'val' =>
  layout === 'treemap' || layout === 'sunburst' ? 'size' : 'val'

/** these fill the plot with their cells and have no axes */
const hasAxes = (layout: ChartexLayout): boolean => layout !== 'treemap' && layout !== 'sunburst'

/**
 * Build a chart part of the 2014 extension (word/charts/chartExN.xml) from
 * data: one cx:data per series, each pairing the shared category column with
 * that series' value column in the embedded workbook.
 */
export function buildChartexPartXml(
  chart: NewChart,
  layout: ChartexLayout,
  externalDataRId?: string,
): string {
  const rows = chart.categories.length
  const lastRow = rows + 1
  const pts = (values: (string | number | null)[]) =>
    values
      .map((v, i) => (v === null ? '' : `<cx:pt idx="${i}">${escapeXmlText(String(v))}</cx:pt>`))
      .join('')
  const data = chart.series
    .map((ser, i) => {
      const col = colLetter(i)
      return (
        `<cx:data id="${i}">` +
        `<cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$${lastRow}</cx:f>` +
        `<cx:lvl ptCount="${rows}">${pts(chart.categories)}</cx:lvl></cx:strDim>` +
        `<cx:numDim type="${sizedBy(layout)}"><cx:f>Sheet1!$${col}$2:$${col}$${lastRow}</cx:f>` +
        `<cx:lvl ptCount="${rows}" formatCode="General">${pts(ser.values.slice(0, rows))}</cx:lvl>` +
        '</cx:numDim></cx:data>'
      )
    })
    .join('')
  const series = chart.series
    .map((ser, i) => {
      const col = colLetter(i)
      const own =
        `<cx:series layoutId="${LAYOUT_ID[layout]}">` +
        `<cx:tx><cx:txData><cx:f>Sheet1!$${col}$1</cx:f><cx:v>${escapeXmlText(ser.name)}</cx:v>` +
        `</cx:txData></cx:tx><cx:dataId val="${i}"/>${LAYOUT_PR[layout]}</cx:series>`
      // a Pareto's line accumulates the columns it belongs to, on the percentage axis
      return layout === 'pareto'
        ? own + `<cx:series layoutId="paretoLine" ownerIdx="${i}"><cx:axisId val="2"/></cx:series>`
        : own
    })
    .join('')
  const axes = hasAxes(layout)
    ? '<cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis>' +
      '<cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis>' +
      (layout === 'pareto'
        ? '<cx:axis id="2"><cx:valScaling max="1" min="0"/><cx:units unit="percentage"/>' +
          '<cx:tickLabels/></cx:axis>'
        : '')
    : ''
  const title = chart.title
    ? '<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich><a:bodyPr/><a:p><a:r>' +
      `<a:t>${escapeXmlText(chart.title)}</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>`
    : ''
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    `xmlns:cx="${CHARTEX_GRAPHIC_URI}">` +
    '<cx:chartData>' +
    (externalDataRId ? `<cx:externalData r:id="${externalDataRId}" cx:autoUpdate="0"/>` : '') +
    `${data}</cx:chartData>` +
    `<cx:chart>${title}<cx:plotArea><cx:plotAreaRegion>${series}</cx:plotAreaRegion>${axes}` +
    '</cx:plotArea><cx:legend pos="b" align="ctr" overlay="0"/></cx:chart></cx:chartSpace>'
  )
}

export const CHART_STYLE_REL_TYPE =
  'http://schemas.microsoft.com/office/2011/relationships/chartStyle'
export const CHART_COLOR_STYLE_REL_TYPE =
  'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'
export const CHART_STYLE_CONTENT_TYPE = 'application/vnd.ms-office.chartstyle+xml'
export const CHART_COLOR_STYLE_CONTENT_TYPE = 'application/vnd.ms-office.chartcolorstyle+xml'

const CHART_STYLE_NS =
  'xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'

// CT_ChartStyle's entries, in schema order
const CHART_STYLE_ENTRIES =
  'axisTitle categoryAxis chartArea dataLabel dataLabelCallout dataPoint dataPoint3D dataPointLine dataPointMarker dataPointMarkerLayout dataPointWireframe dataTable downBar dropLine errorBar floor gridlineMajor gridlineMinor hiLoLine leaderLine legend plotArea plotArea3D seriesAxis seriesLine title trendline trendlineLabel upBar valueAxis wall'.split(
    ' ',
  )

const GREY_TEXT =
  '<a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>'
const HAIRLINE =
  '<cs:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>' +
  '<a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr>' +
  '</a:solidFill><a:round/></a:ln></cs:spPr>'
// a filled data point takes the series colour (phClr) with a thin light outline
const FILLED_POINT = {
  fill: true,
  shape:
    '<cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:ln w="9525"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></cs:spPr>',
}
// a line takes the series colour itself
const SERIES_LINE = {
  line: true,
  shape:
    '<cs:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:round/></a:ln></cs:spPr>',
}

interface StyleEntry {
  /** the fill / the line follows the series colour */
  fill?: boolean
  line?: boolean
  font?: string
  shape?: string
  /** text size, hundredths of a point */
  size?: number
}

// what differs from the plain entry (no line, no fill, text in the theme's text colour)
const CHART_STYLE_LOOK: Record<string, StyleEntry> = {
  axisTitle: { font: GREY_TEXT, size: 900 },
  categoryAxis: { font: GREY_TEXT, shape: HAIRLINE, size: 900 },
  valueAxis: { font: GREY_TEXT, size: 900 },
  seriesAxis: { font: GREY_TEXT, size: 900 },
  chartArea: {
    shape:
      '<cs:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill>' +
      HAIRLINE.slice('<cs:spPr>'.length),
    size: 900,
  },
  dataLabel: { font: GREY_TEXT, size: 850 },
  dataLabelCallout: { font: GREY_TEXT, size: 850 },
  dataPoint: FILLED_POINT,
  dataPoint3D: FILLED_POINT,
  dataPointMarker: FILLED_POINT,
  dataPointLine: SERIES_LINE,
  dataPointWireframe: SERIES_LINE,
  trendline: SERIES_LINE,
  gridlineMajor: { shape: HAIRLINE },
  gridlineMinor: { shape: HAIRLINE },
  legend: { font: GREY_TEXT, size: 900 },
  title: { font: GREY_TEXT, size: 1400 },
}

/**
 * The chart style part of a chart of the 2014 extension. Word refuses to open a
 * document whose chartEx part has no chart style and no colour style (probe
 * 2026-11-18: either one missing is "Word experienced an error trying to open
 * the file"); a classic chart part needs neither. The style is also what paints
 * the chart there: with plain entries the points are drawn with no fill.
 */
export function buildChartStylePartXml(): string {
  const auto = '<cs:styleClr val="auto"/>'
  const entry = (name: string) => {
    if (name === 'dataPointMarkerLayout')
      return '<cs:dataPointMarkerLayout symbol="circle" size="5"/>'
    const look = CHART_STYLE_LOOK[name] ?? {}
    return (
      `<cs:${name}>` +
      (look.line ? `<cs:lnRef idx="0">${auto}</cs:lnRef>` : '<cs:lnRef idx="0"/>') +
      (look.fill ? `<cs:fillRef idx="0">${auto}</cs:fillRef>` : '<cs:fillRef idx="0"/>') +
      '<cs:effectRef idx="0"/>' +
      `<cs:fontRef idx="minor">${look.font ?? '<a:schemeClr val="tx1"/>'}</cs:fontRef>` +
      (look.shape ?? '') +
      (look.size ? `<cs:defRPr sz="${look.size}"/>` : '') +
      `</cs:${name}>`
    )
  }
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<cs:chartStyle ${CHART_STYLE_NS} id="381">${CHART_STYLE_ENTRIES.map(entry).join('')}</cs:chartStyle>`
  )
}

/** the colour style part that goes with it: the theme's six accents in turn */
export function buildChartColorStylePartXml(): string {
  const accents = [1, 2, 3, 4, 5, 6].map((i) => `<a:schemeClr val="accent${i}"/>`).join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<cs:colorStyle ${CHART_STYLE_NS} meth="cycle" id="10">${accents}<cs:variation/></cs:colorStyle>`
  )
}
