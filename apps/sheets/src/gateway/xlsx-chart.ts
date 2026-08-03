import type { WorkbookChartEdit } from '../shared/desktop-api'

/// Surgical edits to a chart part (xl/charts/chartN.xml): title text, chart
/// type conversion within the axis-based family, and series solid colors.
/// Anything outside the supported envelope throws — fail closed, never
/// corrupt a chart part.

export class ChartEditError extends Error {}

const CONVERTIBLE_PLOTS = [
  'barChart',
  'lineChart',
  'areaChart',
  'pieChart',
  'doughnutChart',
] as const
type PlotElement = (typeof CONVERTIBLE_PLOTS)[number]

const TARGET_PLOTS: Record<NonNullable<WorkbookChartEdit['chartType']>, PlotElement> = {
  column: 'barChart',
  bar: 'barChart',
  line: 'lineChart',
  area: 'areaChart',
  pie: 'pieChart',
  doughnut: 'doughnutChart',
}

/// Axis ids synthesized when a pie gains axes; any values work as long as
/// plot and axes agree, and 9xx keeps them clear of authoring-tool ids.
const SYNTH_CAT_AXIS_ID = '900000001'
const SYNTH_VAL_AXIS_ID = '900000002'

export function applyChartEdit(chartXml: string, edit: WorkbookChartEdit): string {
  let xml = chartXml
  if (edit.chartType !== undefined) xml = convertChartType(xml, edit.chartType)
  if (edit.title !== undefined) xml = setChartTitle(xml, edit.title)
  // Full replacement first: the index-keyed edits below target the new set.
  if (edit.seriesSet !== undefined) xml = replaceSeries(xml, edit.seriesSet)
  if (edit.seriesColors !== undefined) {
    for (const [index, color] of Object.entries(edit.seriesColors)) {
      xml = setSeriesColor(xml, Number(index), color)
    }
  }
  if (edit.pointColors !== undefined) {
    for (const [index, points] of Object.entries(edit.pointColors)) {
      xml = setPointColors(xml, Number(index), points)
    }
  }
  if (edit.legend !== undefined) xml = setLegend(xml, edit.legend)
  if (edit.dataLabels !== undefined) xml = setDataLabels(xml, edit.dataLabels)
  // 'none' in the same batch wins: the labels are gone, position/format drop.
  if (
    edit.dataLabels !== 'none' &&
    (edit.dataLabelPosition !== undefined || edit.dataLabelFormat !== undefined)
  ) {
    xml = setDataLabelDetail(xml, edit.dataLabelPosition, edit.dataLabelFormat)
  }
  if (edit.axisTitles !== undefined) {
    if (edit.axisTitles.category !== undefined) {
      xml = setAxisTitle(xml, 'catAx', edit.axisTitles.category)
    }
    if (edit.axisTitles.value !== undefined) {
      xml = setAxisTitle(xml, 'valAx', edit.axisTitles.value)
    }
  }
  if (edit.grouping !== undefined) xml = setGrouping(xml, edit.grouping)
  if (edit.gridlines !== undefined) xml = setGridlines(xml, edit.gridlines)
  if (edit.valueAxis !== undefined) xml = setValueAxisBounds(xml, edit.valueAxis)
  if (edit.gapWidthPct !== undefined) xml = setGapWidth(xml, edit.gapWidthPct)
  if (edit.holeSizePct !== undefined) xml = setHoleSize(xml, edit.holeSizePct)
  if (edit.explosionPct !== undefined) xml = setSeriesExplosion(xml, edit.explosionPct)
  if (edit.pointExplosions !== undefined) xml = setPointExplosions(xml, edit.pointExplosions)
  if (edit.series !== undefined) {
    for (const entry of edit.series) {
      xml = setSeriesData(xml, entry)
    }
  }
  return xml
}

function spliceMatch(xml: string, match: RegExpExecArray, replacement: string): string {
  return xml.slice(0, match.index) + replacement + xml.slice(match.index + match[0].length)
}

const GRIDLINES_PATTERN = /<c:majorGridlines\/>|<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/

const VAL_AXIS_PATTERN = /<c:valAx>([\s\S]*?)<\/c:valAx>/

/// Major gridlines on the first value axis on/off. CT_ValAx order: axId,
/// scaling, delete, axPos, majorGridlines, minorGridlines, title, …
function setGridlines(chartXml: string, on: boolean): string {
  const axisMatch = VAL_AXIS_PATTERN.exec(chartXml)
  if (!axisMatch || axisMatch[1] === undefined) {
    throw new ChartEditError('Chart has no value axis.')
  }
  const inner = axisMatch[1]
  if (!on) {
    return spliceMatch(
      chartXml,
      axisMatch,
      `<c:valAx>${inner.replace(GRIDLINES_PATTERN, '')}</c:valAx>`,
    )
  }
  if (GRIDLINES_PATTERN.test(inner)) return chartXml
  const anchor = /<c:axPos val="[^"]*"\/>/.exec(inner) ?? /<c:delete val="[^"]*"\/>/.exec(inner)
  if (!anchor) throw new ChartEditError('Chart axis is missing its position element.')
  const at = anchor.index + anchor[0].length
  return spliceMatch(
    chartXml,
    axisMatch,
    `<c:valAx>${inner.slice(0, at)}<c:majorGridlines/>${inner.slice(at)}</c:valAx>`,
  )
}

/// Bar/line/area `c:grouping`. 'clustered' maps to 'standard' for line/area
/// (ST_Grouping has no clustered member). Stacked bars are written with
/// full overlap (c:overlap 100); clustered returns to the conventional -27.
function setGrouping(
  chartXml: string,
  grouping: NonNullable<WorkbookChartEdit['grouping']>,
): string {
  const plotMatch = /<c:(barChart|lineChart|areaChart)>([\s\S]*?)<\/c:\1>/.exec(chartXml)
  if (!plotMatch?.[1] || plotMatch[2] === undefined) {
    throw new ChartEditError('Only bar, line, and area charts support stacking.')
  }
  const plot = plotMatch[1]
  const value = grouping === 'clustered' && plot !== 'barChart' ? 'standard' : grouping
  let inner = plotMatch[2]
  if (/<c:grouping val="[^"]*"\/>/.test(inner)) {
    inner = inner.replace(/<c:grouping val="[^"]*"\/>/, () => `<c:grouping val="${value}"/>`)
  } else {
    // CT_*Chart order: (barDir,) grouping, varyColors, ser*…
    const barDir = /<c:barDir val="[^"]*"\/>/.exec(inner)
    const at = barDir ? barDir.index + barDir[0].length : 0
    inner = inner.slice(0, at) + `<c:grouping val="${value}"/>` + inner.slice(at)
  }
  if (plot === 'barChart') {
    const overlap = grouping === 'clustered' ? '-27' : '100'
    if (/<c:overlap val="[^"]*"\/>/.test(inner)) {
      inner = inner.replace(/<c:overlap val="[^"]*"\/>/, () => `<c:overlap val="${overlap}"/>`)
    } else {
      // CT_BarChart puts overlap right before the axId pair.
      const axId = /<c:axId val="[0-9]+"\/>/.exec(inner)
      if (!axId) throw new ChartEditError('Chart has no axes for a stacking edit.')
      inner = inner.slice(0, axId.index) + `<c:overlap val="${overlap}"/>` + inner.slice(axId.index)
    }
  }
  return chartXml.replace(plotMatch[0], () => `<c:${plot}>${inner}</c:${plot}>`)
}

/// Value-axis min/max on the first value axis; null resets a bound to auto.
function setValueAxisBounds(
  chartXml: string,
  bounds: NonNullable<WorkbookChartEdit['valueAxis']>,
): string {
  const axisMatch = VAL_AXIS_PATTERN.exec(chartXml)
  if (!axisMatch || axisMatch[1] === undefined) {
    throw new ChartEditError('Chart has no value axis.')
  }
  const scalingMatch = /<c:scaling>([\s\S]*?)<\/c:scaling>/.exec(axisMatch[1])
  if (!scalingMatch || scalingMatch[1] === undefined) {
    throw new ChartEditError('Chart value axis has no scaling element.')
  }
  let scaling = scalingMatch[1]
  if (bounds.max !== undefined) scaling = patchScalingBound(scaling, 'max', bounds.max)
  if (bounds.min !== undefined) scaling = patchScalingBound(scaling, 'min', bounds.min)
  const inner = spliceMatch(axisMatch[1], scalingMatch, `<c:scaling>${scaling}</c:scaling>`)
  return spliceMatch(chartXml, axisMatch, `<c:valAx>${inner}</c:valAx>`)
}

/// CT_Scaling order: logBase, orientation, max, min.
function patchScalingBound(
  scalingInner: string,
  bound: 'max' | 'min',
  value: number | null,
): string {
  const pattern = new RegExp(`<c:${bound} val="[^"]*"/>`)
  if (value === null) return scalingInner.replace(pattern, '')
  const element = `<c:${bound} val="${value}"/>`
  if (pattern.test(scalingInner)) return scalingInner.replace(pattern, () => element)
  const anchors =
    bound === 'max'
      ? /<c:logBase val="[^"]*"\/>|<c:orientation val="[^"]*"\/>/g
      : /<c:logBase val="[^"]*"\/>|<c:orientation val="[^"]*"\/>|<c:max val="[^"]*"\/>/g
  let insertAt = 0
  for (const found of scalingInner.matchAll(anchors)) {
    insertAt = found.index + found[0].length
  }
  return scalingInner.slice(0, insertAt) + element + scalingInner.slice(insertAt)
}

/// CT_BarChart order: …ser*, dLbls, gapWidth, overlap, axId.
function setGapWidth(chartXml: string, pct: number): string {
  const plotPattern = /<c:barChart>([\s\S]*?)<\/c:barChart>/
  const plotMatch = plotPattern.exec(chartXml)
  if (!plotMatch || plotMatch[1] === undefined) {
    throw new ChartEditError('Only bar charts have a gap width.')
  }
  const inner = plotMatch[1]
  const element = `<c:gapWidth val="${pct}"/>`
  const existing = /<c:gapWidth val="[^"]*"\/>/
  if (existing.test(inner)) {
    return spliceMatch(
      chartXml,
      plotMatch,
      `<c:barChart>${inner.replace(existing, () => element)}</c:barChart>`,
    )
  }
  let insertAt = -1
  for (const found of inner.matchAll(/<\/c:ser>|<\/c:dLbls>/g)) {
    insertAt = found.index + found[0].length
  }
  if (insertAt === -1) throw new ChartEditError('This chart has no series.')
  return spliceMatch(
    chartXml,
    plotMatch,
    `<c:barChart>${inner.slice(0, insertAt)}${element}${inner.slice(insertAt)}</c:barChart>`,
  )
}

/// holeSize is the last CT_DoughnutChart member (before extLst).
function setHoleSize(chartXml: string, pct: number): string {
  const plotPattern = /<c:doughnutChart>([\s\S]*?)<\/c:doughnutChart>/
  const plotMatch = plotPattern.exec(chartXml)
  if (!plotMatch || plotMatch[1] === undefined) {
    throw new ChartEditError('Only doughnut charts have a hole size.')
  }
  const inner = plotMatch[1]
  const element = `<c:holeSize val="${pct}"/>`
  const existing = /<c:holeSize val="[^"]*"\/>/
  let next: string
  if (existing.test(inner)) {
    next = inner.replace(existing, () => element)
  } else {
    const extLst = /<c:extLst[\s/>]/.exec(inner)
    const at = extLst ? extLst.index : inner.length
    next = inner.slice(0, at) + element + inner.slice(at)
  }
  return spliceMatch(chartXml, plotMatch, `<c:doughnutChart>${next}</c:doughnutChart>`)
}

/// Patches the first series of the first pie/doughnut plot.
function withFirstPieSeries(
  chartXml: string,
  what: string,
  patch: (serXml: string) => string,
): string {
  const plotPattern = /<c:(pieChart|doughnutChart)>([\s\S]*?)<\/c:\1>/
  const plotMatch = plotPattern.exec(chartXml)
  if (!plotMatch?.[1] || plotMatch[2] === undefined) {
    throw new ChartEditError(`Only pie and doughnut charts support ${what}.`)
  }
  const serMatch = /<c:ser>[\s\S]*?<\/c:ser>/.exec(plotMatch[2])
  if (!serMatch) throw new ChartEditError('This chart has no series.')
  const inner = spliceMatch(plotMatch[2], serMatch, patch(serMatch[0]))
  return spliceMatch(chartXml, plotMatch, `<c:${plotMatch[1]}>${inner}</c:${plotMatch[1]}>`)
}

/// Series-level explosion (whole ring). CT_PieSer order: …spPr, explosion,
/// dPt* — a dPt's own explosion must stay untouched, so only the slice
/// before the first dPt is considered.
function setSeriesExplosion(chartXml: string, pct: number): string {
  return withFirstPieSeries(chartXml, 'explosion', (serXml) => {
    const element = `<c:explosion val="${pct}"/>`
    const dPtAt = serXml.search(/<c:dPt[>/]/)
    const scope = dPtAt === -1 ? serXml : serXml.slice(0, dPtAt)
    const tail = dPtAt === -1 ? '' : serXml.slice(dPtAt)
    const pattern = /<c:explosion val="[^"]*"\/>/
    if (pattern.test(scope)) return scope.replace(pattern, () => element) + tail
    if (dPtAt !== -1) return scope + element + tail
    const anchor = /<c:dLbls[>/]|<c:cat[>/]|<c:val[>/]/.exec(serXml)
    const at = anchor ? anchor.index : serXml.lastIndexOf('</c:ser>')
    return serXml.slice(0, at) + element + serXml.slice(at)
  })
}

/// Per-slice explosion overrides on the first pie/doughnut series.
function setPointExplosions(chartXml: string, points: Record<string, number>): string {
  return withFirstPieSeries(chartXml, 'per-slice explosion', (serXml) => {
    let next = serXml
    const entries = Object.entries(points)
      .map(([point, pct]) => [Number(point), pct] as const)
      .sort((a, b) => a[0] - b[0])
    for (const [point, pct] of entries) {
      next = upsertPointExplosion(next, point, pct)
    }
    return next
  })
}

function upsertPointExplosion(serXml: string, pointIndex: number, pct: number): string {
  const element = `<c:explosion val="${pct}"/>`
  const dPts = [...serXml.matchAll(/<c:dPt>[\s\S]*?<\/c:dPt>/g)]
  for (const dPt of dPts) {
    const idx = Number(/<c:idx val="([0-9]+)"\/>/.exec(dPt[0])?.[1])
    if (idx !== pointIndex) continue
    const pattern = /<c:explosion val="[^"]*"\/>/
    let patched: string
    if (pattern.test(dPt[0])) {
      patched = dPt[0].replace(pattern, () => element)
    } else {
      // CT_DPt order: idx, invertIfNegative, marker, bubble3D, explosion, spPr.
      const anchor = /<c:spPr[\s/>]|<c:pictureOptions[\s/>]|<c:extLst[\s/>]/.exec(dPt[0])
      const at = anchor ? anchor.index : dPt[0].lastIndexOf('</c:dPt>')
      patched = dPt[0].slice(0, at) + element + dPt[0].slice(at)
    }
    return serXml.slice(0, dPt.index) + patched + serXml.slice(dPt.index + dPt[0].length)
  }
  const created = `<c:dPt><c:idx val="${pointIndex}"/><c:bubble3D val="0"/>${element}</c:dPt>`
  const insertAt = dPtInsertionIndex(serXml, dPts, pointIndex)
  return serXml.slice(0, insertAt) + created + serXml.slice(insertAt)
}

const REPLACEABLE_PLOTS: readonly string[] = [
  'barChart',
  'lineChart',
  'areaChart',
  'pieChart',
  'doughnutChart',
  'radarChart',
  'scatterChart',
]

/// Full series replacement (Select Data) on a single-plot chart: every
/// existing c:ser (with its dPt/trendline baggage) drops, the new set is
/// written in order at the old series slot, other plot children stay.
function replaceSeries(
  chartXml: string,
  entries: NonNullable<WorkbookChartEdit['seriesSet']>,
): string {
  const plotPattern =
    /<c:(barChart|lineChart|areaChart|bar3DChart|line3DChart|area3DChart|pieChart|pie3DChart|doughnutChart|scatterChart|radarChart)>([\s\S]*?)<\/c:\1>/g
  const plots = [...chartXml.matchAll(plotPattern)]
  const plotMatch = plots[0]
  if (!plotMatch?.[1] || plotMatch[2] === undefined) {
    throw new ChartEditError('This chart type does not support series replacement.')
  }
  if (plots.length > 1) {
    throw new ChartEditError('Combo charts cannot have their series replaced.')
  }
  const plot = plotMatch[1]
  if (!REPLACEABLE_PLOTS.includes(plot)) {
    throw new ChartEditError(`Replacing series on a ${plot} is not supported.`)
  }
  const inner = plotMatch[2]
  const firstSer = inner.search(/<c:ser>/)
  if (firstSer === -1) throw new ChartEditError('This chart has no series.')
  const built = entries
    .map((entry, index) =>
      plot === 'scatterChart'
        ? buildScatterSeriesReplacement(entry, index)
        : buildSeriesReplacement(entry, index),
    )
    .join('')
  const nextInner =
    inner.slice(0, firstSer) +
    built +
    inner.slice(firstSer).replace(/<c:ser>[\s\S]*?<\/c:ser>/g, '')
  return (
    chartXml.slice(0, plotMatch.index) +
    `<c:${plot}>${nextInner}</c:${plot}>` +
    chartXml.slice(plotMatch.index + plotMatch[0].length)
  )
}

/// Same cache structure buildSeriesXml (xlsx-drawing-add) produces: refs get
/// strRef/numRef with caches, ref-less data becomes strLit/numLit.
function buildSeriesReplacement(
  entry: NonNullable<WorkbookChartEdit['seriesSet']>[number],
  index: number,
): string {
  const spPr =
    entry.color === undefined
      ? ''
      : `<c:spPr><a:solidFill><a:srgbClr val="${entry.color.slice(1).toUpperCase()}"/></a:solidFill></c:spPr>`
  const categories = entry.categories?.slice(0, entry.values.length)
  const catPoints = (categories ?? [])
    .map((value, idx) => `<c:pt idx="${idx}"><c:v>${escapeXmlText(value)}</c:v></c:pt>`)
    .join('')
  const cat =
    categories === undefined || categories.length === 0
      ? ''
      : '<c:cat>' +
        (entry.categoriesRef !== undefined
          ? `<c:strRef><c:f>${escapeXmlText(entry.categoriesRef)}</c:f>` +
            `<c:strCache><c:ptCount val="${categories.length}"/>${catPoints}</c:strCache></c:strRef>`
          : `<c:strLit><c:ptCount val="${categories.length}"/>${catPoints}</c:strLit>`) +
        '</c:cat>'
  const val = `<c:val>${numDataXml(entry.valuesRef, entry.values)}</c:val>`
  return (
    `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>` +
    `<c:tx><c:v>${escapeXmlText(entry.name)}</c:v></c:tx>${spPr}${cat}${val}</c:ser>`
  )
}

/// numRef with cache when a ref is given, numLit otherwise.
function numDataXml(ref: string | undefined, values: readonly number[]): string {
  const points =
    `<c:ptCount val="${values.length}"/>` +
    values.map((value, idx) => `<c:pt idx="${idx}"><c:v>${value}</c:v></c:pt>`).join('')
  return ref !== undefined
    ? `<c:numRef><c:f>${escapeXmlText(ref)}</c:f>` +
        `<c:numCache><c:formatCode>General</c:formatCode>${points}</c:numCache></c:numRef>`
    : `<c:numLit>${points}</c:numLit>`
}

/// Scatter rebuild, mirroring buildScatterSeriesXml (xlsx-drawing-add):
/// all-numeric categories become the X values, otherwise ordinals 0..n-1.
/// CT_ScatterSer: idx, order, tx, spPr, …, xVal, yVal, smooth.
function buildScatterSeriesReplacement(
  entry: NonNullable<WorkbookChartEdit['seriesSet']>[number],
  index: number,
): string {
  const spPr =
    entry.color === undefined
      ? ''
      : '<c:spPr><a:ln><a:solidFill>' +
        `<a:srgbClr val="${entry.color.slice(1).toUpperCase()}"/></a:solidFill></a:ln></c:spPr>`
  const categories = entry.categories?.slice(0, entry.values.length) ?? []
  const numericCategories = categories.map((value) => Number(value))
  const categoriesAreNumeric =
    categories.length === entry.values.length &&
    categories.length > 0 &&
    numericCategories.every((value) => Number.isFinite(value))
  const xValues = categoriesAreNumeric ? numericCategories : entry.values.map((_, idx) => idx)
  const xVal = `<c:xVal>${numDataXml(categoriesAreNumeric ? entry.categoriesRef : undefined, xValues)}</c:xVal>`
  const yVal = `<c:yVal>${numDataXml(entry.valuesRef, entry.values)}</c:yVal>`
  return (
    `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>` +
    `<c:tx><c:v>${escapeXmlText(entry.name)}</c:v></c:tx>${spPr}${xVal}${yVal}` +
    '<c:smooth val="0"/></c:ser>'
  )
}

const LEGEND_POSITIONS: Record<string, string> = {
  right: 'r',
  bottom: 'b',
  top: 't',
  left: 'l',
}

/// Removes the legend or moves/creates it at the given side.
function setLegend(chartXml: string, legend: NonNullable<WorkbookChartEdit['legend']>): string {
  const withoutLegend = chartXml.replace(/<c:legend>[\s\S]*?<\/c:legend>/, '')
  if (legend === 'none') return withoutLegend
  const position = LEGEND_POSITIONS[legend]
  const element = `<c:legend><c:legendPos val="${position}"/><c:overlay val="0"/></c:legend>`
  const plotAreaEnd = /<\/c:plotArea>/.exec(withoutLegend)
  if (!plotAreaEnd) throw new ChartEditError('Chart part has no plot area.')
  const insertAt = plotAreaEnd.index + plotAreaEnd[0].length
  return withoutLegend.slice(0, insertAt) + element + withoutLegend.slice(insertAt)
}

/// Plot-level data labels: what every slice/bar/point shows. Excel's Quick
/// Layout presets ride on this. 'none' removes the plot-level element
/// (per-series overrides, if any, keep their own settings).
function setDataLabels(
  chartXml: string,
  mode: NonNullable<WorkbookChartEdit['dataLabels']>,
): string {
  const plotPattern =
    /<c:(barChart|lineChart|areaChart|pieChart|doughnutChart)>([\s\S]*?)(<\/c:\1>)/
  const plotMatch = plotPattern.exec(chartXml)
  if (!plotMatch?.[1] || plotMatch[2] === undefined) {
    throw new ChartEditError('This chart type does not support data labels.')
  }
  const inner = plotMatch[2]
  // The plot-level dLbls slot sits right after the last series.
  const withoutLabels = inner.replace(/(<\/c:ser>)<c:dLbls>[\s\S]*?<\/c:dLbls>/, '$1')
  let next = withoutLabels
  if (mode !== 'none') {
    const showValue = mode === 'value'
    const showCategory = mode === 'category-percent'
    const showPercent = mode === 'percent' || mode === 'category-percent'
    const labels =
      '<c:dLbls><c:showLegendKey val="0"/>' +
      `<c:showVal val="${showValue ? '1' : '0'}"/>` +
      `<c:showCatName val="${showCategory ? '1' : '0'}"/>` +
      '<c:showSerName val="0"/>' +
      `<c:showPercent val="${showPercent ? '1' : '0'}"/>` +
      '<c:showBubbleSize val="0"/></c:dLbls>'
    const lastSeriesEnd = withoutLabels.lastIndexOf('</c:ser>')
    if (lastSeriesEnd === -1) throw new ChartEditError('This chart has no series to label.')
    const insertAt = lastSeriesEnd + '</c:ser>'.length
    next = withoutLabels.slice(0, insertAt) + labels + withoutLabels.slice(insertAt)
  }
  return chartXml.replace(
    plotPattern,
    (_full, plot: string, _inner, close: string) => `<c:${plot}>${next}${close}`,
  )
}

const DLBL_POSITIONS: Record<NonNullable<WorkbookChartEdit['dataLabelPosition']>, string> = {
  center: 'ctr',
  'inside-end': 'inEnd',
  'outside-end': 'outEnd',
}

/// Label position (c:dLblPos) and number format (c:numFmt) on the plot-level
/// dLbls, upgrading an existing element or creating a value-labels one.
/// CT_DLbls order: numFmt, spPr, txPr, dLblPos, show*.
function setDataLabelDetail(
  chartXml: string,
  position: WorkbookChartEdit['dataLabelPosition'],
  format: WorkbookChartEdit['dataLabelFormat'],
): string {
  const plotPattern =
    /<c:(barChart|lineChart|areaChart|pieChart|doughnutChart)>([\s\S]*?)(<\/c:\1>)/
  const plotMatch = plotPattern.exec(chartXml)
  if (!plotMatch?.[1] || plotMatch[2] === undefined) {
    throw new ChartEditError('This chart type does not support data labels.')
  }
  const inner = plotMatch[2]
  const numFmt =
    format === undefined
      ? ''
      : `<c:numFmt formatCode="${escapeXmlAttribute(format)}" sourceLinked="0"/>`
  const dLblPos = position === undefined ? '' : `<c:dLblPos val="${DLBL_POSITIONS[position]}"/>`
  const existing = /<\/c:ser><c:dLbls>([\s\S]*?)<\/c:dLbls>/.exec(inner)
  let next: string
  if (existing && existing[1] !== undefined) {
    let labels = existing[1]
    if (numFmt !== '') {
      const pattern = /<c:numFmt [^>]*\/>/
      labels = pattern.test(labels) ? labels.replace(pattern, () => numFmt) : numFmt + labels
    }
    if (dLblPos !== '') {
      const pattern = /<c:dLblPos val="[^"]*"\/>/
      if (pattern.test(labels)) {
        labels = labels.replace(pattern, () => dLblPos)
      } else {
        const anchor = /<c:show|<c:separator[>/]/.exec(labels)
        const at = anchor ? anchor.index : labels.length
        labels = labels.slice(0, at) + dLblPos + labels.slice(at)
      }
    }
    next = spliceMatch(inner, existing, `</c:ser><c:dLbls>${labels}</c:dLbls>`)
  } else {
    const lastSeriesEnd = inner.lastIndexOf('</c:ser>')
    if (lastSeriesEnd === -1) throw new ChartEditError('This chart has no series to label.')
    const at = lastSeriesEnd + '</c:ser>'.length
    const created =
      `<c:dLbls>${numFmt}${dLblPos}<c:showLegendKey val="0"/><c:showVal val="1"/>` +
      '<c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>' +
      '<c:showBubbleSize val="0"/></c:dLbls>'
    next = inner.slice(0, at) + created + inner.slice(at)
  }
  return spliceMatch(chartXml, plotMatch, `<c:${plotMatch[1]}>${next}${plotMatch[3]}`)
}

/// Sets or removes (null) the title of the first category/value axis.
function setAxisTitle(chartXml: string, axis: 'catAx' | 'valAx', title: string | null): string {
  const axisPattern = new RegExp(`<c:${axis}>([\\s\\S]*?)</c:${axis}>`)
  const axisMatch = axisPattern.exec(chartXml)
  if (!axisMatch || axisMatch[1] === undefined) {
    // Date axes stand in for category axes on time-series charts.
    if (axis === 'catAx' && /<c:dateAx>/.test(chartXml)) {
      return setDateAxisTitle(chartXml, title)
    }
    throw new ChartEditError(`Chart has no ${axis === 'catAx' ? 'category' : 'value'} axis.`)
  }
  return chartXml.replace(
    axisPattern,
    (_full, inner: string) => `<c:${axis}>${patchAxisInnerTitle(inner, title)}</c:${axis}>`,
  )
}

function setDateAxisTitle(chartXml: string, title: string | null): string {
  return chartXml.replace(
    /<c:dateAx>([\s\S]*?)<\/c:dateAx>/,
    (_full, inner: string) => `<c:dateAx>${patchAxisInnerTitle(inner, title)}</c:dateAx>`,
  )
}

/// CT_*Ax order: the title slot follows axPos / gridlines.
function patchAxisInnerTitle(axisInner: string, title: string | null): string {
  const stripped = axisInner.replace(/<c:title>[\s\S]*?<\/c:title>/, '')
  if (title === null) return stripped
  const anchor =
    /<c:minorGridlines\/>|<c:minorGridlines>[\s\S]*?<\/c:minorGridlines>|<c:majorGridlines\/>|<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>|<c:axPos val="[^"]*"\/>/g
  let insertAt = -1
  for (const found of stripped.matchAll(anchor)) {
    insertAt = found.index + found[0].length
  }
  if (insertAt === -1) throw new ChartEditError('Chart axis is missing its position element.')
  return stripped.slice(0, insertAt) + builtTitle(escapeXmlText(title)) + stripped.slice(insertAt)
}

/// Rewrites a series' name and/or data (reference + cached values together).
/// On a scatter plot the categories become numeric X values (c:xVal) and the
/// values land in c:yVal.
function setSeriesData(
  chartXml: string,
  entry: NonNullable<WorkbookChartEdit['series']>[number],
): string {
  const isScatter = /<c:scatterChart[\s>]/.test(chartXml)
  const seriesPattern = /<c:ser>[\s\S]*?<\/c:ser>/g
  let found = false
  let position = -1
  const result = chartXml.replace(seriesPattern, (serXml) => {
    position += 1
    if (position !== entry.index) return serXml
    found = true
    let next = serXml
    if (entry.name !== undefined) {
      const tx = `<c:tx><c:v>${escapeXmlText(entry.name)}</c:v></c:tx>`
      if (/<c:tx>[\s\S]*?<\/c:tx>/.test(next)) {
        next = next.replace(/<c:tx>[\s\S]*?<\/c:tx>/, () => tx)
      } else {
        const anchor = /<c:order val="[0-9]+"\/>/.exec(next)
        if (!anchor) throw new ChartEditError('Chart series is missing idx/order structure.')
        next =
          next.slice(0, anchor.index + anchor[0].length) +
          tx +
          next.slice(anchor.index + anchor[0].length)
      }
    }
    if (entry.categories !== undefined) {
      if (isScatter) {
        next = setScatterXValues(next, entry.categories, entry.categoriesRef)
      } else {
        const catPattern = /<c:cat\/>|<c:cat>[\s\S]*?<\/c:cat>/
        const previousCat = catPattern.exec(next)?.[0] ?? ''
        // Numeric categories cached as numRef/numLit (dates) keep that
        // shape and its formatCode; a strRef rewrite drops the format for good.
        const numeric = entry.categories.map((value) => Number(value))
        const keepNumeric =
          /<c:num(?:Ref|Lit)[\s>]/.test(previousCat) &&
          entry.categories.length > 0 &&
          entry.categories.every(
            (value, idx) => value.trim() !== '' && Number.isFinite(numeric[idx]),
          )
        let cat: string
        if (keepNumeric) {
          const formatCode =
            /<c:formatCode>([^<]*)<\/c:formatCode>/.exec(previousCat)?.[1] ?? 'General'
          const data =
            `<c:formatCode>${formatCode}</c:formatCode>` +
            `<c:ptCount val="${entry.categories.length}"/>` +
            numeric.map((value, idx) => `<c:pt idx="${idx}"><c:v>${value}</c:v></c:pt>`).join('')
          // c:f is mandatory inside numRef, so ref-less data must be a literal.
          cat =
            entry.categoriesRef === undefined
              ? `<c:cat><c:numLit>${data}</c:numLit></c:cat>`
              : `<c:cat><c:numRef><c:f>${escapeXmlText(entry.categoriesRef)}</c:f>` +
                `<c:numCache>${data}</c:numCache></c:numRef></c:cat>`
        } else {
          const points =
            `<c:ptCount val="${entry.categories.length}"/>` +
            entry.categories
              .map((value, idx) => `<c:pt idx="${idx}"><c:v>${escapeXmlText(value)}</c:v></c:pt>`)
              .join('')
          // c:f is mandatory inside strRef, so ref-less data must be a literal.
          cat =
            entry.categoriesRef === undefined
              ? `<c:cat><c:strLit>${points}</c:strLit></c:cat>`
              : `<c:cat><c:strRef><c:f>${escapeXmlText(entry.categoriesRef)}</c:f>` +
                `<c:strCache>${points}</c:strCache></c:strRef></c:cat>`
        }
        if (previousCat !== '') {
          next = next.replace(catPattern, () => cat)
        } else {
          const valElement = /<c:val\/>|<c:val>/.exec(next)
          if (!valElement) throw new ChartEditError('Chart series has no value element.')
          next = next.slice(0, valElement.index) + cat + next.slice(valElement.index)
        }
      }
    }
    if (entry.values !== undefined) {
      const tag = isScatter ? 'yVal' : 'val'
      const points = entry.values
        .map((value, idx) => `<c:pt idx="${idx}"><c:v>${value}</c:v></c:pt>`)
        .join('')
      // Keep the original number format: the axis-scale renderer reads it.
      const valPattern = new RegExp(`<c:${tag}/>|<c:${tag}>[\\s\\S]*?</c:${tag}>`)
      const previousVal = valPattern.exec(next)?.[0] ?? ''
      const formatCode = /<c:formatCode>([^<]*)<\/c:formatCode>/.exec(previousVal)?.[1] ?? 'General'
      const data =
        `<c:formatCode>${formatCode}</c:formatCode>` +
        `<c:ptCount val="${entry.values.length}"/>${points}`
      // c:f is mandatory inside numRef, so ref-less data must be a literal.
      const val =
        entry.valuesRef === undefined
          ? `<c:${tag}><c:numLit>${data}</c:numLit></c:${tag}>`
          : `<c:${tag}><c:numRef><c:f>${escapeXmlText(entry.valuesRef)}</c:f>` +
            `<c:numCache>${data}</c:numCache></c:numRef></c:${tag}>`
      if (!valPattern.test(next)) throw new ChartEditError('Chart series has no value element.')
      next = next.replace(valPattern, () => val)
    }
    return next
  })
  if (!found) throw new ChartEditError(`Chart has no series with index ${entry.index}.`)
  return result
}

/// Scatter X data (c:xVal): categories numericized, a non-numeric entry
/// falls back to its own index.
function setScatterXValues(
  serXml: string,
  categories: readonly string[],
  categoriesRef: string | undefined,
): string {
  const points =
    `<c:formatCode>General</c:formatCode><c:ptCount val="${categories.length}"/>` +
    categories
      .map((value, idx) => {
        const numeric = Number(value)
        const x = value.trim() !== '' && Number.isFinite(numeric) ? numeric : idx
        return `<c:pt idx="${idx}"><c:v>${x}</c:v></c:pt>`
      })
      .join('')
  const xVal =
    categoriesRef === undefined
      ? `<c:xVal><c:numLit>${points}</c:numLit></c:xVal>`
      : `<c:xVal><c:numRef><c:f>${escapeXmlText(categoriesRef)}</c:f>` +
        `<c:numCache>${points}</c:numCache></c:numRef></c:xVal>`
  const xValPattern = /<c:xVal\/>|<c:xVal>[\s\S]*?<\/c:xVal>/
  if (xValPattern.test(serXml)) return serXml.replace(xValPattern, () => xVal)
  // CT_ScatterSer: xVal precedes yVal and smooth.
  const anchor = /<c:yVal[>/]|<c:smooth[\s/>]/.exec(serXml)
  const at = anchor ? anchor.index : serXml.lastIndexOf('</c:ser>')
  return serXml.slice(0, at) + xVal + serXml.slice(at)
}

/// Replaces the first title run's text (preserving its styling), blanks any
/// further runs, or inserts a fresh title when the chart has none.
/// Only the chart-level title qualifies — axes carry their own c:title
/// inside the plot area, so the search stops there.
function setChartTitle(chartXml: string, title: string): string {
  const escaped = escapeXmlText(title)
  const xml = chartXml.replace(/<c:autoTitleDeleted val="1"\/>/, '<c:autoTitleDeleted val="0"/>')
  const plotAreaAt = xml.search(/<c:plotArea[\s>]/)
  const scope = plotAreaAt === -1 ? xml : xml.slice(0, plotAreaAt)
  const titleElement = /<c:title>[\s\S]*?<\/c:title>/.exec(scope)
  if (titleElement) {
    let isFirstRun = true
    const rewritten = titleElement[0].replace(
      /(<a:t>)[\s\S]*?(<\/a:t>)|<a:t\/>/g,
      (_full, open: string | undefined, close: string | undefined) => {
        const text = isFirstRun ? escaped : ''
        isFirstRun = false
        return `${open ?? '<a:t>'}${text}${close ?? '</a:t>'}`
      },
    )
    if (isFirstRun) {
      // A title without any text run (e.g. formula-linked): replace whole tx.
      return xml.replace(/<c:title>[\s\S]*?<\/c:title>/, () => builtTitle(escaped))
    }
    return xml.replace(/<c:title>[\s\S]*?<\/c:title>/, () => rewritten)
  }
  const chartOpen = /<c:chart>/.exec(xml)
  if (!chartOpen) throw new ChartEditError('Chart part has no c:chart element.')
  const insertAt = chartOpen.index + chartOpen[0].length
  const autoTitle = /<c:autoTitleDeleted /.test(xml) ? '' : '<c:autoTitleDeleted val="0"/>'
  return xml.slice(0, insertAt) + `${builtTitle(escaped)}${autoTitle}` + xml.slice(insertAt)
}

function builtTitle(escapedText: string): string {
  return (
    '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>' +
    `<a:t>${escapedText}</a:t>` +
    '</a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
  )
}

/// Converts between column/bar/line/area/pie/doughnut. Bar direction changes
/// rewrite only c:barDir; cross-type conversion rebuilds the plot element
/// keeping series and dropping type-specific leftovers. Axis↔pie conversions
/// also remove or synthesize the axis pair, since pie plots reference none.
function convertChartType(
  chartXml: string,
  target: NonNullable<WorkbookChartEdit['chartType']>,
): string {
  const plotPattern =
    /<c:(barChart|lineChart|areaChart|bar3DChart|line3DChart|area3DChart|pieChart|pie3DChart|doughnutChart|scatterChart|radarChart)>([\s\S]*?)<\/c:\1>/g
  const plots = [...chartXml.matchAll(plotPattern)]
  const plotMatch = plots[0]
  if (!plotMatch?.[1] || plotMatch[2] === undefined) {
    throw new ChartEditError('This chart type cannot be converted.')
  }
  if (plots.length > 1) {
    throw new ChartEditError('Combo charts cannot be converted.')
  }
  const current = plotMatch[1]
  if (!(CONVERTIBLE_PLOTS as readonly string[]).includes(current)) {
    throw new ChartEditError(`Converting a ${current} is not supported.`)
  }
  const inner = plotMatch[2]
  const targetElement = TARGET_PLOTS[target]
  const currentIsPie = current === 'pieChart' || current === 'doughnutChart'
  const targetIsPie = targetElement === 'pieChart' || targetElement === 'doughnutChart'

  if (current === 'barChart' && targetElement === 'barChart') {
    const direction = target === 'column' ? 'col' : 'bar'
    return chartXml.replace(/<c:barDir val="[^"]*"\/>/, () => `<c:barDir val="${direction}"/>`)
  }
  if (targetElement === current) return chartXml

  const series = [...inner.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((match) =>
    sanitizeSeries(match[0], targetIsPie),
  )
  if (series.length === 0) {
    throw new ChartEditError('This chart has no series to convert.')
  }
  const parts: string[] = []
  if (targetElement === 'barChart') {
    parts.push(`<c:barDir val="${target === 'bar' ? 'bar' : 'col'}"/>`)
  }
  if (!targetIsPie) {
    const grouping = /<c:grouping val="(stacked|percentStacked)"\/>/.exec(inner)?.[1]
    parts.push(
      `<c:grouping val="${grouping ?? (targetElement === 'barChart' ? 'clustered' : 'standard')}"/>`,
    )
  }
  parts.push(`<c:varyColors val="${targetIsPie ? '1' : '0'}"/>`, ...series)
  // The plot-level data labels survive the conversion; only the position is
  // dropped (its allowed values differ per chart type).
  const plotLabels = /<\/c:ser>(<c:dLbls>[\s\S]*?<\/c:dLbls>)/.exec(inner)?.[1]
  if (plotLabels) parts.push(plotLabels.replace(/<c:dLblPos val="[^"]*"\/>/, ''))
  if (targetElement === 'barChart') parts.push('<c:gapWidth val="150"/>')
  if (targetElement === 'lineChart') parts.push('<c:marker val="1"/>')
  if (targetIsPie) {
    parts.push('<c:firstSliceAng val="0"/>')
    if (targetElement === 'doughnutChart') parts.push('<c:holeSize val="50"/>')
    const plotReplaced = chartXml.replace(
      plotMatch[0],
      () => `<c:${targetElement}>${parts.join('')}</c:${targetElement}>`,
    )
    return plotReplaced.replace(
      /<c:(catAx|valAx|dateAx|serAx)>[\s\S]*?<\/c:\1>|<c:(?:catAx|valAx|dateAx|serAx)\/>/g,
      '',
    )
  }
  if (!currentIsPie) {
    const axisIds = [...inner.matchAll(/<c:axId val="[0-9]+"\/>/g)].map((match) => match[0])
    if (axisIds.length === 0) {
      throw new ChartEditError('This chart has no axes to convert.')
    }
    parts.push(...axisIds)
    return chartXml.replace(
      plotMatch[0],
      () => `<c:${targetElement}>${parts.join('')}</c:${targetElement}>`,
    )
  }
  // Pie → axis chart: the plot gains a synthesized category/value axis pair
  // (bar charts flip them onto the other sides).
  parts.push(`<c:axId val="${SYNTH_CAT_AXIS_ID}"/>`, `<c:axId val="${SYNTH_VAL_AXIS_ID}"/>`)
  const catPos = target === 'bar' ? 'l' : 'b'
  const valPos = target === 'bar' ? 'b' : 'l'
  const axes =
    `<c:catAx><c:axId val="${SYNTH_CAT_AXIS_ID}"/>` +
    '<c:scaling><c:orientation val="minMax"/></c:scaling>' +
    `<c:delete val="0"/><c:axPos val="${catPos}"/>` +
    `<c:crossAx val="${SYNTH_VAL_AXIS_ID}"/></c:catAx>` +
    `<c:valAx><c:axId val="${SYNTH_VAL_AXIS_ID}"/>` +
    '<c:scaling><c:orientation val="minMax"/></c:scaling>' +
    `<c:delete val="0"/><c:axPos val="${valPos}"/><c:majorGridlines/>` +
    `<c:crossAx val="${SYNTH_CAT_AXIS_ID}"/></c:valAx>`
  return chartXml.replace(
    plotMatch[0],
    () => `<c:${targetElement}>${parts.join('')}</c:${targetElement}>${axes}`,
  )
}

/// Strips per-type series options that are invalid on other chart types.
function sanitizeSeries(serXml: string, forPie: boolean): string {
  const cleaned = serXml
    .replace(/<c:invertIfNegative val="[^"]*"\/>/g, '')
    .replace(/<c:marker>[\s\S]*?<\/c:marker>/g, '')
    .replace(/<c:smooth val="[^"]*"\/>/g, '')
    .replace(/<c:shape val="[^"]*"\/>/g, '')
    .replace(/<c:explosion val="[^"]*"\/>/g, '')
  // CT_PieSer has no trendline slot.
  return forPie ? cleaned.replace(/<c:trendline>[\s\S]*?<\/c:trendline>/g, '') : cleaned
}

/// Sets a series' solid color (fill + line, so bars and lines both take it).
function setSeriesColor(chartXml: string, seriesIndex: number, color: string): string {
  const argb = color.slice(1).toUpperCase()
  const spPr =
    `<c:spPr><a:solidFill><a:srgbClr val="${argb}"/></a:solidFill>` +
    `<a:ln><a:solidFill><a:srgbClr val="${argb}"/></a:solidFill></a:ln></c:spPr>`
  const seriesPattern = /<c:ser>[\s\S]*?<\/c:ser>/g
  let found = false
  // Document position, not c:idx: the sidecar numbers series by document
  // order, so a file whose c:idx values are shuffled must still edit the
  // series the UI showed as N.
  let position = -1
  const result = chartXml.replace(seriesPattern, (serXml) => {
    position += 1
    if (position !== seriesIndex) return serXml
    found = true
    // The series-level spPr precedes any c:dPt; don't clobber a point's spPr.
    const dPtAt = serXml.indexOf('<c:dPt>')
    const scope = dPtAt === -1 ? serXml : serXml.slice(0, dPtAt)
    const replaced = scope.replace(/<c:spPr>[\s\S]*?<\/c:spPr>/, () => spPr)
    if (replaced !== scope) return replaced + (dPtAt === -1 ? '' : serXml.slice(dPtAt))
    // No spPr yet: schema order puts it right after c:tx (or c:order when
    // the series has no name element).
    const anchor = /<c:tx>[\s\S]*?<\/c:tx>/.exec(serXml) ?? /<c:order val="[0-9]+"\/>/.exec(serXml)
    if (!anchor) throw new ChartEditError('Chart series is missing idx/order structure.')
    const insertAt = anchor.index + anchor[0].length
    return serXml.slice(0, insertAt) + spPr + serXml.slice(insertAt)
  })
  if (!found) throw new ChartEditError(`Chart has no series with index ${seriesIndex}.`)
  return result
}

/// Per-point fills via c:dPt on the series at the given document position.
function setPointColors(
  chartXml: string,
  seriesIndex: number,
  points: Record<string, string>,
): string {
  const seriesPattern = /<c:ser>[\s\S]*?<\/c:ser>/g
  let found = false
  let position = -1
  const result = chartXml.replace(seriesPattern, (serXml) => {
    position += 1
    if (position !== seriesIndex) return serXml
    found = true
    let next = serXml
    const entries = Object.entries(points)
      .map(([point, color]) => [Number(point), color] as const)
      .sort((a, b) => a[0] - b[0])
    for (const [point, color] of entries) {
      next = upsertPointColor(next, point, color)
    }
    return next
  })
  if (!found) throw new ChartEditError(`Chart has no series with index ${seriesIndex}.`)
  return result
}

function upsertPointColor(serXml: string, pointIndex: number, color: string): string {
  const fill = `<a:solidFill><a:srgbClr val="${color.slice(1).toUpperCase()}"/></a:solidFill>`
  const dPts = [...serXml.matchAll(/<c:dPt>[\s\S]*?<\/c:dPt>/g)]
  for (const dPt of dPts) {
    const idx = Number(/<c:idx val="([0-9]+)"\/>/.exec(dPt[0])?.[1])
    if (idx !== pointIndex) continue
    return (
      serXml.slice(0, dPt.index) +
      patchPointFill(dPt[0], fill) +
      serXml.slice(dPt.index + dPt[0].length)
    )
  }
  const element =
    `<c:dPt><c:idx val="${pointIndex}"/><c:bubble3D val="0"/>` + `<c:spPr>${fill}</c:spPr></c:dPt>`
  const insertAt = dPtInsertionIndex(serXml, dPts, pointIndex)
  return serXml.slice(0, insertAt) + element + serXml.slice(insertAt)
}

/// Where a new dPt with this idx belongs: ascending among existing dPt
/// entries, otherwise before dLbls/cat/val.
function dPtInsertionIndex(
  serXml: string,
  dPts: readonly RegExpExecArray[],
  pointIndex: number,
): number {
  let insertAt = -1
  for (const dPt of dPts) {
    const idx = Number(/<c:idx val="([0-9]+)"\/>/.exec(dPt[0])?.[1])
    if (idx > pointIndex) return dPt.index
    insertAt = dPt.index + dPt[0].length
  }
  if (insertAt !== -1) return insertAt
  const anchor = /<c:dLbls[>/]|<c:cat[>/]|<c:val[>/]|<c:xVal[>/]/.exec(serXml)
  if (!anchor) throw new ChartEditError('Chart series has no data element to anchor a point.')
  return anchor.index
}

/// Existing dPt: swap only the spPr fill, keeping every other child.
function patchPointFill(dPtXml: string, fill: string): string {
  const spPrMatch = /<c:spPr>([\s\S]*?)<\/c:spPr>/.exec(dPtXml)
  if (!spPrMatch || spPrMatch[1] === undefined) {
    const anchor = /<c:pictureOptions[\s/>]|<c:extLst[\s/>]/.exec(dPtXml)
    const insertAt = anchor ? anchor.index : dPtXml.lastIndexOf('</c:dPt>')
    return dPtXml.slice(0, insertAt) + `<c:spPr>${fill}</c:spPr>` + dPtXml.slice(insertAt)
  }
  const inner = spPrMatch[1]
  // Only touch the shape fill, never a fill nested inside a:ln.
  const lnAt = inner.search(/<a:ln[\s>]/)
  const head = lnAt === -1 ? inner : inner.slice(0, lnAt)
  const tail = lnAt === -1 ? '' : inner.slice(lnAt)
  const fillPattern =
    /<a:noFill\/>|<a:solidFill>[\s\S]*?<\/a:solidFill>|<a:gradFill[\s\S]*?<\/a:gradFill>|<a:pattFill[\s\S]*?<\/a:pattFill>|<a:blipFill[\s\S]*?<\/a:blipFill>|<a:grpFill\/>/
  const nextHead = fillPattern.test(head) ? head.replace(fillPattern, () => fill) : head + fill
  return dPtXml.replace(spPrMatch[0], () => `<c:spPr>${nextHead}${tail}</c:spPr>`)
}

function escapeXmlText(input: string): string {
  return input.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function escapeXmlAttribute(input: string): string {
  return escapeXmlText(input).replaceAll('"', '&quot;')
}
