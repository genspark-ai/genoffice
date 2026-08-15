/**
 * Chart insertion — writes the chart part (ppt/charts/chartN.xml) + Content_Types
 * Override + slide rels + graphicFrame fragment, going through appendRawElements to
 * reuse the existing chart parsing/rendering.
 *
 * The chartSpace template mirrors docx-engine's buildChartPartXml (data goes through
 * strCache/numCache caches, no embedded workbook attached; PowerPoint renders it
 * fine, but "Edit Data" is unavailable).
 */
import type { EmuRect, Slide } from './types'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'
import { relsPathFor } from './zip'
import { appendRawElements, type OpenedPptx } from './index'
import { nextCNvPrId } from './insert'

export type NewChartKind =
  | 'bar'
  | 'barStacked'
  /** Percent-stacked column (no insert entry point; used to preserve the type subdivision when rebuilding an externally created chart during edits) */
  | 'barPercentStacked'
  | 'line'
  | 'area'
  | 'pie'
  | 'doughnut'
  | 'scatter'
  | 'radar'
  /** Combo chart: first N-1 series as clustered columns, last series as a line (on the right secondary value axis) */
  | 'comboBarLine'
  /** 3-D pie (c:pie3DChart + c:view3D; this app renders the pseudo-3D projection, PowerPoint renders true 3-D) */
  | 'pie3D'
  /** 3-D clustered column (c:bar3DChart + c:view3D) */
  | 'bar3D'

/** Chart element/style toggles (unset = current defaults: legend at bottom, no gridlines, no data labels). */
export interface ChartStyleOptions {
  /** Legend position; 'none' = do not write c:legend */
  legendPos?: 'b' | 't' | 'r' | 'l' | 'none'
  /** Data labels (plot-level c:dLbls showVal) */
  dataLabels?: boolean
  /** Major gridlines on the value axis */
  gridlines?: boolean
  catAxisTitle?: string
  valAxisTitle?: string
  /** Gap between bars (% of bar width, c:gapWidth, PowerPoint default 150) */
  gapWidthPct?: number
}

export interface NewChartOptions extends ChartStyleOptions {
  kind: NewChartKind
  title?: string
  categories: string[]
  series: Array<{ name: string; values: number[] }>
  offset: EmuRect
  /** Bar direction (bar = horizontal bar chart; no insert entry point, used to preserve the direction when rebuilding an externally created chart during edits) */
  barDir?: 'col' | 'bar'
  /** Per-point fills, [seriesIdx][pointIdx] (sparse; written as <c:dPt>, wins over the series color) */
  pointColors?: Array<Array<string | undefined> | undefined>
}

const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
const CHART_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const colLetter = (i: number) => String.fromCharCode(66 + i) // B, C, D…

function strCacheXml(values: string[], f: string): string {
  return (
    `<c:strRef><c:f>${escapeXmlText(f)}</c:f><c:strCache><c:ptCount val="${values.length}"/>` +
    values
      .map((v, i) => (v === '' ? '' : `<c:pt idx="${i}"><c:v>${escapeXmlText(v)}</c:v></c:pt>`))
      .join('') +
    '</c:strCache></c:strRef>'
  )
}

function numCacheXml(values: (number | null | undefined)[], f: string): string {
  return (
    `<c:numRef><c:f>${escapeXmlText(f)}</c:f><c:numCache><c:formatCode>General</c:formatCode>` +
    `<c:ptCount val="${values.length}"/>` +
    values
      .map((v, i) =>
        v == null || !Number.isFinite(v) ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`,
      )
      .join('') +
    '</c:numCache></c:numRef>'
  )
}

/** Data label fragment (dLbls follows ser in the schema; shared by all chart types). */
const DLBLS_XML =
  '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>' +
  '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'

/** Axis title fragment (value axis vertical, rot=-5400000). */
function axTitleXml(text: string, vertical: boolean): string {
  return (
    '<c:title><c:tx><c:rich>' +
    `<a:bodyPr${vertical ? ' rot="-5400000" vert="horz"' : ''}/><a:lstStyle/>` +
    `<a:p><a:r><a:t>${escapeXmlText(text)}</a:t></a:r></a:p>` +
    '</c:rich></c:tx><c:overlay val="0"/></c:title>'
  )
}

/** Full c:chartSpace part XML. */
export function buildChartSpaceXml(opts: NewChartOptions): string {
  const rows = opts.categories.length
  const dLbls = opts.dataLabels ? DLBLS_XML : ''
  const grid = opts.gridlines ? '<c:majorGridlines/>' : ''
  const catTitle = opts.catAxisTitle ? axTitleXml(opts.catAxisTitle, false) : ''
  const valTitle = opts.valAxisTitle ? axTitleXml(opts.valAxisTitle, true) : ''
  const gapWidth =
    opts.gapWidthPct != null ? `<c:gapWidth val="${Math.round(opts.gapWidthPct)}"/>` : ''
  // Empty names stay empty: no <c:tx> for an unnamed series, no <c:cat> when every category name is empty
  const txXml = (name: string, i: number) =>
    name === '' ? '' : `<c:tx>${strCacheXml([name], `Sheet1!$${colLetter(i)}$1`)}</c:tx>`
  const catXml = opts.categories.some((c) => c !== '')
    ? `<c:cat>${strCacheXml(opts.categories, `Sheet1!$A$2:$A$${rows + 1}`)}</c:cat>`
    : ''
  // Per-point <c:dPt> fills (schema position: after tx/spPr, before dLbls/cat/val)
  const dPtXml = (i: number): string => {
    const colors = opts.pointColors?.[i]
    if (!colors) return ''
    let out = ''
    colors.forEach((c, pi) => {
      if (!c) return
      const hex = c.replace('#', '').toUpperCase()
      out +=
        `<c:dPt><c:idx val="${pi}"/><c:bubble3D val="0"/>` +
        `<c:spPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></c:spPr></c:dPt>`
    })
    return out
  }
  // Single-series fragment (idx/order use global ordinals so colors stay in order when a combo chart splits into two plots)
  const serXml = (ser: { name: string; values: number[] }, i: number): string => {
    const col = colLetter(i)
    return (
      `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` +
      txXml(ser.name, i) +
      dPtXml(i) +
      catXml +
      `<c:val>${numCacheXml(ser.values.slice(0, rows), `Sheet1!$${col}$2:$${col}$${rows + 1}`)}</c:val>` +
      '</c:ser>'
    )
  }
  const sers = opts.series.map(serXml).join('')

  let plot: string
  if (opts.kind === 'comboBarLine') {
    // Combo chart (columns + line): with ≥2 series the last is a line, the rest are
    // clustered columns; a single series degrades to plain columns.
    // Columns use the primary axes (category axis 1111… + left value axis 2222…), the line
    // uses the secondary axes (hidden category axis 3333… + right secondary value axis 4444…),
    // matching how PowerPoint writes "combo chart + secondary axis".
    const lineCount = opts.series.length >= 2 ? 1 : 0
    const barEnd = opts.series.length - lineCount
    const barSers = opts.series.slice(0, barEnd).map(serXml).join('')
    const lineSers = opts.series
      .slice(barEnd)
      .map((ser, k) => serXml(ser, barEnd + k))
      .join('')
    plot =
      '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
      `${barSers}${dLbls}${gapWidth}<c:axId val="111111111"/><c:axId val="222222222"/></c:barChart>` +
      (lineSers
        ? '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' +
          `${lineSers}${dLbls}<c:marker val="1"/><c:axId val="333333333"/><c:axId val="444444444"/></c:lineChart>`
        : '') +
      '<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      `<c:delete val="0"/><c:axPos val="b"/>${catTitle}<c:crossAx val="222222222"/></c:catAx>` +
      '<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      `<c:delete val="0"/><c:axPos val="l"/>${grid}${valTitle}<c:crossAx val="111111111"/></c:valAx>` +
      (lineSers
        ? '<c:valAx><c:axId val="444444444"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
          '<c:delete val="0"/><c:axPos val="r"/><c:crossAx val="333333333"/><c:crosses val="max"/></c:valAx>' +
          '<c:catAx><c:axId val="333333333"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
          '<c:delete val="1"/><c:axPos val="b"/><c:crossAx val="444444444"/></c:catAx>'
        : '')
  } else if (opts.kind === 'pie') {
    plot = `<c:pieChart><c:varyColors val="1"/>${sers}${dLbls}<c:firstSliceAng val="0"/></c:pieChart>`
  } else if (opts.kind === 'pie3D') {
    plot = `<c:pie3DChart><c:varyColors val="1"/>${sers}${dLbls}</c:pie3DChart>`
  } else if (opts.kind === 'bar3D') {
    // bar3D takes three axes (category / value / series); the series axis is required by the schema
    const horizontal = opts.barDir === 'bar'
    plot =
      `<c:bar3DChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
      `${sers}${dLbls}${gapWidth}<c:shape val="box"/>` +
      '<c:axId val="111111111"/><c:axId val="222222222"/><c:axId val="333333333"/></c:bar3DChart>' +
      '<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      `<c:delete val="0"/><c:axPos val="${horizontal ? 'l' : 'b'}"/>${catTitle}<c:crossAx val="222222222"/></c:catAx>` +
      '<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      `<c:delete val="0"/><c:axPos val="${horizontal ? 'b' : 'l'}"/>${grid}${valTitle}<c:crossAx val="111111111"/></c:valAx>` +
      '<c:serAx><c:axId val="333333333"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:serAx>'
  } else if (opts.kind === 'doughnut') {
    plot =
      `<c:doughnutChart><c:varyColors val="1"/>${sers}${dLbls}` +
      '<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>'
  } else if (opts.kind === 'scatter') {
    // Scatter (XY): x values come from categories (numeric strings use their value,
    // otherwise ordinals 1..n), y values from the series values;
    // series data uses c:xVal/c:yVal, with dual value axes (x axis axPos=b, y axis axPos=l)
    const xs = opts.categories.map((c, i) => {
      const v = Number(c)
      return c.trim() !== '' && Number.isFinite(v) ? v : i + 1
    })
    const scatterSers = opts.series
      .map((ser, i) => {
        const col = colLetter(i)
        return (
          `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>` +
          txXml(ser.name, i) +
          `<c:xVal>${numCacheXml(xs, `Sheet1!$A$2:$A$${rows + 1}`)}</c:xVal>` +
          `<c:yVal>${numCacheXml(ser.values.slice(0, rows), `Sheet1!$${col}$2:$${col}$${rows + 1}`)}</c:yVal>` +
          '<c:smooth val="0"/></c:ser>'
        )
      })
      .join('')
    plot =
      '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>' +
      `${scatterSers}${dLbls}<c:axId val="111111111"/><c:axId val="222222222"/></c:scatterChart>` +
      '<c:valAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      `<c:delete val="0"/><c:axPos val="b"/>${catTitle}<c:crossAx val="222222222"/></c:valAx>` +
      '<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      `<c:delete val="0"/><c:axPos val="l"/>${grid}${valTitle}<c:crossAx val="111111111"/></c:valAx>`
  } else {
    // Horizontal bar chart (barDir=bar): category axis on the left, value axis at the bottom (matches how PowerPoint writes it)
    const isBarKind =
      opts.kind === 'bar' || opts.kind === 'barStacked' || opts.kind === 'barPercentStacked'
    const horizontal = isBarKind && opts.barDir === 'bar'
    const axes =
      `<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
      `<c:delete val="0"/><c:axPos val="${horizontal ? 'l' : 'b'}"/>${catTitle}<c:crossAx val="222222222"/></c:catAx>` +
      '<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      `<c:delete val="0"/><c:axPos val="${horizontal ? 'b' : 'l'}"/>${grid}${valTitle}<c:crossAx val="111111111"/></c:valAx>`
    const axIds = '<c:axId val="111111111"/><c:axId val="222222222"/>'
    let inner: string
    if (opts.kind === 'radar') {
      // Radar: categories are vertices (clockwise from 12 o'clock), standard style (unfilled lines)
      inner = `<c:radarChart><c:radarStyle val="standard"/><c:varyColors val="0"/>${sers}${dLbls}${axIds}</c:radarChart>`
    } else if (opts.kind === 'line') {
      inner = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${dLbls}<c:marker val="1"/>${axIds}</c:lineChart>`
    } else if (opts.kind === 'area') {
      inner = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${dLbls}${axIds}</c:areaChart>`
    } else {
      const grouping =
        opts.kind === 'barPercentStacked'
          ? 'percentStacked'
          : opts.kind === 'barStacked'
            ? 'stacked'
            : 'clustered'
      const overlap = grouping === 'clustered' ? '' : '<c:overlap val="100"/>'
      inner =
        `<c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>` +
        `${sers}${dLbls}${gapWidth}${overlap}${axIds}</c:barChart>`
    }
    plot = inner + axes
  }

  const title = opts.title
    ? '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>' +
      `<a:t>${escapeXmlText(opts.title)}</a:t></a:r></a:p></c:rich></c:tx>` +
      '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
    : '<c:autoTitleDeleted val="1"/>'

  // 3-D view settings (schema position: after the title block, before plotArea);
  // rotX/rotY follow PowerPoint's defaults for each type
  const view3D =
    opts.kind === 'pie3D'
      ? '<c:view3D><c:rotX val="30"/><c:rotY val="0"/><c:rAngAx val="0"/><c:perspective val="30"/></c:view3D>'
      : opts.kind === 'bar3D'
        ? '<c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="1"/><c:perspective val="30"/></c:view3D>'
        : ''

  const legendPos = opts.legendPos ?? 'b'
  const legend =
    legendPos === 'none'
      ? ''
      : `<c:legend><c:legendPos val="${legendPos}"/><c:overlay val="0"/></c:legend>`
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">` +
    `<c:chart>${title}${view3D}<c:plotArea><c:layout/>${plot}</c:plotArea>` +
    legend +
    '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
    '</c:chartSpace>'
  )
}

/**
 * Insert a chart: part surgery + graphicFrame fragment append + reparse.
 * Returns the new slide and element id (all ids on the slide refreshed).
 */
export function addChart(
  opened: OpenedPptx,
  slideIndex: number,
  opts: NewChartOptions,
): { slide: Slide; elementId: string } | null {
  const { archive } = opened
  const slide = opened.deck.slides[slideIndex]
  if (!slide || !opts.categories.length || !opts.series.length) return null

  // 1) chart part: number = current max + 1
  let maxNum = 0
  for (const path of archive.entries.keys()) {
    const m = /^ppt\/charts\/chart(\d+)\.xml$/.exec(path)
    if (m) maxNum = Math.max(maxNum, Number(m[1]))
  }
  const chartPath = `ppt/charts/chart${maxNum + 1}.xml`
  archive.entries.set(chartPath, Buffer.from(buildChartSpaceXml(opts), 'utf8'))

  // 2) [Content_Types].xml Override
  const ctPath = '[Content_Types].xml'
  const ct = archive.readText(ctPath)
  if (ct && !ct.includes(`PartName="/${chartPath}"`)) {
    const override = `<Override PartName="/${chartPath}" ContentType="${CHART_CONTENT_TYPE}"/>`
    archive.entries.set(ctPath, Buffer.from(ct.replace('</Types>', `${override}</Types>`), 'utf8'))
  }

  // 3) slide rels
  const relsPath = relsPathFor(slide.path)
  const rels =
    archive.readText(relsPath) ??
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  let maxRid = 0
  for (const m of rels.matchAll(/Id="rId(\d+)"/g)) maxRid = Math.max(maxRid, Number(m[1]))
  const rid = `rId${maxRid + 1}`
  const relXml = `<Relationship Id="${rid}" Type="${CHART_REL_TYPE}" Target="../charts/chart${maxNum + 1}.xml"/>`
  archive.entries.set(
    relsPath,
    Buffer.from(rels.replace('</Relationships>', `${relXml}</Relationships>`), 'utf8'),
  )

  // 4) graphicFrame fragment + append reparse
  const id = nextCNvPrId(slide)
  const o = opts.offset
  const name = opts.title ? `Chart ${id} - ${opts.title}` : `Chart ${id}`
  // descr="aislides-chart" marks charts created by this app (like the ink marker); recognized as editable charts on reopen
  const frameXml =
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${escapeXmlAttr(name)}" descr="aislides-chart"/>` +
    '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
    `<p:xfrm><a:off x="${o.x}" y="${o.y}"/><a:ext cx="${o.cx}" cy="${o.cy}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${C_NS}">` +
    `<c:chart xmlns:c="${C_NS}" xmlns:r="${R_NS}" r:id="${rid}"/>` +
    '</a:graphicData></a:graphic></p:graphicFrame>'

  const r = appendRawElements(opened, slideIndex, [frameXml])
  return r ? { slide: r.slide, elementId: r.elementIds[r.elementIds.length - 1]! } : null
}
