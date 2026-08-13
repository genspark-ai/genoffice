/**
 * Chart parsing (ppt/charts/chartN.xml → ChartModel).
 *
 * Read-only semantic parsing: on a slide a chart is a separate part referenced by
 * a <p:graphicFrame>; byte fidelity is guaranteed by passing the graphicFrame's
 * anchor bytes through verbatim, and the chart part is never rewritten.
 *
 * Supports data + explicit styling (series colors, axis label styles, gridlines)
 * for lineChart / barChart (incl. horizontal bars) / pieChart (incl. doughnut) /
 * areaChart / scatterChart / radarChart; missing styles fall back to the theme
 * palette. bar/area/line can be combined in one plotArea (column+line combo);
 * combined series carry plotKind and the primary type is picked bar > area > line;
 * series on the secondary value axis (axPos=r, not deleted, matched by the plot's
 * c:axId) carry secondaryAxis, with the secondary axis style/range parsed into
 * valAxis2.
 */
import { XMLParser } from 'fast-xml-parser'
import { type Theme } from './theme'
import { resolveColorNode } from './color'
import type { Fill } from './types'

const chartParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
  parseTagValue: false,
  isArray: (name) => ['c:ser', 'c:pt', 'c:lvl', 'c:dPt'].includes(name),
})

export type ChartKind = 'line' | 'bar' | 'pie' | 'area' | 'scatter' | 'radar' | 'unknown'

export interface ChartSeries {
  name?: string
  /** Series main color #RRGGBB (explicit spPr color; render layer fills in from the theme palette otherwise) */
  color?: string
  values: Array<number | null>
  /** Combo chart (e.g. bar+line): the plot type this series belongs to; default = ChartModel.kind */
  plotKind?: 'line' | 'bar' | 'area'
  /** Scatter: x values (c:xVal numeric cache; y values in values). Empty → render layer uses ordinals 1..n */
  xValues?: Array<number | null>
  /** Line: smoothed curve */
  smooth?: boolean
  /** Whether to draw data point markers (line defaults to false; scatter/radar default
   *  comes from the style; only set when <c:marker><c:symbol> is explicit) */
  marker?: boolean
  /** Explicit per-point colors <c:dPt> (common for pies; render layer palette otherwise) */
  pointColors?: Array<string | undefined>
  /** Combo dual axes: this series is on the secondary value axis (independent right-side range; decided by the plot's c:axId) */
  secondaryAxis?: boolean
}

export interface ChartAxisStyle {
  /** Explicit min/max (render layer computes a nice range from the data otherwise) */
  min?: number
  max?: number
  labelColor?: string
  labelSizePt?: number
  lineColor?: string
  gridColor?: string
  gridDash?: boolean
  title?: string
  /** <c:orientation val="maxMin"/>: categories/values reversed (common for bar charts with the first category on top) */
  reversed?: boolean
}

export interface ChartModel {
  kind: ChartKind
  /** bar only: col = vertical columns, bar = horizontal bars */
  barDir?: 'col' | 'bar'
  grouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard'
  /** Gap between bars (% of bar width, c:gapWidth, default 150) */
  gapWidthPct?: number
  categories: string[]
  series: ChartSeries[]
  /** Legend position (undefined when there is no c:legend) */
  legendPos?: 't' | 'b' | 'l' | 'r' | 'tr'
  valAxis?: ChartAxisStyle
  /** Secondary value axis (right side, combo column+line dual axes; undefined without a right value axis or style info) */
  valAxis2?: ChartAxisStyle
  catAxis?: ChartAxisStyle
  /** Doughnut hole (% of radius, c:holeSize; pie = 0) */
  holePct?: number
  /** First slice start angle (degrees, 12 o'clock = 0, clockwise; c:firstSliceAng) */
  firstSliceAngDeg?: number
  /** Scatter style (c:scatterStyle: line/lineMarker/marker/smooth/smoothMarker/none) */
  scatterStyle?: string
  /** Radar style (c:radarStyle) */
  radarStyle?: 'standard' | 'marker' | 'filled'
  /** Data labels (showVal or showPercent on plot/ser-level c:dLbls) */
  dataLabels?: boolean
  /** Data labels show percentages only (showPercent only, common for pies) */
  dataLabelsPct?: boolean
  /** Chart title (concatenated rich text of c:chart/c:title) */
  title?: string
  /** chartSpace-level <c:spPr> fill (whole-chart background, e.g. picture fill) */
  bgFill?: Fill
  /** Theme accent1..6 resolved to #RRGGBB (default series/wedge colors, PowerPoint varyColors order) */
  themePalette?: string[]
  /** chartSpace-level <c:txPr> default text size (pt); chart text without its own txPr uses this */
  defaultTextPt?: number
  /** 3D chart type (pie3D/bar3D/…): render layer draws a pseudo-3D look on the 2D pipeline */
  pseudo3D?: boolean
  /** c:view3D rotX (degrees) — controls the pie tilt / bar extrusion feel */
  rotXDeg?: number
}

/** Parse one chartN.xml. Returns null for unrecognized plot types (caller falls back to a placeholder chip). */
export function parseChartXml(
  xml: string,
  theme?: Theme,
  /** DrawingML fill resolver (injected by parse.ts so chart-part blip rIds resolve correctly) */
  resolveFill?: (spPr: unknown) => Fill | undefined,
): ChartModel | null {
  let doc: any
  try {
    doc = chartParser.parse(xml)
  } catch {
    return null
  }
  const chart = doc['c:chartSpace']?.['c:chart']
  const plotArea = chart?.['c:plotArea']
  if (!plotArea) return null

  // Cartesian types (bar/area/line) may coexist combined (e.g. column+line combo); pie/scatter/radar stand alone.
  // 3D variants map onto the 2D pipelines (same data/palette/legend); the render layer adds a
  // pseudo-3D look (elliptical pie with a rim, extruded bars) driven by pseudo3D + c:view3D rotX.
  const cartesian: Array<{ kind: 'bar' | 'area' | 'line'; plot: any }> = []
  const barPlot = plotArea['c:barChart'] ?? plotArea['c:bar3DChart']
  const areaPlot = plotArea['c:areaChart'] ?? plotArea['c:area3DChart']
  const linePlot = plotArea['c:lineChart'] ?? plotArea['c:line3DChart']
  if (barPlot) cartesian.push({ kind: 'bar', plot: barPlot })
  if (areaPlot) cartesian.push({ kind: 'area', plot: areaPlot })
  if (linePlot) cartesian.push({ kind: 'line', plot: linePlot })

  const piePlot = plotArea['c:pieChart'] ?? plotArea['c:pie3DChart'] ?? plotArea['c:doughnutChart']

  const is3D = !!(
    plotArea['c:bar3DChart'] ||
    plotArea['c:area3DChart'] ||
    plotArea['c:line3DChart'] ||
    plotArea['c:pie3DChart']
  )

  let kind: ChartKind
  let plot: any
  if (cartesian.length) {
    // Primary type is the first (bar > area > line): axis/bar params read from it; other combo series carry plotKind
    kind = cartesian[0]!.kind
    plot = cartesian[0]!.plot
  } else if (piePlot) {
    kind = 'pie'
    plot = piePlot
  } else if (plotArea['c:scatterChart']) {
    kind = 'scatter'
    plot = plotArea['c:scatterChart']
  } else if (plotArea['c:radarChart']) {
    kind = 'radar'
    plot = plotArea['c:radarChart']
  } else {
    return null
  }

  // Extract value axis nodes up front: combo dual axes need the secondary value
  // axis's axId (axPos=r and not deleted) first, so series parsing can decide
  // "primary or secondary axis" by the owning plot's c:axId
  const valAxRaw = plotArea['c:valAx']
  const valAxes: any[] = Array.isArray(valAxRaw) ? valAxRaw : valAxRaw ? [valAxRaw] : []
  const secValAxNode =
    kind !== 'scatter' && cartesian.length > 1
      ? valAxes.find((a) => a?.['c:axPos']?.['@_val'] === 'r' && a?.['c:delete']?.['@_val'] !== '1')
      : undefined
  const secAxId: string | undefined = secValAxNode?.['c:axId']?.['@_val']
  // Axes attached to a plot node (two c:axIds: category axis + value axis)
  const plotAxIds = (plotNode: any): string[] => {
    const raw = plotNode?.['c:axId']
    const arr: any[] = Array.isArray(raw) ? raw : raw ? [raw] : []
    return arr.map((a) => a?.['@_val']).filter((v): v is string => v != null)
  }

  const series: ChartSeries[] = []
  let categories: string[] = []
  const parsePlotSeries = (
    plotNode: any,
    plotKind: ChartKind,
    tagPlotKind: boolean,
    secondary = false,
  ) => {
    const sersRaw = plotNode['c:ser']
    const sers: any[] = Array.isArray(sersRaw) ? sersRaw : sersRaw ? [sersRaw] : []
    for (const ser of sers) {
      // Scatter: y values in c:yVal, x values in c:xVal; other types use c:val
      const s: ChartSeries = {
        values: readNumPoints(plotKind === 'scatter' ? ser['c:yVal'] : ser['c:val']),
      }
      if (tagPlotKind) s.plotKind = plotKind as 'line' | 'bar' | 'area'
      if (secondary) s.secondaryAxis = true
      if (plotKind === 'scatter') {
        const xs = readNumPoints(ser['c:xVal'])
        if (xs.length) s.xValues = xs
      }
      const name = readStrPoints(ser['c:tx'])[0]
      if (name != null) s.name = name
      const color = serColor(ser, theme)
      if (color) s.color = color
      if (ser['c:smooth']?.['@_val'] === '1') s.smooth = true
      const markerSym = ser['c:marker']?.['c:symbol']?.['@_val']
      if (plotKind === 'line') s.marker = markerSym != null && markerSym !== 'none'
      // scatter/radar: default marker decided by style; only set for explicit symbol (none → false)
      else if ((plotKind === 'scatter' || plotKind === 'radar') && markerSym != null)
        s.marker = markerSym !== 'none'
      // Per-data-point colors (one color per pie slice)
      const dPts: any[] = ser['c:dPt'] ?? []
      if (dPts.length) {
        const pointColors: Array<string | undefined> = []
        for (const dPt of dPts) {
          const idx = parseInt(dPt['c:idx']?.['@_val'], 10)
          if (Number.isNaN(idx)) continue
          const c = resolveColorNode(dPt['c:spPr']?.['a:solidFill'], theme)
          if (c != null) pointColors[idx] = c
        }
        if (pointColors.length) s.pointColors = pointColors
      }
      series.push(s)
      // Categories: take the first non-empty series' cat
      if (!categories.length) categories = readStrPoints(ser['c:cat'])
    }
  }
  if (cartesian.length > 1) {
    for (const c of cartesian)
      parsePlotSeries(c.plot, c.kind, true, secAxId != null && plotAxIds(c.plot).includes(secAxId))
  } else parsePlotSeries(plot, kind, false)
  if (!series.length) return null
  if (!categories.length) {
    // With no category cache, keep names empty (length from the longest series); never inject placeholders
    const n = Math.max(...series.map((s) => s.values.length), 0)
    categories = Array.from({ length: n }, () => '')
  }

  const model: ChartModel = { kind, categories, series }

  if (is3D) {
    model.pseudo3D = true
    const rotX = parseInt(chart['c:view3D']?.['c:rotX']?.['@_val'], 10)
    if (Number.isFinite(rotX)) model.rotXDeg = rotX
  }

  if (kind === 'bar') {
    const dir = plot['c:barDir']?.['@_val']
    model.barDir = dir === 'bar' ? 'bar' : 'col'
    const grouping = plot['c:grouping']?.['@_val']
    if (grouping) model.grouping = grouping
    const gap = plot['c:gapWidth']?.['@_val']
    model.gapWidthPct = gap != null ? parseInt(gap, 10) : 150
  }

  if (kind === 'pie') {
    const hole = plot['c:holeSize']?.['@_val']
    model.holePct = hole != null ? parseInt(hole, 10) || 0 : plotArea['c:doughnutChart'] ? 50 : 0
    const first = plot['c:firstSliceAng']?.['@_val']
    if (first != null) model.firstSliceAngDeg = parseInt(first, 10) || 0
  }

  if (kind === 'scatter') {
    const st = plot['c:scatterStyle']?.['@_val']
    if (st) model.scatterStyle = String(st)
  }

  if (kind === 'radar') {
    const st = plot['c:radarStyle']?.['@_val']
    model.radarStyle = st === 'filled' ? 'filled' : st === 'marker' ? 'marker' : 'standard'
  }

  const legendPos = chart['c:legend']?.['c:legendPos']?.['@_val']
  if (chart['c:legend']) model.legendPos = (legendPos as ChartModel['legendPos']) ?? 'r'

  // Whole-chart background (chartSpace-level spPr), e.g. picture fill
  const bgFill = resolveFill?.(doc['c:chartSpace']?.['c:spPr'])
  if (bgFill && bgFill.type !== 'none') model.bgFill = bgFill

  // Theme accents: default series/wedge colors follow the file's theme, not a fixed palette
  if (theme) {
    const accents = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
      .map((k) => theme.colors?.[k])
      .filter((c): c is string => !!c)
    if (accents.length === 6) model.themePalette = accents
  }
  // chartSpace-level default text size (hundredths of a pt)
  const txP = doc['c:chartSpace']?.['c:txPr']?.['a:p']
  const defSz = parseInt((Array.isArray(txP) ? txP[0] : txP)?.['a:pPr']?.['a:defRPr']?.['@_sz'], 10)
  if (Number.isFinite(defSz) && defSz > 0) model.defaultTextPt = defSz / 100

  // Title text: rich text, or the cached cell-linked string (c:tx/c:strRef)
  const chartTitle =
    collectText(chart['c:title']?.['c:tx']?.['c:rich']) ||
    readStrPoints(chart['c:title']?.['c:tx'])[0]
  if (chartTitle) model.title = chartTitle
  // Auto title: a <c:title> with no c:tx at all (and autoTitleDeleted != 1) takes the sole series name
  else if (
    chart['c:title'] &&
    !chart['c:title']?.['c:tx'] &&
    chart['c:autoTitleDeleted']?.['@_val'] !== '1' &&
    series.length === 1 &&
    series[0]!.name
  ) {
    model.title = series[0]!.name
  }

  // Data labels: plot-level or any series-level c:dLbls (delete=1 counts as none)
  const dLblsInfo = (owner: any): { on: boolean; pct: boolean } => {
    const d = owner?.['c:dLbls']
    if (!d || typeof d !== 'object' || d['c:delete']?.['@_val'] === '1')
      return { on: false, pct: false }
    const showVal = d['c:showVal']?.['@_val'] === '1'
    const showPct = d['c:showPercent']?.['@_val'] === '1'
    return { on: showVal || showPct, pct: showPct && !showVal }
  }
  const dLblOwners: any[] = (cartesian.length > 1 ? cartesian.map((c) => c.plot) : [plot]).flatMap(
    (p) => [
      p,
      ...((Array.isArray(p['c:ser']) ? p['c:ser'] : p['c:ser'] ? [p['c:ser']] : []) as any[]),
    ],
  )
  const found = dLblOwners.map(dLblsInfo).find((r) => r.on)
  if (found) {
    model.dataLabels = true
    if (found.pct) model.dataLabelsPct = true
  }

  // Axes: scatter charts have dual value axes (x at the bottom axPos=b, y on the left); the x axis goes in the catAxis slot
  if (kind === 'scatter' && valAxes.length >= 2) {
    const xAxNode = valAxes.find((a) => a?.['c:axPos']?.['@_val'] === 'b') ?? valAxes[0]
    const yAxNode = valAxes.find((a) => a !== xAxNode) ?? valAxes[1]
    const xAx = parseAxis(xAxNode, theme)
    if (xAx) model.catAxis = xAx
    const yAx = parseAxis(yAxNode, theme)
    if (yAx) model.valAxis = yAx
  } else {
    // A combo chart (column + line secondary axis) has two axis pairs in plotArea:
    // the value axis is the left primary (axPos=l), the category axis is the
    // non-deleted one (c:delete≠1); the secondary system's hidden category axis is skipped
    const valAxNode = valAxes.find((a) => a?.['c:axPos']?.['@_val'] === 'l') ?? valAxes[0]
    const valAx = parseAxis(valAxNode, theme)
    if (valAx) model.valAxis = valAx
    // Secondary value axis (combo dual axes): min/max/style handed to the render layer to draw the right-side ticks
    if (secValAxNode) {
      const valAx2 = parseAxis(secValAxNode, theme)
      if (valAx2) model.valAxis2 = valAx2
    }
    const catAxRaw = plotArea['c:catAx']
    const catAxes: any[] = Array.isArray(catAxRaw) ? catAxRaw : catAxRaw ? [catAxRaw] : []
    const catAxNode = catAxes.find((a) => a?.['c:delete']?.['@_val'] !== '1') ?? catAxes[0]
    const catAx = parseAxis(catAxNode, theme)
    if (catAx) model.catAxis = catAx
  }

  return model
}

/** Numeric cache inside <c:val>/<c:cat>/<c:tx> → number[] (idx order kept, empty points null). */
function readNumPoints(node: any): Array<number | null> {
  const cache = node?.['c:numRef']?.['c:numCache'] ?? node?.['c:numLit']
  if (!cache) return []
  return readPoints(cache).map((v) => {
    if (v == null || v === '') return null
    const n = Number(v)
    return Number.isFinite(n) ? n : null
  })
}

/** String cache (strRef/strCache or the innermost lvl of multiLvlStrRef) → string[]. */
function readStrPoints(node: any): string[] {
  const strCache = node?.['c:strRef']?.['c:strCache']
  if (strCache) return readPoints(strCache).map((v) => v ?? '')
  const multi = node?.['c:multiLvlStrRef']?.['c:multiLvlStrCache']
  if (multi) {
    const lvls: any[] = Array.isArray(multi['c:lvl'])
      ? multi['c:lvl']
      : multi['c:lvl']
        ? [multi['c:lvl']]
        : []
    // The innermost (first lvl) holds the leaf categories
    if (lvls.length) return readPoints(lvls[0]).map((v) => v ?? '')
  }
  const numCache = node?.['c:numRef']?.['c:numCache']
  if (numCache) return readPoints(numCache).map((v) => v ?? '')
  return []
}

/** c:pt list → value array ordered by idx. */
function readPoints(cache: any): Array<string | null> {
  const ptsRaw = cache?.['c:pt']
  const pts: any[] = Array.isArray(ptsRaw) ? ptsRaw : ptsRaw ? [ptsRaw] : []
  const count = cache?.['c:ptCount']?.['@_val']
  const n = count != null ? parseInt(count, 10) : pts.length
  const out: Array<string | null> = new Array(Math.max(n, pts.length)).fill(null)
  for (const pt of pts) {
    const idx = parseInt(pt['@_idx'], 10) || 0
    const v = pt['c:v']
    out[idx] = typeof v === 'string' ? v : v != null ? String(v['#text'] ?? v) : null
  }
  return out
}

/** Series main color: ln stroke first (lines), otherwise solidFill (bars/pies). */
function serColor(ser: any, theme?: Theme): string | undefined {
  const spPr = ser['c:spPr']
  if (!spPr) return undefined
  const lnColor = resolveColorNode(spPr['a:ln']?.['a:solidFill'], theme)
  const fillColor = resolveColorNode(spPr['a:solidFill'], theme)
  return lnColor ?? fillColor
}

function parseAxis(ax: any, theme?: Theme): ChartAxisStyle | undefined {
  if (!ax || typeof ax !== 'object') return undefined
  const out: ChartAxisStyle = {}
  const scaling = ax['c:scaling']
  if (scaling?.['c:min']?.['@_val'] != null) out.min = Number(scaling['c:min']['@_val'])
  if (scaling?.['c:max']?.['@_val'] != null) out.max = Number(scaling['c:max']['@_val'])
  if (scaling?.['c:orientation']?.['@_val'] === 'maxMin') out.reversed = true
  const defRPr =
    ax['c:txPr']?.['a:p']?.[0]?.['a:pPr']?.['a:defRPr'] ??
    ax['c:txPr']?.['a:p']?.['a:pPr']?.['a:defRPr']
  if (defRPr) {
    const c = resolveColorNode(defRPr['a:solidFill'], theme)
    if (c) out.labelColor = c
    if (defRPr['@_sz']) out.labelSizePt = parseInt(defRPr['@_sz'], 10) / 100
  }
  const lineColor = resolveColorNode(ax['c:spPr']?.['a:ln']?.['a:solidFill'], theme)
  if (lineColor) out.lineColor = lineColor
  // A self-closing <c:majorGridlines/> (no spPr) parses to an empty string; still counts as having gridlines
  const grid = ax['c:majorGridlines']
  if (grid !== undefined) {
    const spPr = typeof grid === 'object' ? grid['c:spPr'] : undefined
    const gc = resolveColorNode(spPr?.['a:ln']?.['a:solidFill'], theme)
    out.gridColor = gc ?? '#E6E6E6'
    if (spPr?.['a:ln']?.['a:prstDash']?.['@_val'] === 'dash') out.gridDash = true
  }
  // Axis title (all a:t inside c:title/c:tx/c:rich concatenated)
  const title = collectText(ax['c:title']?.['c:tx']?.['c:rich'])
  if (title) out.title = title
  return Object.keys(out).length ? out : undefined
}

/** Collect all a:t inside a rich text node. */
function collectText(rich: any): string | undefined {
  if (!rich) return undefined
  const paras: any[] = Array.isArray(rich['a:p']) ? rich['a:p'] : rich['a:p'] ? [rich['a:p']] : []
  const parts: string[] = []
  for (const p of paras) {
    const runs: any[] = Array.isArray(p['a:r']) ? p['a:r'] : p['a:r'] ? [p['a:r']] : []
    for (const r of runs) {
      const t = r['a:t']
      parts.push(typeof t === 'string' ? t : String(t?.['#text'] ?? ''))
    }
  }
  const s = parts.join('')
  return s.trim() ? s : undefined
}
