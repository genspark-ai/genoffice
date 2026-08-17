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
import { resolveColorNode, scaleLuminance } from './color'
import type { Fill } from './types'

const chartParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
  parseTagValue: false,
  isArray: (name) => ['c:ser', 'c:pt', 'c:lvl', 'c:dPt'].includes(name),
})

export type ChartKind =
  'line' | 'bar' | 'pie' | 'area' | 'scatter' | 'radar' | 'funnel' | 'sunburst' | 'unknown'

export interface ChartSeries {
  name?: string
  /** Series main color #RRGGBB (explicit spPr color; render layer fills in from the theme palette otherwise) */
  color?: string
  values: Array<number | null>
  /** Combo chart (e.g. bar+line): the plot type this series belongs to; default = ChartModel.kind */
  plotKind?: 'line' | 'bar' | 'area'
  /** Scatter: x values (c:xVal numeric cache; y values in values). Empty → render layer uses ordinals 1..n */
  xValues?: Array<number | null>
  /** Bubble: per-point sizes (c:bubbleSize); marker radius ∝ √size (area represents value) */
  bubbleSizes?: Array<number | null>
  /** Line: smoothed curve */
  smooth?: boolean
  /** c:idx: automatic palette colors key off this, not document order */
  paletteIdx?: number
  /** Per-series data-label visibility (own c:dLbls, else the plot-level c:dLbls); undefined = fall back to the chart-level flag */
  dataLabels?: boolean
  /** Line/scatter series stroke dash (c:spPr a:prstDash, non-solid only) */
  dash?: string
  /** Line/scatter series stroke width (pt, c:spPr a:ln @w) */
  lineWidthPt?: number
  /** Whether to draw data point markers (line defaults to false; scatter/radar default
   *  comes from the style; only set when <c:marker><c:symbol> is explicit) */
  marker?: boolean
  /** Explicit per-point colors <c:dPt> (common for pies; render layer palette otherwise) */
  pointColors?: Array<string | undefined>
  /** Pie: slice offset from center as percent of diameter (series-level c:explosion → all slices) */
  explosionPct?: number
  /** Pie: per-point explosion overrides (c:dPt/c:explosion) */
  pointExplosionPct?: Array<number | undefined>
  /** Combo dual axes: this series is on the secondary value axis (independent right-side range; decided by the plot's c:axId) */
  secondaryAxis?: boolean
  /** Series belongs to c:stockChart: no connecting lines, participates in whisker/up-down-bar roles */
  fromStock?: boolean
}

export interface ChartAxisStyle {
  /** Explicit min/max (render layer computes a nice range from the data otherwise) */
  min?: number
  max?: number
  labelColor?: string
  labelSizePt?: number
  labelBold?: boolean
  /** Explicit tick-label rotation from txPr bodyPr rot (degrees, e.g. -45) */
  labelRotDeg?: number
  /** <c:delete val="1"/>: axis hidden (no tick labels or axis line) but its scale still applies */
  hidden?: boolean
  /** Tick labels off (<c:tickLblPos val="none"/>): no labels, no reserved space */
  tickLblHidden?: boolean
  /** Labels not rendered but their space still reserved (invalid txPr baseline sentinel) */
  tickLblGarbage?: boolean
  /** <c:numFmt formatCode>: source-linked data labels format numbers with this (omitted for "General") */
  numFmt?: string
  lineColor?: string
  gridColor?: string
  gridDash?: boolean
  /** prstDash value (dash/dashDot/sysDash/...) for a finer dash pattern */
  gridDashVal?: string
  /** Explicit major gridline width (EMU) */
  gridWidthEmu?: number
  /** Minor gridlines: only drawn with an explicit color (PowerPoint hides them by default) */
  minorGridColor?: string
  minorGridWidthEmu?: number
  title?: string
  /** <c:orientation val="maxMin"/>: categories/values reversed (common for bar charts with the first category on top) */
  reversed?: boolean
  /** The category axis is a c:dateAx (PowerPoint rotates its labels as soon as they collide) */
  isDate?: boolean
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
  /** Legend text size (pt) from c:legend/c:txPr default run properties */
  legendPt?: number
  legendBold?: boolean
  /** <c:legend><c:overlay val="1"/>: the legend floats over the plot, reserving no space */
  legendOverlay?: boolean
  /** Plot-area inner rectangle (c:plotArea/c:layout/c:manualLayout layoutTarget=inner), fractions of the chart frame */
  plotLayout?: { x: number; y: number; w: number; h: number }
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
  /** Multi-level category axis: outer-level group labels with their starting leaf index */
  categoryGroups?: Array<{ label: string; start: number }>
  /** Data labels (showVal or showPercent on plot/ser-level c:dLbls) */
  dataLabels?: boolean
  /** Data labels show percentages only (showPercent only, common for pies) */
  dataLabelsPct?: boolean
  /** Data labels include the series name / category name (c:showSerName / c:showCatName) */
  dataLabelSerName?: boolean
  dataLabelCatName?: boolean
  /** Data labels omit the value itself (name-only labels) */
  dataLabelNoValue?: boolean
  /** Data-label number format (c:dLbls/c:numFmt formatCode, else the value axis's; omitted for "General") */
  dataLabelFmt?: string
  /** Data-label text size (pt) / bold from c:dLbls/c:txPr */
  dataLabelPt?: number
  dataLabelBold?: boolean
  /** Chart title (concatenated rich text of c:chart/c:title) */
  title?: string
  /** Title font size (pt) from c:title/c:txPr default run properties */
  titlePt?: number
  /** <c:title><c:overlay val="1"/>: the title floats over the plot, reserving no space */
  titleOverlay?: boolean
  /** chartSpace-level <c:spPr> fill (whole-chart background, e.g. picture fill) */
  bgFill?: Fill
  /** Plot-area fill (c:plotArea's own spPr) */
  plotFill?: Fill
  /** Plot-area border (c:plotArea spPr ln) */
  plotBorder?: { color: string; widthEmu: number }
  /** chartSpace-level <c:spPr><a:ln> frame border around the whole chart */
  border?: { color: string; widthEmu: number }
  /** Theme accent1..6 resolved to #RRGGBB (default series/wedge colors, PowerPoint varyColors order) */
  themePalette?: string[]
  /** Pie: explicit <c:varyColors val="0"/> — every wedge uses the series color */
  varyColors?: boolean
  /** Bubble: <c:bubbleScale> % (default 100), scales the largest bubble's diameter */
  bubbleScale?: number
  /** Bubble: <c:sizeRepresents val="w"/> — size maps to diameter instead of area */
  bubbleSizeIsWidth?: boolean
  /** Stock chart: series are (open,)high,low,close roles in document order */
  stock?: { hiLowLines: boolean; upDownBars: boolean; gapWidthPct?: number }
  /** Sunburst (chartEx): levels leaf-first as stored (levels[0][i] = leaf label of point i,
   *  later levels = ancestors); '' = the point ends at a shallower depth. sizes align by point. */
  sunburst?: {
    levels: string[][]
    sizes: Array<number | null>
    pointColors?: Array<string | undefined>
  }
  /** chartUserShapes straight-line overlays; coordinates are fractions of the chart frame */
  userLines?: Array<{
    x1: number
    y1: number
    x2: number
    y2: number
    color: string
    widthEmu: number
  }>
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
  // Stock (open-)high-low-close rides the line pipeline: category axis + one value series
  // per role; whiskers/up-down bars come from the stock flags, connecting lines stay off
  const stockPlot = plotArea['c:stockChart']
  if (barPlot) cartesian.push({ kind: 'bar', plot: barPlot })
  if (areaPlot) cartesian.push({ kind: 'area', plot: areaPlot })
  if (linePlot) cartesian.push({ kind: 'line', plot: linePlot })
  if (stockPlot) cartesian.push({ kind: 'line', plot: stockPlot })

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
  } else if (plotArea['c:scatterChart'] || plotArea['c:bubbleChart']) {
    // Bubble rides the scatter pipeline: same x/y value model, sized markers via bubbleSizes
    kind = 'scatter'
    plot = plotArea['c:scatterChart'] ?? plotArea['c:bubbleChart']
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
  let categoryGroups: Array<{ label: string; start: number }> | undefined
  const parsePlotSeries = (
    plotNode: any,
    plotKind: ChartKind,
    tagPlotKind: boolean,
    secondary = false,
    fromStock = false,
  ) => {
    const sersRaw = plotNode['c:ser']
    const sers: any[] = Array.isArray(sersRaw) ? sersRaw : sersRaw ? [sersRaw] : []
    for (const ser of sers) {
      // Scatter: y values in c:yVal, x values in c:xVal; other types use c:val
      const s: ChartSeries = {
        values: readNumPoints(plotKind === 'scatter' ? ser['c:yVal'] : ser['c:val']),
      }
      if (tagPlotKind) s.plotKind = plotKind as 'line' | 'bar' | 'area'
      // Excel/PowerPoint pick automatic series colors by c:idx, not document order
      const palIdx = parseInt(ser['c:idx']?.['@_val'], 10)
      if (Number.isFinite(palIdx)) s.paletteIdx = palIdx
      if (secondary) s.secondaryAxis = true
      if (fromStock) s.fromStock = true
      if (plotKind === 'scatter') {
        const xs = readNumPoints(ser['c:xVal'])
        if (xs.length) s.xValues = xs
        const sizes = readNumPoints(ser['c:bubbleSize'])
        if (sizes.length) s.bubbleSizes = sizes
      }
      const name = readStrPoints(ser['c:tx'])[0]
      if (name != null) s.name = name
      // Bubble color is the fill (the line is only the outline), unlike line/scatter/radar
      const color = serColor(
        ser,
        theme,
        !s.bubbleSizes && (plotKind === 'line' || plotKind === 'scatter' || plotKind === 'radar'),
      )
      if (color) s.color = color
      if (ser['c:smooth']?.['@_val'] === '1') s.smooth = true
      const serLn = ser['c:spPr']?.['a:ln']
      const serDash = serLn?.['a:prstDash']?.['@_val']
      if (typeof serDash === 'string' && serDash !== 'solid') s.dash = serDash
      const serLnW = parseInt(serLn?.['@_w'], 10)
      if (Number.isFinite(serLnW) && serLnW > 0) s.lineWidthPt = serLnW / 12700
      // Data-label visibility is per series: a combo chart shows labels only on the
      // series that carry c:dLbls (acsa: bar series labeled, the index line not)
      const dlOn = (d: any) =>
        !!d &&
        typeof d === 'object' &&
        d['c:delete']?.['@_val'] !== '1' &&
        (d['c:showVal']?.['@_val'] === '1' || d['c:showPercent']?.['@_val'] === '1')
      const serDl = ser['c:dLbls']
      s.dataLabels = serDl && typeof serDl === 'object' ? dlOn(serDl) : dlOn(plotNode['c:dLbls'])
      const markerSym = ser['c:marker']?.['c:symbol']?.['@_val']
      if (plotKind === 'line') s.marker = markerSym != null && markerSym !== 'none'
      // scatter/radar: default marker decided by style; only set for explicit symbol (none → false)
      else if ((plotKind === 'scatter' || plotKind === 'radar') && markerSym != null)
        s.marker = markerSym !== 'none'
      const expl = parseInt(ser['c:explosion']?.['@_val'], 10)
      if (Number.isFinite(expl) && expl > 0) s.explosionPct = expl
      // Per-data-point colors (one color per pie slice)
      const dPts: any[] = ser['c:dPt'] ?? []
      if (dPts.length) {
        const pointColors: Array<string | undefined> = []
        const pointExpl: Array<number | undefined> = []
        for (const dPt of dPts) {
          const idx = parseInt(dPt['c:idx']?.['@_val'], 10)
          if (Number.isNaN(idx)) continue
          const c = resolveColorNode(dPt['c:spPr']?.['a:solidFill'], theme)
          if (c != null) pointColors[idx] = c
          const pe = parseInt(dPt['c:explosion']?.['@_val'], 10)
          if (Number.isFinite(pe)) pointExpl[idx] = pe
        }
        if (pointColors.length) s.pointColors = pointColors
        if (pointExpl.length) s.pointExplosionPct = pointExpl
      }
      series.push(s)
      // Categories: take the first non-empty series' cat
      if (!categories.length) categories = readStrPoints(ser['c:cat'])
      // Multi-level category axis: the outer level groups leaf categories (CA | SF, LA)
      if (!categoryGroups) {
        const multi = ser['c:cat']?.['c:multiLvlStrRef']?.['c:multiLvlStrCache']
        const lvlsRaw = multi?.['c:lvl']
        const lvls: any[] = Array.isArray(lvlsRaw) ? lvlsRaw : lvlsRaw ? [lvlsRaw] : []
        if (lvls.length > 1) {
          const ptsRaw = lvls[1]?.['c:pt']
          const pts: any[] = Array.isArray(ptsRaw) ? ptsRaw : ptsRaw ? [ptsRaw] : []
          const groups = pts
            .map((pt) => {
              const v = pt?.['c:v']
              return {
                label: typeof v === 'string' ? v : v != null ? String(v['#text'] ?? '') : '',
                start: parseInt(pt?.['@_idx'], 10) || 0,
              }
            })
            .sort((a, b) => a.start - b.start)
          if (groups.length) categoryGroups = groups
        }
      }
    }
  }
  if (cartesian.length > 1) {
    for (const c of cartesian)
      parsePlotSeries(
        c.plot,
        c.kind,
        true,
        secAxId != null && plotAxIds(c.plot).includes(secAxId),
        c.plot === stockPlot,
      )
  } else parsePlotSeries(plot, kind, false, false, plot === stockPlot)
  if (!series.length) return null
  if (!categories.length) {
    // With no category cache, keep names empty (length from the longest series); never inject placeholders
    const n = Math.max(...series.map((s) => s.values.length), 0)
    categories = Array.from({ length: n }, () => '')
  }

  const model: ChartModel = { kind, categories, series }
  if (categoryGroups) model.categoryGroups = categoryGroups

  if (is3D) {
    model.pseudo3D = true
    const rotX = parseInt(chart['c:view3D']?.['c:rotX']?.['@_val'], 10)
    if (Number.isFinite(rotX)) model.rotXDeg = rotX
  }

  if (kind === 'bar' || kind === 'area') {
    const grouping = plot['c:grouping']?.['@_val']
    if (grouping) model.grouping = grouping
  }
  if (kind === 'bar') {
    const dir = plot['c:barDir']?.['@_val']
    model.barDir = dir === 'bar' ? 'bar' : 'col'
    const gap = plot['c:gapWidth']?.['@_val']
    model.gapWidthPct = gap != null ? parseInt(gap, 10) : 150
  }

  if (kind === 'pie') {
    const hole = plot['c:holeSize']?.['@_val']
    model.holePct = hole != null ? parseInt(hole, 10) || 0 : plotArea['c:doughnutChart'] ? 50 : 0
    const first = plot['c:firstSliceAng']?.['@_val']
    if (first != null) model.firstSliceAngDeg = parseInt(first, 10) || 0
    const vary = plot['c:varyColors']?.['@_val']
    if (vary === '0' || vary === 'false') model.varyColors = false
  }

  if (kind === 'scatter') {
    const st = plot['c:scatterStyle']?.['@_val']
    if (st) model.scatterStyle = String(st)
    const bScale = parseInt(plot['c:bubbleScale']?.['@_val'], 10)
    if (Number.isFinite(bScale) && bScale > 0) model.bubbleScale = bScale
    if (plot['c:sizeRepresents']?.['@_val'] === 'w') model.bubbleSizeIsWidth = true
  }

  if (kind === 'radar') {
    const st = plot['c:radarStyle']?.['@_val']
    model.radarStyle = st === 'filled' ? 'filled' : st === 'marker' ? 'marker' : 'standard'
  }

  if (stockPlot) {
    // Empty elements (<c:hiLowLines/>) parse to "" — presence check must not be truthiness
    const udb = stockPlot['c:upDownBars']
    const gw = parseInt(udb?.['c:gapWidth']?.['@_val'], 10)
    model.stock = {
      hiLowLines: stockPlot['c:hiLowLines'] !== undefined,
      upDownBars: udb !== undefined,
      ...(Number.isFinite(gw) && gw >= 0 ? { gapWidthPct: gw } : {}),
    }
  }

  const legendNode = chart['c:legend']
  const legendPos = legendNode?.['c:legendPos']?.['@_val']
  if (legendNode) {
    model.legendPos = (legendPos as ChartModel['legendPos']) ?? 'r'
    // <c:overlay val="1"/>: the legend floats over the plot, reserving no space
    if (legendNode['c:overlay']?.['@_val'] === '1') model.legendOverlay = true
    const legP = chart['c:legend']?.['c:txPr']?.['a:p']
    const legRPr = (Array.isArray(legP) ? legP[0] : legP)?.['a:pPr']?.['a:defRPr']
    // INT_MIN baseline sentinel: PowerPoint renders no legend at all
    if (legRPr?.['@_baseline'] === '-2147483648') delete model.legendPos
    const legSz = parseInt(legRPr?.['@_sz'], 10)
    if (Number.isFinite(legSz) && legSz > 0) model.legendPt = legSz / 100
    if (legRPr?.['@_b'] === '1') model.legendBold = true
  }

  // Plot-area inner rectangle (edge-mode fractions of the chart frame); PowerPoint positions
  // gridlines/bars exactly here, with axis labels outside it
  const mLay = plotArea['c:layout']?.['c:manualLayout']
  if (mLay?.['c:layoutTarget']?.['@_val'] === 'inner') {
    const frac = (k: string) => Number(mLay[k]?.['@_val'])
    const edgeMode = (k: string) => {
      const m = mLay[k]?.['@_val']
      return m == null || m === 'edge'
    }
    const [lx, ly, lw, lh] = [frac('c:x'), frac('c:y'), frac('c:w'), frac('c:h')]
    if (
      edgeMode('c:xMode') &&
      edgeMode('c:yMode') &&
      [lx, ly, lw, lh].every(Number.isFinite) &&
      lw > 0 &&
      lh > 0
    ) {
      model.plotLayout = { x: lx, y: ly, w: lw, h: lh }
    }
  }

  // Plot-area fill/border (c:plotArea's own spPr)
  const paSpPr = plotArea['c:spPr']
  const paFill =
    resolveFill?.(paSpPr) ??
    (() => {
      const c = resolveColorNode(paSpPr?.['a:solidFill'], theme)
      return c ? ({ type: 'solid', color: c } as Fill) : undefined
    })()
  if (paFill && paFill.type !== 'none') model.plotFill = paFill
  const paLn = paSpPr?.['a:ln']
  const paBorderColor = paLn && !paLn['a:noFill'] && resolveColorNode(paLn['a:solidFill'], theme)
  if (paBorderColor) {
    const w = parseInt(paLn['@_w'], 10)
    model.plotBorder = { color: paBorderColor, widthEmu: Number.isFinite(w) && w > 0 ? w : 9525 }
  }

  // Whole-chart background (chartSpace-level spPr), e.g. picture fill; no spPr →
  // transparent (PowerPoint-verified: overlapping frames show through, aspose Chart2)
  const bgFill = resolveFill?.(doc['c:chartSpace']?.['c:spPr'])
  if (bgFill && bgFill.type !== 'none') model.bgFill = bgFill
  // Whole-chart frame border (chartSpace-level ln)
  const spLn = doc['c:chartSpace']?.['c:spPr']?.['a:ln']
  const borderColor = spLn && !spLn['a:noFill'] && resolveColorNode(spLn['a:solidFill'], theme)
  if (borderColor) {
    const w = parseInt(spLn['@_w'], 10)
    model.border = { color: borderColor, widthEmu: Number.isFinite(w) && w > 0 ? w : 9525 }
  }

  // Theme accents: default series/wedge colors follow the file's theme, not a fixed palette
  if (theme) {
    const accents = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
      .map((k) => theme.colors?.[k])
      .filter((c): c is string => !!c)
    if (accents.length === 6) model.themePalette = accents
  }
  // Legacy 2007 <c:style>: styles arrange in 8 columns; columns 3-8 color every series
  // with a monochrome luminance ramp of accent1-6 (dark→light in plot order). Ramp
  // anchors measured against PowerPoint for Mac (style 12, 4 series).
  const styleVal = parseInt(doc['c:chartSpace']?.['c:style']?.['@_val'], 10)
  const styleCol = Number.isFinite(styleVal) ? (styleVal - 1) % 8 : -1
  if (styleCol >= 2 && model.themePalette && series.length && series.every((s) => !s.color)) {
    const base = model.themePalette[styleCol - 2]
    if (base) {
      const anchors = [0.77, 0.93, 1.44, 1.92]
      const lumAt = (p: number): number => {
        const x = p * (anchors.length - 1)
        const i = Math.min(Math.floor(x), anchors.length - 2)
        return anchors[i]! + (anchors[i + 1]! - anchors[i]!) * (x - i)
      }
      series.forEach((s, i) => {
        s.color = scaleLuminance(base, series.length === 1 ? 1 : lumAt(i / (series.length - 1)))
      })
    }
  }
  // chartSpace-level default text size (hundredths of a pt)
  const txP = doc['c:chartSpace']?.['c:txPr']?.['a:p']
  const defSz = parseInt((Array.isArray(txP) ? txP[0] : txP)?.['a:pPr']?.['a:defRPr']?.['@_sz'], 10)
  if (Number.isFinite(defSz) && defSz > 0) model.defaultTextPt = defSz / 100

  // Title text: rich text, or the cached cell-linked string (c:tx/c:strRef)
  const chartTitle =
    collectText(chart['c:title']?.['c:tx']?.['c:rich']) ||
    readStrPoints(chart['c:title']?.['c:tx'])[0]
  if (chart['c:title']?.['c:overlay']?.['@_val'] === '1') model.titleOverlay = true
  if (chartTitle) model.title = chartTitle
  // Auto title: a <c:title> with no c:tx at all (and autoTitleDeleted != 1) takes the
  // sole series name; with several series PowerPoint shows the literal "Chart Title"
  else if (
    chart['c:title'] &&
    !chart['c:title']?.['c:tx'] &&
    chart['c:autoTitleDeleted']?.['@_val'] !== '1'
  ) {
    model.title = (series.length === 1 && series[0]!.name) || 'Chart Title'
  }
  if (model.title) {
    const titP = chart['c:title']?.['c:txPr']?.['a:p']
    const titleSz = parseInt(
      (Array.isArray(titP) ? titP[0] : titP)?.['a:pPr']?.['a:defRPr']?.['@_sz'],
      10,
    )
    if (Number.isFinite(titleSz) && titleSz > 0) model.titlePt = titleSz / 100
    // Rich-text titles usually carry the size on the first run, not on txPr
    if (!model.titlePt) {
      const rich = chart['c:title']?.['c:tx']?.['c:rich']
      const p0 = Array.isArray(rich?.['a:p']) ? rich['a:p'][0] : rich?.['a:p']
      const r0 = Array.isArray(p0?.['a:r']) ? p0['a:r'][0] : p0?.['a:r']
      const runSz = parseInt(r0?.['a:rPr']?.['@_sz'], 10)
      if (Number.isFinite(runSz) && runSz > 0) model.titlePt = runSz / 100
    }
  }

  // Data labels: plot-level or any series-level c:dLbls (delete=1 counts as none)
  const dLblsInfo = (
    owner: any,
  ): {
    on: boolean
    pct: boolean
    ser?: boolean
    cat?: boolean
    val?: boolean
    fmt?: string
    sizePt?: number
    bold?: boolean
  } => {
    const d = owner?.['c:dLbls']
    if (!d || typeof d !== 'object' || d['c:delete']?.['@_val'] === '1')
      return { on: false, pct: false }
    const showVal = d['c:showVal']?.['@_val'] === '1'
    const showPct = d['c:showPercent']?.['@_val'] === '1'
    const showSer = d['c:showSerName']?.['@_val'] === '1'
    const showCat = d['c:showCatName']?.['@_val'] === '1'
    const fmt = d['c:numFmt']?.['@_formatCode']
    const dP = d['c:txPr']?.['a:p']
    const dRPr = (Array.isArray(dP) ? dP[0] : dP)?.['a:pPr']?.['a:defRPr']
    const dSz = parseInt(dRPr?.['@_sz'], 10)
    return {
      on: showVal || showPct || showSer || showCat,
      pct: showPct && !showVal,
      ser: showSer,
      cat: showCat,
      val: showVal || showPct,
      ...(typeof fmt === 'string' && fmt && fmt !== 'General' ? { fmt } : {}),
      ...(Number.isFinite(dSz) && dSz > 0 ? { sizePt: dSz / 100 } : {}),
      ...(dRPr?.['@_b'] === '1' ? { bold: true } : {}),
    }
  }
  const dLblOwners: any[] = (cartesian.length > 1 ? cartesian.map((c) => c.plot) : [plot]).flatMap(
    (p) => [
      p,
      ...((Array.isArray(p['c:ser']) ? p['c:ser'] : p['c:ser'] ? [p['c:ser']] : []) as any[]),
    ],
  )
  const dLblResults = dLblOwners.map(dLblsInfo)
  const found = dLblResults.find((r) => r.on)
  if (found) {
    model.dataLabels = true
    if (found.pct) model.dataLabelsPct = true
    if (found.ser) model.dataLabelSerName = true
    if (found.cat) model.dataLabelCatName = true
    if (found.on && !found.val) model.dataLabelNoValue = true
    const fmt = found.fmt ?? dLblResults.find((r) => r.fmt)?.fmt
    if (fmt) model.dataLabelFmt = fmt
    const sizePt = found.sizePt ?? dLblResults.find((r) => r.sizePt)?.sizePt
    if (sizePt) model.dataLabelPt = sizePt
    if (found.bold ?? dLblResults.some((r) => r.bold)) model.dataLabelBold = true
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
    // A date axis is a category axis with time semantics — same label/gridline model.
    // Both kinds may coexist (a deleted secondary catAx next to the primary dateAx):
    // pick the first non-deleted candidate across the two lists
    const toArr = (raw: any): any[] => (Array.isArray(raw) ? raw : raw ? [raw] : [])
    const catCands = [
      ...toArr(plotArea['c:catAx']).map((n) => ({ n, date: false })),
      ...toArr(plotArea['c:dateAx']).map((n) => ({ n, date: true })),
    ]
    const pick = catCands.find((c) => c.n?.['c:delete']?.['@_val'] !== '1') ?? catCands[0]
    const catAx = parseAxis(pick?.n, theme)
    if (catAx || pick?.date)
      model.catAxis = { ...(catAx ?? {}), ...(pick?.date ? { isDate: true } : {}) }
  }
  // Source-linked data labels inherit the value axis's number format
  if (model.dataLabels && !model.dataLabelFmt && model.valAxis?.numFmt)
    model.dataLabelFmt = model.valAxis.numFmt

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
  if (numCache) {
    // Date categories: numeric serials + a date formatCode in the cache (PowerPoint
    // displays the formatted date, e.g. 37261 + m/d/yyyy → 1/5/2002)
    const fmt = numCache['c:formatCode']
    const fmtStr = typeof fmt === 'string' ? fmt : String(fmt?.['#text'] ?? '')
    const isDate = /[ymd]/i.test(fmtStr) && !/[#0?]/.test(fmtStr)
    return readPoints(numCache).map((v) => {
      if (v == null) return ''
      if (!isDate) return v
      const serial = parseFloat(v)
      return Number.isFinite(serial) ? formatDateSerial(serial, fmtStr) : v
    })
  }
  return []
}

/** Excel date serial (days since 1899-12-30) formatted per the common date codes. */
function formatDateSerial(serial: number, fmt: string): string {
  const ms = (serial - 25569) * 86400000 // 25569 = days 1899-12-30 → 1970-01-01
  const d = new Date(ms)
  const yyyy = d.getUTCFullYear()
  const mNum = d.getUTCMonth() + 1
  const dayNum = d.getUTCDate()
  const MONTHS = [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec',
  ]
  return fmt
    .replace(/\\/g, '')
    .replace(/yyyy/gi, String(yyyy))
    .replace(/yy/gi, String(yyyy % 100).padStart(2, '0'))
    .replace(/mmm/gi, MONTHS[d.getUTCMonth()]!)
    .replace(/mm/gi, String(mNum).padStart(2, '0'))
    .replace(/(?<![a-z])m(?![a-z])/gi, String(mNum))
    .replace(/dd/gi, String(dayNum).padStart(2, '0'))
    .replace(/(?<![a-z])d(?![a-z])/gi, String(dayNum))
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

/** Series main color: line-family series read the ln stroke first; bars/pies read the
 *  solidFill first (their ln is a segment outline, e.g. black borders around bars). */
function serColor(ser: any, theme: Theme | undefined, preferLine: boolean): string | undefined {
  const spPr = ser['c:spPr']
  if (!spPr) return undefined
  const lnColor = resolveColorNode(spPr['a:ln']?.['a:solidFill'], theme)
  const fillColor = resolveColorNode(spPr['a:solidFill'], theme)
  return preferLine ? (lnColor ?? fillColor) : (fillColor ?? lnColor)
}

function parseAxis(ax: any, theme?: Theme): ChartAxisStyle | undefined {
  if (!ax || typeof ax !== 'object') return undefined
  const out: ChartAxisStyle = {}
  if (ax['c:delete']?.['@_val'] === '1') out.hidden = true
  const scaling = ax['c:scaling']
  if (scaling?.['c:min']?.['@_val'] != null) out.min = Number(scaling['c:min']['@_val'])
  if (scaling?.['c:max']?.['@_val'] != null) out.max = Number(scaling['c:max']['@_val'])
  if (scaling?.['c:orientation']?.['@_val'] === 'maxMin') out.reversed = true
  if (ax['c:tickLblPos']?.['@_val'] === 'none') out.tickLblHidden = true
  const defRPr =
    ax['c:txPr']?.['a:p']?.[0]?.['a:pPr']?.['a:defRPr'] ??
    ax['c:txPr']?.['a:p']?.['a:pPr']?.['a:defRPr']
  if (defRPr) {
    const c = resolveColorNode(defRPr['a:solidFill'], theme)
    if (c) out.labelColor = c
    if (defRPr['@_sz']) out.labelSizePt = parseInt(defRPr['@_sz'], 10) / 100
    if (defRPr['@_b'] === '1') out.labelBold = true
    // INT_MIN baseline sentinel (Aspose-written): PowerPoint reserves the label space
    // but renders nothing there
    if (defRPr['@_baseline'] === '-2147483648') out.tickLblGarbage = true
  }
  // rot outside +-90 deg (e.g. Office's auto marker -60000000) is not an explicit rotation
  const rot = parseInt(ax['c:txPr']?.['a:bodyPr']?.['@_rot'], 10)
  if (Number.isFinite(rot) && rot !== 0 && Math.abs(rot) <= 5400000) out.labelRotDeg = rot / 60000
  const numFmt = ax['c:numFmt']?.['@_formatCode']
  if (typeof numFmt === 'string' && numFmt && numFmt !== 'General') out.numFmt = numFmt
  const lineColor = resolveColorNode(ax['c:spPr']?.['a:ln']?.['a:solidFill'], theme)
  if (lineColor) out.lineColor = lineColor
  // A self-closing <c:majorGridlines/> (no spPr) parses to an empty string; still counts as having gridlines
  const grid = ax['c:majorGridlines']
  if (grid !== undefined) {
    const ln = typeof grid === 'object' ? grid['c:spPr']?.['a:ln'] : undefined
    const gc = resolveColorNode(ln?.['a:solidFill'], theme)
    out.gridColor = gc ?? '#E6E6E6'
    const dash = ln?.['a:prstDash']?.['@_val']
    if (dash && dash !== 'solid') {
      out.gridDash = true
      out.gridDashVal = String(dash)
    }
    const w = parseInt(ln?.['@_w'], 10)
    if (Number.isFinite(w) && w > 0) out.gridWidthEmu = w
  }
  const minor = ax['c:minorGridlines']
  if (minor !== undefined && typeof minor === 'object') {
    const ln = minor['c:spPr']?.['a:ln']
    const gc = resolveColorNode(ln?.['a:solidFill'], theme)
    if (gc) {
      out.minorGridColor = gc
      const w = parseInt(ln?.['@_w'], 10)
      if (Number.isFinite(w) && w > 0) out.minorGridWidthEmu = w
    }
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
