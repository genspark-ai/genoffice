/**
 * Chart layouter — ChartModel (pptx-engine parse) → ChartRenderNode (draw primitives).
 *
 * Same idea as text layout: all geometry (ticks, bar widths, polyline points,
 * legend, wedge angles) is computed to px here; Konva only draws. Supports
 * line / area / bar (columns + horizontal bars: clustered + stacked + percent
 * stacked) / pie (incl. doughnut) / scatter / radar, and bar+line/area combos
 * (series.plotKind). Combos support a secondary value axis (series.secondaryAxis:
 * independent right-side range + tick labels). Unrecognized types fall back to a
 * placeholder chip upstream.
 */
import type { ChartModel } from '@genoffice/pptx-engine'
import type { ChartRenderNode } from './render-tree'
import type { PlacedBox } from './coords'
import { emuToPx, ptToPx, type Viewport } from './coords'
import type { FontMetricsProvider, RunStyle } from './metrics'

/** Default series palette (approximation of PowerPoint's default theme accent sequence). */
const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']

const LABEL_FONT = 'Arial'

/** Series/wedge default colors: the file theme's accent1..6 when available, else the fixed approximation. */
function chartPalette(model: ChartModel): string[] {
  return model.themePalette?.length ? model.themePalette : PALETTE
}

/** Chart body text size (pt): chartSpace-level c:txPr default, else 10pt. */
function chartTextPt(model: ChartModel): number {
  return model.defaultTextPt ?? 10
}

/** Multiply an #RRGGBB color's channels (pseudo-3D face shading); non-hex passes through. */
function shade(color: string, f: number): string {
  const m = /^#([0-9a-f]{6})$/i.exec(color)
  if (!m) return color
  const ch = (i: number) =>
    Math.min(255, Math.round(parseInt(m[1]!.slice(i, i + 2), 16) * f))
      .toString(16)
      .padStart(2, '0')
  return `#${ch(0)}${ch(2)}${ch(4)}`
}

/** Pseudo-3D extrusion faces for one bar (top + right, painter-ready before the front rect). */
function barFaces(
  x: number,
  y: number,
  w: number,
  h: number,
  d: number,
  color: string,
): Array<{ d: string; fill: string }> {
  const top = `M ${x} ${y} L ${x + w} ${y} L ${x + w + d} ${y - d} L ${x + d} ${y - d} Z`
  const side = `M ${x + w} ${y} L ${x + w + d} ${y - d} L ${x + w + d} ${y + h - d} L ${x + w} ${y + h} Z`
  return [
    { d: top, fill: shade(color, 1.18) },
    { d: side, fill: shade(color, 0.78) },
  ]
}

/** prstDash → Konva dash array, scaled by line width (PowerPoint-like proportions). */
function dashArray(val: string | undefined, widthPx: number): number[] {
  const u = Math.max(widthPx, 1)
  switch (val) {
    case 'dot':
      return [u, 3 * u]
    case 'sysDot':
      return [u, u]
    case 'sysDash':
      return [3 * u, u]
    case 'lgDash':
      return [8 * u, 3 * u]
    case 'dashDot':
      return [4 * u, 3 * u, u, 3 * u]
    case 'lgDashDot':
      return [8 * u, 3 * u, u, 3 * u]
    case 'lgDashDotDot':
      return [8 * u, 3 * u, u, 3 * u, u, 3 * u]
    case 'sysDashDot':
      return [3 * u, u, u, u]
    case 'sysDashDotDot':
      return [3 * u, u, u, u, u, u]
    default:
      return [4 * u, 3 * u]
  }
}

export function buildChartNode(
  id: string,
  sourceId: string,
  model: ChartModel,
  box: PlacedBox,
  vp: Viewport,
  metrics: FontMetricsProvider,
): ChartRenderNode | null {
  if (!model.title) {
    const node = buildChartNodeInner(id, sourceId, model, box, vp, metrics)
    if (node) extrudeBars(node, model)
    return node
  }
  // With a title: the content area yields a top row; shift everything down and draw the title
  // centered on top. Explicit c:title size wins (txPr default, else the first rich run);
  // otherwise PowerPoint sizes the auto title at 1.2× the chart's default text.
  const titleSizePx = ptToPx(model.titlePt ?? chartTextPt(model) * 1.2, vp.scale)
  const measureTitle = (t: string) =>
    metrics.measure(t, {
      fontFamily: LABEL_FONT,
      fontSizePx: titleSizePx,
      bold: true,
      italic: false,
    })
  const titleLines = wrapToWidth(model.title, Math.max(box.w - 16, 40), measureTitle)
  const titleH = titleSizePx * 1.4 * titleLines.length + titleSizePx * 0.3
  // A manual plot layout already positions the plot in full-frame fractions (title space
  // included), so the content must not shrink or shift — but only the cartesian builder
  // consumes plotLayout; pie/scatter/radar/hbar still lay out from the (shrunk) frame
  const manual =
    (!!model.plotLayout &&
      (model.kind === 'line' ||
        model.kind === 'area' ||
        (model.kind === 'bar' && model.barDir !== 'bar'))) ||
    // <c:overlay val="1"/>: the title floats over the plot without reserving space
    !!model.titleOverlay
  const node = buildChartNodeInner(
    id,
    sourceId,
    model,
    manual ? box : { ...box, h: Math.max(box.h - titleH, 10) },
    vp,
    metrics,
  )
  if (!node) return null
  if (!manual) shiftChartNode(node, titleH)
  extrudeBars(node, model)
  node.box = box
  titleLines.forEach((line, i) => {
    node.labels.push({
      text: line,
      x: Math.max((box.w - measureTitle(line)) / 2, 4),
      y: titleSizePx * 0.3 + i * titleSizePx * 1.4,
      fontSizePx: titleSizePx,
      color: '#333333',
      bold: true,
    })
  })
  return node
}

/** Greedy word-wrap to a pixel width; a single overlong word stays on its own line. */
function wrapToWidth(text: string, maxW: number, measure: (t: string) => number): string[] {
  if (measure(text) <= maxW) return [text]
  const words = text.split(/\s+/).filter(Boolean)
  const lines: string[] = []
  let cur = ''
  for (const w of words) {
    const cand = cur ? `${cur} ${w}` : w
    if (cur && measure(cand) > maxW) {
      lines.push(cur)
      cur = w
    } else {
      cur = cand
    }
  }
  if (cur) lines.push(cur)
  return lines
}

/** Pseudo-3D bars: extrusion faces (top + right) behind every bar rect. */
function extrudeBars(node: ChartRenderNode, model: ChartModel): void {
  if (!model.pseudo3D || !node.bars.length) return
  const faces: NonNullable<ChartRenderNode['paths']> = []
  for (const b of node.bars) {
    const d = Math.min(Math.max(Math.min(b.w, b.h) * 0.35, 3), 14)
    faces.push(...barFaces(b.x, b.y, b.w, b.h, d, b.color))
  }
  node.paths = [...faces, ...(node.paths ?? [])]
}

/** Shifts all chart draw primitives vertically (moving the content area down after the title claims space). */
function shiftChartNode(node: ChartRenderNode, dy: number): void {
  for (const g of node.gridLines) {
    g.y1 += dy
    g.y2 += dy
  }
  for (const a of node.axisLines) {
    a.y1 += dy
    a.y2 += dy
  }
  for (const l of node.labels) l.y += dy
  for (const b of node.bars) b.y += dy
  for (const p of node.polylines)
    for (let i = 1; i < p.points.length; i += 2) p.points[i] = p.points[i]! + dy
  for (const m of node.markers) m.y += dy
  for (const s of node.swatches) s.y += dy
  for (const w of node.wedges ?? []) w.cy += dy
  for (const p of node.paths ?? []) p.dy = (p.dy ?? 0) + dy
  if (node.plotRect) node.plotRect.y += dy
}

function buildChartNodeInner(
  id: string,
  sourceId: string,
  model: ChartModel,
  box: PlacedBox,
  vp: Viewport,
  metrics: FontMetricsProvider,
): ChartRenderNode | null {
  if (model.kind === 'pie') return buildPieNode(id, sourceId, model, box, vp, metrics)
  if (model.kind === 'scatter') return buildScatterNode(id, sourceId, model, box, vp, metrics)
  if (model.kind === 'radar') return buildRadarNode(id, sourceId, model, box, vp, metrics)
  if (model.kind === 'funnel') return buildFunnelNode(id, sourceId, model, box, vp, metrics)
  if (model.kind === 'sunburst') return buildSunburstNode(id, sourceId, model, box, vp, metrics)
  if (model.kind === 'bar' && model.barDir === 'bar') {
    return buildHBarNode(id, sourceId, model, box, vp, metrics)
  }
  if (model.kind !== 'line' && model.kind !== 'bar' && model.kind !== 'area') return null

  const grouping =
    model.kind === 'bar'
      ? (model.grouping ?? 'clustered')
      : model.kind === 'area'
        ? (model.grouping ?? 'standard')
        : 'standard'
  const stacked = grouping === 'stacked' || grouping === 'percentStacked'
  // Combo charts: each series dispatches by its own plotKind (bars draw bars, line/area draw lines), sharing the category axis
  const serKind = (ser: (typeof model.series)[number]) => ser.plotKind ?? model.kind
  const barSeriesIdx = model.series
    .map((s, i) => i)
    .filter((i) => serKind(model.series[i]!) === 'bar')
  // Dual axis: line/area series on the secondary axis use an independent right-side range (bar series always use the primary axis so stacking semantics hold)
  const numVals = (s: (typeof model.series)[number]) =>
    s.values.filter((v): v is number => v != null)
  let secVals = model.series.filter((s) => s.secondaryAxis && serKind(s) !== 'bar').flatMap(numVals)
  let priVals = model.series
    .filter((s) => !(s.secondaryAxis && serKind(s) !== 'bar'))
    .flatMap(numVals)
  // No data on the primary axis (all series on the secondary) → degrade to a single axis to avoid an empty range
  if (!priVals.length) {
    priVals = secVals
    secVals = []
  }
  const onSecAxis = (ser: (typeof model.series)[number]) =>
    secVals.length > 0 && !!ser.secondaryAxis && serKind(ser) !== 'bar'

  const node: ChartRenderNode = {
    id,
    type: 'chart',
    box,
    sourceId,
    gridLines: [],
    axisLines: [],
    labels: [],
    bars: [],
    polylines: [],
    markers: [],
    swatches: [],
  }

  const labelSizePx = ptToPx(model.valAxis?.labelSizePt ?? chartTextPt(model), vp.scale)
  const labelColor = model.valAxis?.labelColor ?? '#666666'
  const catLabelSizePx = ptToPx(
    model.catAxis?.labelSizePt ?? model.valAxis?.labelSizePt ?? chartTextPt(model),
    vp.scale,
  )
  const catLabelColor = model.catAxis?.labelColor ?? labelColor
  const style = (sizePx: number): RunStyle => ({
    fontFamily: LABEL_FONT,
    fontSizePx: sizePx,
    bold: false,
    italic: false,
  })
  const measure = (text: string, sizePx: number) => metrics.measure(text, style(sizePx))

  const palette = chartPalette(model)
  const seriesColor = (i: number) =>
    model.series[i]?.color ?? palette[(model.series[i]?.paletteIdx ?? i) % palette.length]!

  // ── Value range + nice ticks (primary/secondary axes independent) ──
  if (!priVals.length) return null
  const catCount = Math.max(model.categories.length, ...model.series.map((s) => s.values.length), 1)
  // Percent stacked: normalize each value to its category share (%) — stacking only applies to bar series
  const catAbsTotals = Array.from({ length: catCount }, (_, i) =>
    barSeriesIdx.reduce((a, si) => a + Math.abs(model.series[si]!.values[i] ?? 0), 0),
  )
  const valueAt = (si: number, i: number): number | null => {
    const v = model.series[si]?.values[i]
    if (v == null) return null
    if (grouping !== 'percentStacked' || serKind(model.series[si]!) !== 'bar') return v
    const total = catAbsTotals[i] || 1
    return (v / total) * 100
  }
  // Stacked: bar series ranges use per-category positive/negative stack sums; overlaid line/area series use raw values
  let dataMax: number
  let dataMin: number
  if (stacked) {
    // Bars and areas both stack; overlay lines stay raw
    const stackIdx = model.series
      .map((s, i) => (serKind(s) === 'bar' || serKind(s) === 'area' ? i : -1))
      .filter((i) => i >= 0)
    const posSums = Array.from({ length: catCount }, (_, i) =>
      stackIdx.reduce((a, si) => a + Math.max(valueAt(si, i) ?? 0, 0), 0),
    )
    const negSums = Array.from({ length: catCount }, (_, i) =>
      stackIdx.reduce((a, si) => a + Math.min(valueAt(si, i) ?? 0, 0), 0),
    )
    // Overlaid line series on the secondary axis don't feed into the primary range
    const overlayVals = model.series
      .filter((s) => serKind(s) !== 'bar' && serKind(s) !== 'area' && !onSecAxis(s))
      .flatMap(numVals)
    dataMax = Math.max(...posSums, ...overlayVals, 0)
    dataMin = Math.min(...negSums, ...overlayVals, 0)
  } else {
    dataMax = Math.max(...priVals, 0)
    dataMin = Math.min(...priVals, 0)
  }
  // Percent stacked: the value axis is exactly 0-100% (−100% with negative stacks)
  if (grouping === 'percentStacked') {
    dataMax = 100
    dataMin = dataMin < 0 ? -100 : 0
  }
  // PowerPoint spaces vertical ticks by the physical axis length (>= ~1.75 label
  // line-heights per interval, python-pptx sheets measured: 385px/32px charts give
  // range 13 -> unit 5 and 33.6 -> 10, while a 569px chart keeps range 4.6 -> 0.5).
  // Large charts stay on the 5-10 interval ratio rule (poi/aspose calibration
  // untouched - the cap only ever makes ticks sparser)
  const maxIntervals = Math.max(3, Math.min(10, Math.floor(box.h / (labelSizePx * 1.75))))
  const { min, max, ticks } = ppTicks(
    model.valAxis?.min ?? dataMin,
    model.valAxis?.max ?? dataMax,
    model.valAxis?.min == null && grouping !== 'percentStacked',
    model.valAxis?.max == null && grouping !== 'percentStacked',
    false,
    maxIntervals,
  )
  // Secondary value axis ticks (right side): range fully independent of the primary (e.g. left axis revenue 0-250, right axis growth% 0-30)
  const sec = secVals.length
    ? ppTicks(
        model.valAxis2?.min ?? Math.min(...secVals, 0),
        model.valAxis2?.max ?? Math.max(...secVals, 0),
        model.valAxis2?.min == null,
        model.valAxis2?.max == null,
        false,
        maxIntervals,
      )
    : undefined

  // ── Layout frame ────────────────────────────────────────────────
  const pad = Math.max(4, box.w * 0.01)
  const legendPos = model.legendPos
  const legendSizePx = ptToPx(
    model.legendPt ?? model.valAxis?.labelSizePt ?? chartTextPt(model),
    vp.scale,
  )
  const legendBold = !!model.legendBold
  const legendH =
    (legendPos === 't' || legendPos === 'b') && !model.legendOverlay ? legendSizePx * 1.6 : 0
  const legendItems = model.series.some((s) => s.name) ? model.series.map((s) => s.name ?? '') : []
  // Side legends claim their true width (swatch + text at the legend's own size), so
  // the plot area shrinks instead of the legend clipping at the frame edge
  const legendW =
    legendItems.length &&
    !model.legendOverlay &&
    (legendPos === 'r' || legendPos === 'l' || legendPos === 'tr')
      ? legendSizePx * 0.5 +
        4 +
        Math.max(
          ...legendItems.map((t) =>
            metrics.measure(t, {
              fontFamily: LABEL_FONT,
              fontSizePx: legendSizePx,
              bold: legendBold,
              italic: false,
            }),
          ),
          0,
        ) +
        8
      : 0
  const valHidden = !!model.valAxis?.hidden
  // Tick labels can be off two ways: none = no labels and no reserved space;
  // garbage txPr baseline = space reserved, nothing drawn
  const valLabelsOff =
    valHidden || !!model.valAxis?.tickLblHidden || !!model.valAxis?.tickLblGarbage
  const catLabelsOff =
    !!model.catAxis?.hidden || !!model.catAxis?.tickLblHidden || !!model.catAxis?.tickLblGarbage
  const valNoReserve = valHidden || !!model.valAxis?.tickLblHidden
  const catNoReserve = !!model.catAxis?.hidden || !!model.catAxis?.tickLblHidden
  const tickLabels = ticks.map((t) => (grouping === 'percentStacked' ? `${fmtNum(t)}%` : fmtNum(t)))
  const yLabelW = valNoReserve ? 0 : Math.max(...tickLabels.map((t) => measure(t, labelSizePx)), 0)
  // Vertical axis-title placeholder ≈2.2 line heights (left margin measured from reference renderings)
  const axisTitleW = model.valAxis?.title ? labelSizePx * 2.2 : 0
  // Secondary axis (right): tick label width + optional axis-title placeholder
  const secLabelSizePx = ptToPx(
    model.valAxis2?.labelSizePt ?? model.valAxis?.labelSizePt ?? chartTextPt(model),
    vp.scale,
  )
  const secTickLabels = sec ? sec.ticks.map((t) => fmtNum(t)) : []
  const y2LabelW = sec ? Math.max(...secTickLabels.map((t) => measure(t, secLabelSizePx)), 0) : 0
  const axisTitle2W = sec && model.valAxis2?.title ? secLabelSizePx * 2.2 : 0

  const plotX = pad + axisTitleW + yLabelW + 10
  const plotY = pad + (legendPos === 't' ? legendH + 4 : 0) + labelSizePx * 0.6
  // Right side: no secondary axis → small gap (PowerPoint measured ≈1.5% width); with one, reserve space for tick labels + title
  const plotR = box.w - pad - (sec ? y2LabelW + axisTitle2W + 10 : labelSizePx * 0.7) - legendW
  const nCats = Math.max(model.categories.length, 1)
  const maxCatW = Math.max(...model.categories.map((c) => measure(c, catLabelSizePx)), 1)
  // Category label mode: an explicit txPr rotation wins; otherwise crowded labels wrap
  // to ≤3 horizontal lines when the words fit the slot, and rotate ~45° as a last resort
  // (PowerPoint thins labels only after that)
  const heuristicSlotW = Math.max(plotR - plotX, 10) / nCats
  const catRotDeg = model.catAxis?.labelRotDeg
  const catWrapLines = (slotW: number): string[][] | null => {
    if (model.categories.length < 2 || maxCatW <= slotW * 1.05) return null
    const maxLineW = slotW * 0.95
    const wrapped = model.categories.map((c) =>
      wrapToWidth(c, maxLineW, (t) => measure(t, catLabelSizePx)),
    )
    const fits = wrapped.every(
      (ls) => ls.length <= 3 && ls.every((l) => measure(l, catLabelSizePx) <= slotW * 1.05),
    )
    return fits && wrapped.some((ls) => ls.length > 1) ? wrapped : null
  }
  const heuristicWrap = catRotDeg == null ? catWrapLines(heuristicSlotW) : null
  // Date-axis labels rotate as soon as they collide; plain category labels tolerate
  // mild overflow before rotating (both PowerPoint-observed)
  const rotFactor = model.catAxis?.isDate ? 1 : 1.3
  const rotateCats =
    (catRotDeg != null && catRotDeg !== 0) ||
    (catRotDeg == null &&
      !heuristicWrap &&
      model.categories.length > 1 &&
      maxCatW > heuristicSlotW * rotFactor)
  const catRotUsed = catRotDeg != null && catRotDeg !== 0 ? catRotDeg : -45
  const rotReserve = Math.abs(Math.sin((catRotUsed * Math.PI) / 180))
  const catReserve = rotateCats
    ? Math.min(maxCatW * rotReserve, box.h * 0.3) + catLabelSizePx * 0.6
    : catLabelSizePx * 1.2 * (heuristicWrap ? Math.max(...heuristicWrap.map((l) => l.length)) : 1) +
      catLabelSizePx * 0.3
  const catTitleH = model.catAxis?.title ? catLabelSizePx * 1.6 : 0
  const catGroupH = model.categoryGroups && !catNoReserve ? catLabelSizePx * 1.5 : 0
  const plotB =
    box.h -
    pad -
    (catNoReserve ? 0 : catReserve) -
    catGroupH -
    catTitleH -
    (legendPos === 'b' ? legendH : 0)
  const L = model.plotLayout
  const plot = L
    ? {
        x: L.x * box.w,
        y: L.y * box.h,
        w: Math.max(L.w * box.w, 10),
        h: Math.max(L.h * box.h, 10),
      }
    : {
        x: plotX,
        y: plotY,
        w: Math.max(plotR - plotX, 10),
        h: Math.max(plotB - plotY, 10),
      }

  // 3D bar charts (default right-angle axes): PowerPoint draws bars on a front plane
  // and gridlines/ticks on a back wall shifted (+d, -d) from it — both walls the same
  // size, the front floor flush with the stage bottom. d measured against PowerPoint
  // (workarea deck, bar3D 21 cats): ~0.8x one bar width. `plot` becomes the front plane.
  let depth3d = 0
  if (model.pseudo3D && barSeriesIdx.length && model.kind === 'bar') {
    // Stacked draws one full-slot bar per category, so its width divisor is 1, not the series count
    const barW0 =
      plot.w /
      Math.max(model.categories.length, 1) /
      ((stacked ? 1 : Math.max(barSeriesIdx.length, 1)) + (model.gapWidthPct ?? 150) / 100)
    depth3d = Math.min(barW0 * 0.8, plot.w * 0.06, plot.h * 0.15)
    // Bars sit at the mid-depth of the 3D box: one full d right of the stage left edge
    plot.x += depth3d
    plot.y += depth3d
    plot.w -= depth3d * 3
    plot.h -= depth3d
  }

  const yOf = (v: number) => plot.y + plot.h * (1 - (v - min) / (max - min || 1))
  // Secondary axis mapping: same plot-area height, independent range
  const yOf2 = (v: number) =>
    plot.y + plot.h * (1 - (v - (sec?.min ?? 0)) / ((sec?.max ?? 1) - (sec?.min ?? 0) || 1))

  // ── Plot-area fill/border + gridlines + tick labels + axes (back wall for 3D: +d, -d) ──
  if (model.plotFill || model.plotBorder) {
    node.plotRect = {
      x: plot.x + depth3d * 0.5,
      y: plot.y - depth3d,
      w: plot.w,
      h: plot.h,
      ...(model.plotBorder
        ? {
            borderColor: model.plotBorder.color,
            borderWidthPx: Math.max(emuToPx(model.plotBorder.widthEmu, vp.scale), 0.75),
          }
        : {}),
    }
  }
  const gridColor = model.valAxis?.gridColor
  const gridW = model.valAxis?.gridWidthEmu
    ? Math.max(emuToPx(model.valAxis.gridWidthEmu, vp.scale), 0.75)
    : undefined
  // Minor gridlines only with an explicit color; PowerPoint subdivides each major gap in 5
  const minorColor = model.valAxis?.minorGridColor
  const minorW = model.valAxis?.minorGridWidthEmu
    ? Math.max(emuToPx(model.valAxis.minorGridWidthEmu, vp.scale), 0.75)
    : undefined
  const gxa = plot.x + depth3d * 0.5
  for (let i = 0; i < ticks.length; i++) {
    const t = ticks[i]!
    const y = yOf(t) - depth3d
    if (gridColor && t !== min) {
      node.gridLines.push({
        x1: gxa,
        y1: y,
        x2: gxa + plot.w,
        y2: y,
        color: gridColor,
        ...(model.valAxis?.gridDash
          ? { dash: dashArray(model.valAxis?.gridDashVal, gridW ?? 1) }
          : {}),
        ...(gridW ? { widthPx: gridW } : {}),
      })
    }
    if (minorColor && i + 1 < ticks.length) {
      const gap = (yOf(ticks[i + 1]!) - yOf(t)) / 5
      for (let k = 1; k < 5; k++) {
        const my = yOf(t) + gap * k - depth3d
        node.gridLines.push({
          x1: gxa,
          y1: my,
          x2: gxa + plot.w,
          y2: my,
          color: minorColor,
          ...(minorW ? { widthPx: minorW } : {}),
        })
      }
    }
    if (valLabelsOff) continue
    const text = tickLabels[i]!
    node.labels.push({
      text,
      x: plot.x + depth3d * 0.5 - 6 - measure(text, labelSizePx),
      y: y - labelSizePx * 0.55,
      fontSizePx: labelSizePx,
      color: labelColor,
      ...(model.valAxis?.labelBold ? { bold: true } : {}),
    })
  }
  const axisColor = model.valAxis?.lineColor ?? model.catAxis?.lineColor ?? '#888888'
  const axisW = Math.max(1, ptToPx(1, vp.scale))
  // x axis (bottom, at the front floor edge) + y axis (left, on the back wall; skipped when deleted)
  node.axisLines.push({
    x1: plot.x,
    y1: plot.y + plot.h,
    x2: plot.x + plot.w,
    y2: plot.y + plot.h,
    color: axisColor,
    widthPx: axisW,
  })
  if (!valHidden)
    node.axisLines.push({
      x1: plot.x + depth3d * 0.5,
      y1: plot.y - depth3d,
      x2: plot.x + depth3d * 0.5,
      y2: plot.y - depth3d + plot.h,
      color: axisColor,
      widthPx: axisW,
    })

  // Secondary value axis (right): axis line + independent tick labels; gridlines are drawn for the primary only, so two grids don't fight
  if (sec) {
    const secLabelColor = model.valAxis2?.labelColor ?? labelColor
    node.axisLines.push({
      x1: plot.x + plot.w,
      y1: plot.y,
      x2: plot.x + plot.w,
      y2: plot.y + plot.h,
      color: model.valAxis2?.lineColor ?? axisColor,
      widthPx: axisW,
    })
    sec.ticks.forEach((t, i) => {
      node.labels.push({
        text: secTickLabels[i]!,
        x: plot.x + plot.w + 6,
        y: yOf2(t) - secLabelSizePx * 0.55,
        fontSizePx: secLabelSizePx,
        color: secLabelColor,
      })
    })
    // Secondary axis title (vertical on the right, 90° clockwise)
    if (model.valAxis2?.title) {
      const tw = measure(model.valAxis2.title, secLabelSizePx)
      node.labels.push({
        text: model.valAxis2.title,
        x: box.w - pad,
        y: plot.y + plot.h / 2 - tw / 2,
        fontSizePx: secLabelSizePx,
        color: secLabelColor,
        rotationDeg: 90,
      })
    }
  }

  // Category-axis title (horizontal, centered under the category labels)
  if (model.catAxis?.title) {
    const tw = measure(model.catAxis.title, catLabelSizePx)
    node.labels.push({
      text: model.catAxis.title,
      x: plot.x + (plot.w - tw) / 2,
      // Sits above a bottom legend (both were subtracted from the plot height)
      y: box.h - pad - (legendPos === 'b' ? legendH : 0) - catLabelSizePx * 1.2,
      fontSizePx: catLabelSizePx,
      color: model.catAxis.labelColor ?? labelColor,
    })
  }
  // Value-axis title (vertical, 90° counterclockwise; Konva rotates around the top-left)
  if (model.valAxis?.title) {
    const tw = measure(model.valAxis.title, labelSizePx)
    node.labels.push({
      text: model.valAxis.title,
      x: pad,
      y: plot.y + plot.h / 2 + tw / 2,
      fontSizePx: labelSizePx,
      color: labelColor,
      rotationDeg: -90,
    })
  }

  // ── Category labels (all shown: horizontal, wrapped horizontal, or rotated when crowded) ──
  const n = Math.max(model.categories.length, 1)
  const slotW = plot.w / n
  // Category-axis vertical gridlines: majors at slot boundaries, minors at slot midpoints
  const catGrid = model.catAxis?.gridColor
  const catMinor = model.catAxis?.minorGridColor
  if (catGrid || catMinor) {
    const catW = model.catAxis?.gridWidthEmu
      ? Math.max(emuToPx(model.catAxis.gridWidthEmu, vp.scale), 0.75)
      : undefined
    const catMinorW = model.catAxis?.minorGridWidthEmu
      ? Math.max(emuToPx(model.catAxis.minorGridWidthEmu, vp.scale), 0.75)
      : undefined
    const gy = plot.y - depth3d
    for (let i = 0; i <= n; i++) {
      const gx = plot.x + depth3d * 0.5 + i * slotW
      if (catGrid && i > 0) {
        node.gridLines.push({
          x1: gx,
          y1: gy,
          x2: gx,
          y2: gy + plot.h,
          color: catGrid,
          ...(model.catAxis?.gridDash
            ? { dash: dashArray(model.catAxis?.gridDashVal, catW ?? 1) }
            : {}),
          ...(catW ? { widthPx: catW } : {}),
        })
      }
      if (catMinor && i < n) {
        node.gridLines.push({
          x1: gx + slotW / 2,
          y1: gy,
          x2: gx + slotW / 2,
          y2: gy + plot.h,
          color: catMinor,
          ...(catMinorW ? { widthPx: catMinorW } : {}),
        })
      }
    }
  }
  const catBold = model.catAxis?.labelBold ? { bold: true as const } : {}
  // Re-derive the mode from the final plot width (a manual plot layout can differ from the heuristic frame)
  const drawWrap = catRotDeg == null ? catWrapLines(slotW) : null
  const drawRotate =
    (catRotDeg != null && catRotDeg !== 0) ||
    (catRotDeg == null && !drawWrap && model.categories.length > 1 && maxCatW > slotW * rotFactor)
  const rotCos = Math.cos((Math.abs(catRotUsed) * Math.PI) / 180)
  const rotSin = Math.sin((Math.abs(catRotUsed) * Math.PI) / 180)
  if (!catLabelsOff)
    model.categories.forEach((cat, i) => {
      const cx = plot.x + (i + 0.5) * slotW
      if (drawRotate) {
        // Rotation is about the text's origin (left end). Negative angles slant up-right:
        // place the origin down-left so the right end lands just below the tick. Positive
        // angles slant down-right: the left end itself sits at the tick.
        const w = measure(cat, catLabelSizePx)
        const neg = catRotUsed < 0
        node.labels.push({
          text: cat,
          x: neg ? cx - w * rotCos - catLabelSizePx * 0.5 : cx - catLabelSizePx * 0.5,
          y: plot.y + plot.h + catLabelSizePx * 0.5 + (neg ? w * rotSin : 0),
          fontSizePx: catLabelSizePx,
          color: catLabelColor,
          rotationDeg: catRotUsed,
          ...catBold,
        })
        return
      }
      const lines = drawWrap?.[i] ?? [cat]
      lines.forEach((line, li) => {
        node.labels.push({
          text: line,
          x: cx - measure(line, catLabelSizePx) / 2,
          y: plot.y + plot.h + catLabelSizePx * 0.35 + li * catLabelSizePx * 1.2,
          fontSizePx: catLabelSizePx,
          color: catLabelColor,
          ...catBold,
        })
      })
    })

  // Multi-level category axis: outer-level group labels centered under their span,
  // with divider ticks between groups (Excel's boxed second row, simplified)
  if (model.categoryGroups && !catLabelsOff) {
    const gy = plot.y + plot.h + catReserve + catLabelSizePx * 0.35
    const nCatsG = Math.max(model.categories.length, 1)
    const slotWG = plot.w / nCatsG
    model.categoryGroups.forEach((g, gi) => {
      const end = model.categoryGroups![gi + 1]?.start ?? nCatsG
      const cxG = plot.x + ((g.start + end) / 2) * slotWG
      node.labels.push({
        text: g.label,
        x: cxG - measure(g.label, catLabelSizePx) / 2,
        y: gy,
        fontSizePx: catLabelSizePx,
        color: catLabelColor,
      })
      if (g.start > 0)
        node.axisLines.push({
          x1: plot.x + g.start * slotWG,
          y1: plot.y + plot.h,
          x2: plot.x + g.start * slotWG,
          y2: plot.y + plot.h + catReserve + catLabelSizePx * 1.5,
          color: model.catAxis?.lineColor ?? '#888888',
          widthPx: 1,
        })
    })
  }

  // ── Data series (bar series + line/area series can coexist: combo charts) ──
  const dlSize = model.dataLabelPt ? ptToPx(model.dataLabelPt, vp.scale) : labelSizePx * 0.9
  const dlBold = !!model.dataLabelBold
  // Data labels (c:dLbls showVal): white centered inside stacked segments, dark gray outside points/bars otherwise
  const dLbl = (si: number, catIdx: number, cx: number, y: number, v: number, inside: boolean) => {
    if (!(model.series[si]?.dataLabels ?? model.dataLabels)) return
    const text = composeDataLabel(model, si, catIdx, fmtDataLabel(v, model.dataLabelFmt))
    node.labels.push({
      text,
      x: cx - measure(text, dlSize) / 2,
      y,
      fontSizePx: dlSize,
      color: inside ? '#FFFFFF' : dlBold ? '#000000' : '#404040',
      ...(dlBold ? { bold: true } : {}),
    })
  }
  if (barSeriesIdx.length && stacked) {
    // Stacked columns: one bar per category, series accumulate along the value axis (positive up, negative down)
    const gap = (model.gapWidthPct ?? 150) / 100
    const barW = slotW / (1 + gap)
    for (let i = 0; i < n; i++) {
      const x = plot.x + i * slotW + (slotW - barW) / 2
      let posAcc = 0
      let negAcc = 0
      barSeriesIdx.forEach((si) => {
        const ser = model.series[si]!
        const v = valueAt(si, i)
        if (v == null || v === 0) return
        const color = ser.pointColors?.[i] ?? seriesColor(si)!
        const from = v > 0 ? posAcc : negAcc
        const to = from + v
        if (v > 0) posAcc = to
        else negAcc = to
        const yTop = yOf(Math.max(from, to))
        const yBot = yOf(Math.min(from, to))
        node.bars.push({ x, y: yTop, w: barW, h: Math.max(yBot - yTop, 0.5), color })
        dLbl(si, i, x + barW / 2, (yTop + yBot) / 2 - dlSize * 0.55, ser.values[i]!, true)
      })
    }
  } else if (barSeriesIdx.length) {
    // clustered col: gapWidth is the category-group gap as a percentage of one bar's width
    const gap = (model.gapWidthPct ?? 150) / 100
    const sCount = Math.max(barSeriesIdx.length, 1)
    const barW = slotW / (sCount + gap)
    const groupW = barW * sCount
    const base = Math.max(min, 0) // autoZero baseline; when axis min>0, bars start from the axis bottom
    barSeriesIdx.forEach((si, slot) => {
      const ser = model.series[si]!
      const color = seriesColor(si)
      ser.values.forEach((v, i) => {
        if (v == null || i >= n) return
        const x = plot.x + i * slotW + (slotW - groupW) / 2 + slot * barW
        const yTop = yOf(Math.max(v, base))
        const yBot = yOf(Math.min(v, base))
        // Explicit per-point colors (c:dPt, varyColors multi-color single-series bars) win over the series color
        node.bars.push({
          x,
          y: yTop,
          w: barW,
          h: Math.max(yBot - yTop, 0.5),
          color: ser.pointColors?.[i] ?? color,
        })
        // 3D bars: the label clears the box's top face (its back edge rises depth3d above the front top)
        dLbl(
          si,
          i,
          x + barW / 2,
          v >= 0 ? yTop - dlSize * 1.15 - depth3d : yBot + dlSize * 0.15 + depth3d,
          v,
          false,
        )
      })
    })
  }
  {
    const lineW = Math.max(1.5, ptToPx(1.5, vp.scale))
    const markerR = Math.max(2, ptToPx(3, vp.scale))
    // Stacked areas accumulate per category; percentStacked normalizes to column totals
    const areaCum: number[] = new Array(n).fill(0)
    const areaTotals: number[] = Array.from({ length: n }, (_, i) =>
      model.series.reduce(
        (a, s) => a + (serKind(s) === 'area' ? Math.abs(s.values[i] ?? 0) : 0),
        0,
      ),
    )
    model.series.forEach((ser, si) => {
      const k = serKind(ser)
      if (k === 'bar') return
      // Dual axis: series on the secondary axis map y via the right-axis range
      const secSer = onSecAxis(ser)
      const yOfSer = secSer ? yOf2 : yOf
      const color = seriesColor(si)
      // Stacked area: bands accumulate per category (even for sparse series) and label
      // at the band center; handled before the pts-based path and its length guard
      if (k === 'area' && stacked && !secSer) {
        const top: number[] = []
        const bottom: number[] = []
        for (let i = 0; i < n; i++) {
          let v = ser.values[i] ?? 0
          if (grouping === 'percentStacked') v = (v / (areaTotals[i] || 1)) * 100
          const x = plot.x + (i + 0.5) * slotW
          const y0 = yOf(areaCum[i]!)
          bottom.push(x, y0)
          areaCum[i]! += v
          const y1 = yOf(areaCum[i]!)
          top.push(x, y1)
          if (ser.values[i] != null)
            dLbl(si, i, x, (y0 + y1) / 2 - dlSize * 0.55, ser.values[i]!, true)
        }
        for (let i = bottom.length - 2; i >= 0; i -= 2) top.push(bottom[i]!, bottom[i + 1]!)
        node.polylines.push({ points: top, color, widthPx: 1, closed: true, fill: color })
        return
      }
      const pts: number[] = []
      ser.values.forEach((v, i) => {
        if (v == null || i >= n) return
        const x = plot.x + (i + 0.5) * slotW
        const y = yOfSer(v)
        pts.push(x, y)
        if (k === 'line' && ser.marker) node.markers.push({ x, y, r: markerR, color })
        dLbl(si, i, x, y - dlSize * 1.3, v, false)
      })
      if (pts.length < 4) return
      // Stock series never connect points; whiskers/up-down bars are drawn below
      // (overlay line series from a combined c:lineChart keep their connectors)
      if (ser.fromStock) return
      if (k === 'area') {
        // Area chart: the polyline closes to the zero baseline (clamped to the axis when out of range), solid fill
        const sMin = secSer ? sec!.min : min
        const sMax = secSer ? sec!.max : max
        const baseY = yOfSer(Math.min(Math.max(0, sMin), sMax))
        node.polylines.push({
          points: [pts[0]!, baseY, ...pts, pts[pts.length - 2]!, baseY],
          color,
          widthPx: 1,
          closed: true,
          fill: color,
        })
      } else {
        const w = ser.lineWidthPt ? Math.max(1, ptToPx(ser.lineWidthPt, vp.scale)) : lineW
        node.polylines.push({
          points: pts,
          color,
          widthPx: w,
          ...(ser.smooth ? { smooth: true } : {}),
          ...(ser.dash ? { dash: dashArray(ser.dash, w) } : {}),
        })
      }
    })
  }

  // ── Stock whiskers / up-down bars ────────────────────────────────
  if (model.stock) {
    const stockSers = model.series.filter((s) => s.fromStock)
    // Volume-OHLC combos put prices on the secondary axis — map through the stock series' own axis
    const yOfS = stockSers[0] && onSecAxis(stockSers[0]) ? yOf2 : yOf
    // Roles by document order: OHLC with 4+ series, HLC with 3, HL with 2
    const openSer = stockSers.length >= 4 ? stockSers[0] : undefined
    const closeSer = stockSers.length >= 3 ? stockSers[stockSers.length - 1] : undefined
    const lineWidth = Math.max(1, ptToPx(1, vp.scale))
    const barW = slotW / (1 + (model.stock.gapWidthPct ?? 150) / 100)
    for (let i = 0; i < n; i++) {
      const vals = stockSers.map((s) => s.values[i]).filter((v): v is number => v != null)
      if (vals.length < 2) continue
      const x = plot.x + (i + 0.5) * slotW
      if (model.stock.hiLowLines) {
        node.axisLines.push({
          x1: x,
          y1: yOfS(Math.max(...vals)),
          x2: x,
          y2: yOfS(Math.min(...vals)),
          color: '#000000',
          widthPx: lineWidth,
        })
      }
      const open = openSer?.values[i]
      const close = closeSer?.values[i]
      if (model.stock.upDownBars && open != null && close != null && open !== close) {
        const y1 = yOfS(Math.max(open, close))
        const y2 = yOfS(Math.min(open, close))
        node.polylines.push({
          points: [x - barW / 2, y1, x + barW / 2, y1, x + barW / 2, y2, x - barW / 2, y2],
          color: '#000000',
          widthPx: lineWidth,
          closed: true,
          fill: close >= open ? '#FFFFFF' : '#404040',
        })
      }
    }
  }

  // ── Legend ──────────────────────────────────────────────────────
  if (legendPos && model.series.some((s) => s.name)) {
    // PowerPoint legend swatches are ~0.5em squares (measured 10px at 18pt text)
    const sw = legendSizePx * 0.5
    const legendTextStyle: RunStyle = {
      fontFamily: LABEL_FONT,
      fontSizePx: legendSizePx,
      bold: legendBold,
      italic: false,
    }
    const legendMeasure = (t: string) => metrics.measure(t, legendTextStyle)
    const items = model.series.map((s, i) => ({
      label: s.name ?? '',
      color: seriesColor(i),
    }))
    const itemWs = items.map((it) => sw + 4 + legendMeasure(it.label) + legendSizePx * 0.5)
    const legendEntry = (x: number, y: number, it: (typeof items)[number]) => {
      node.swatches.push({
        x,
        y: y + legendSizePx * 0.3,
        w: sw,
        h: legendSizePx * 0.5,
        color: it.color,
      })
      node.labels.push({
        text: it.label,
        x: x + sw + 4,
        y,
        fontSizePx: legendSizePx,
        color: labelColor,
        ...(legendBold ? { bold: true } : {}),
      })
    }
    if (legendPos === 't' || legendPos === 'b') {
      const total = itemWs.reduce((a, b) => a + b, 0)
      let x = Math.max((box.w - total) / 2, pad)
      const y = legendPos === 't' ? pad : box.h - pad - legendSizePx * 1.2
      items.forEach((it, i) => {
        legendEntry(x, y, it)
        x += itemWs[i]!
      })
    } else {
      // Right column (l/r laid out vertically centered like PowerPoint, tr pinned at the top);
      // with a secondary axis, shift right to clear its tick labels
      const colH = items.length * legendSizePx * 1.5
      let y = legendPos === 'tr' ? plot.y : Math.max(plot.y + (plot.h - colH) / 2, pad)
      // Overlay legends float inside the right frame edge (the plot reserved no space)
      const x = model.legendOverlay
        ? Math.max(box.w - Math.max(...itemWs, 0) - pad, plot.x + 4)
        : plot.x + plot.w + 8 + (sec ? y2LabelW + axisTitle2W + 8 : 0)
      items.forEach((it) => {
        legendEntry(x, y, it)
        y += legendSizePx * 1.5
      })
    }
  }

  return node
}

// ── Pie / doughnut charts ──────────────────────────────────────────

function buildPieNode(
  id: string,
  sourceId: string,
  model: ChartModel,
  box: PlacedBox,
  vp: Viewport,
  metrics: FontMetricsProvider,
): ChartRenderNode | null {
  // Pie uses the first series; wedge colors prefer per-point colors (c:dPt), default palette by category index
  const ser = model.series[0]
  if (!ser) return null
  const vals = ser.values.map((v) => (v != null && v > 0 ? v : 0))
  const total = vals.reduce((a, b) => a + b, 0)
  if (total <= 0) return null

  const node: ChartRenderNode = {
    id,
    type: 'chart',
    box,
    sourceId,
    gridLines: [],
    axisLines: [],
    labels: [],
    bars: [],
    polylines: [],
    markers: [],
    swatches: [],
    wedges: [],
  }

  const labelSizePx = ptToPx(chartTextPt(model), vp.scale)
  const style: RunStyle = {
    fontFamily: LABEL_FONT,
    fontSizePx: labelSizePx,
    bold: false,
    italic: false,
  }
  const measure = (text: string) => metrics.measure(text, style)
  const palette = chartPalette(model)
  const sliceColor = (i: number) =>
    ser.pointColors?.[i] ??
    (model.varyColors === false ? (ser.color ?? palette[0]!) : palette[i % palette.length]!)
  const pad = Math.max(6, Math.min(box.w, box.h) * 0.03)

  // Legend space (without a legend, the whole box goes to the pie)
  const legendPos = model.legendPos
  const legendItems = model.categories.map((cat, i) => ({
    label: cat,
    color: sliceColor(i),
  }))
  const legendRowH = labelSizePx * 1.5
  let plotW = box.w - pad * 2
  let plotH = box.h - pad * 2
  let plotX = pad
  let plotY = pad
  const sideLegendW =
    legendPos === 'l' || legendPos === 'r' || legendPos === 'tr'
      ? Math.min(
          box.w * 0.4,
          Math.max(...legendItems.map((it) => measure(it.label)), 0) + labelSizePx * 2.2,
        )
      : 0
  if (legendPos === 'r' || legendPos === 'tr') plotW -= sideLegendW
  else if (legendPos === 'l') {
    plotW -= sideLegendW
    plotX += sideLegendW
  } else if (legendPos === 't') {
    plotY += legendRowH
    plotH -= legendRowH
  } else if (legendPos === 'b') plotH -= legendRowH

  // Explosion: slice offset = (explosion/100)·diameter along its mid-angle; the radius
  // shrinks so the exploded footprint stays ≈ the unexploded one (PowerPoint keeps bounds)
  const explAt = (i: number) =>
    Math.max(ser.pointExplosionPct?.[i] ?? ser.explosionPct ?? 0, 0) / 100
  const maxExpl = vals.reduce((m, v, i) => (v > 0 ? Math.max(m, explAt(i)) : m), 0)
  const outerR = Math.max(Math.min(plotW, plotH) / 2, 5) / (1 + 2 * maxExpl)
  const cx = plotX + plotW / 2
  let cy = plotY + plotH / 2
  const innerR = (outerR * Math.min(Math.max(model.holePct ?? 0, 0), 90)) / 100

  // Pseudo-3D pie: the tilted disc projects to an ellipse (ry = sin(rotX) · rx) with a darker
  // side rim below the front-facing arcs; wedges become elliptical sector paths.
  // Rim depth 0.38·ry, solid (ellipse+rim) at 81% of plot height and vertically centered —
  // all three calibrated against PowerPoint (3D Pie-O12-PPT-Charts, rotX=30).
  const p3d = !!model.pseudo3D && innerR === 0
  const kY = p3d
    ? Math.min(Math.max(Math.sin(((model.rotXDeg ?? 30) * Math.PI) / 180), 0.35), 0.9)
    : 1
  const DEPTH_K = 0.38
  const rx = p3d
    ? Math.max(Math.min(plotW / 2, (plotH * 0.81) / (kY * (2 + DEPTH_K))), 5) / (1 + 2 * maxExpl)
    : outerR
  const ry = rx * kY
  const depth = p3d ? ry * DEPTH_K : 0
  // cy intentionally uses the explosion-shrunk ry: PowerPoint re-centers the exploded solid
  // (measured 10px down at explosion=25; this formula lands within 0.4px, fixed cy is ~11px off)
  if (p3d) cy = plotY + (plotH - ry * (2 + DEPTH_K)) / 2 + ry
  const explOffset = (startDeg: number, sweep: number, i: number) => {
    const mid = ((startDeg + sweep / 2) * Math.PI) / 180
    const off = explAt(i) * 2 * rx
    return { dx: Math.cos(mid) * off, dy: Math.sin(mid) * off * kY }
  }
  const ptAt = (deg: number, dx = 0, dy = 0) => {
    const t = (deg * Math.PI) / 180
    return { x: cx + dx + Math.cos(t) * rx, y: cy + dy + Math.sin(t) * ry }
  }
  if (p3d) {
    node.paths = node.paths ?? []
    // Rim: the front half is the parametric range [0°, 180°) (screen lower half)
    let a = -90 + (model.firstSliceAngDeg ?? 0)
    vals.forEach((v, i) => {
      if (v <= 0) return
      const sweep = (v / total) * 360
      const { dx, dy } = explOffset(a, sweep, i)
      // normalize wedge interval into [-180, 180) then clamp to the front range [0, 180]
      for (const off of [-360, 0, 360]) {
        const b1 = Math.max(a + off, 0)
        const b2 = Math.min(a + off + sweep, 180)
        if (b2 <= b1) continue
        const p1 = ptAt(b1, dx, dy)
        const p2 = ptAt(b2, dx, dy)
        const large = b2 - b1 > 180 ? 1 : 0
        node.paths!.push({
          d:
            `M ${p1.x} ${p1.y} A ${rx} ${ry} 0 ${large} 1 ${p2.x} ${p2.y} ` +
            `L ${p2.x} ${p2.y + depth} A ${rx} ${ry} 0 ${large} 0 ${p1.x} ${p1.y + depth} Z`,
          fill: shade(sliceColor(i), 0.72),
        })
      }
      a += sweep
    })
  }

  // Wedges: start at 12 o'clock (Konva rotation 0 = 3 o'clock, hence -90°), clockwise
  let angle = -90 + (model.firstSliceAngDeg ?? 0)
  vals.forEach((v, i) => {
    if (v <= 0) return
    const sweep = (v / total) * 360
    const { dx, dy } = explOffset(angle, sweep, i)
    if (p3d) {
      const p1 = ptAt(angle, dx, dy)
      if (sweep >= 359.999) {
        // Full circle: an SVG arc with coincident endpoints renders nothing, so use two half arcs
        const pm = ptAt(angle + 180, dx, dy)
        node.paths!.push({
          d:
            `M ${p1.x} ${p1.y} A ${rx} ${ry} 0 1 1 ${pm.x} ${pm.y} ` +
            `A ${rx} ${ry} 0 1 1 ${p1.x} ${p1.y} Z`,
          fill: sliceColor(i),
          stroke: '#ffffff',
        })
      } else {
        const p2 = ptAt(angle + sweep, dx, dy)
        const large = sweep > 180 ? 1 : 0
        node.paths!.push({
          d: `M ${cx + dx} ${cy + dy} L ${p1.x} ${p1.y} A ${rx} ${ry} 0 ${large} 1 ${p2.x} ${p2.y} Z`,
          fill: sliceColor(i),
          stroke: '#ffffff',
        })
      }
    } else {
      node.wedges!.push({
        cx: cx + dx,
        cy: cy + dy,
        outerR,
        innerR,
        startDeg: angle,
        sweepDeg: sweep,
        color: sliceColor(i),
      })
    }
    if (model.series[0]?.dataLabels ?? model.dataLabels) {
      // At the wedge midline radius: doughnut uses the ring-band midpoint, pie uses 2/3 radius
      const midRad = ((angle + sweep / 2) * Math.PI) / 180
      const r = innerR > 0 ? (innerR + outerR) / 2 : outerR * 0.66
      const valueText = model.dataLabelsPct ? `${Math.round((v / total) * 100)}%` : fmtNum(v)
      const text = composeDataLabel(model, 0, i, valueText)
      node.labels.push({
        text,
        x: cx + dx + Math.cos(midRad) * (p3d ? rx * 0.66 : r) - measure(text) / 2,
        y: cy + dy + Math.sin(midRad) * (p3d ? ry * 0.66 : r) - labelSizePx * 0.55,
        fontSizePx: labelSizePx * 0.9,
        color: '#FFFFFF',
      })
    }
    angle += sweep
  })

  // Legend
  if (legendPos) {
    const sw = labelSizePx * 0.5
    if (legendPos === 't' || legendPos === 'b') {
      const itemWs = legendItems.map((it) => sw + 4 + measure(it.label) + labelSizePx * 0.5)
      const totalW = itemWs.reduce((a, b) => a + b, 0)
      let x = Math.max((box.w - totalW) / 2, pad)
      const y = legendPos === 't' ? pad * 0.5 : box.h - pad * 0.5 - labelSizePx * 1.2
      legendItems.forEach((it, i) => {
        node.swatches.push({
          x,
          y: y + labelSizePx * 0.25,
          w: sw,
          h: labelSizePx * 0.5,
          color: it.color,
        })
        node.labels.push({
          text: it.label,
          x: x + sw + 4,
          y,
          fontSizePx: labelSizePx,
          color: '#666666',
        })
        x += itemWs[i]!
      })
    } else {
      const x = legendPos === 'l' ? pad : box.w - sideLegendW
      let y = Math.max(cy - (legendItems.length * legendRowH) / 2, pad)
      for (const it of legendItems) {
        node.swatches.push({
          x,
          y: y + labelSizePx * 0.25,
          w: sw,
          h: labelSizePx * 0.5,
          color: it.color,
        })
        node.labels.push({
          text: it.label,
          x: x + sw + 4,
          y,
          fontSizePx: labelSizePx,
          color: '#666666',
        })
        y += legendRowH
      }
    }
  }

  return node
}

// ── Horizontal bars (barDir='bar': categories on the y axis, values on the x axis) ──

function buildHBarNode(
  id: string,
  sourceId: string,
  model: ChartModel,
  box: PlacedBox,
  vp: Viewport,
  metrics: FontMetricsProvider,
): ChartRenderNode | null {
  const grouping = model.grouping ?? 'clustered'
  const stacked = grouping === 'stacked' || grouping === 'percentStacked'
  const node = emptyChartNode(id, sourceId, box)

  const labelSizePx = ptToPx(model.valAxis?.labelSizePt ?? chartTextPt(model), vp.scale)
  const labelColor = model.valAxis?.labelColor ?? '#666666'
  const catLabelSizePx = ptToPx(
    model.catAxis?.labelSizePt ?? model.valAxis?.labelSizePt ?? chartTextPt(model),
    vp.scale,
  )
  const catLabelColor = model.catAxis?.labelColor ?? labelColor
  const style = (sizePx: number): RunStyle => ({
    fontFamily: LABEL_FONT,
    fontSizePx: sizePx,
    bold: false,
    italic: false,
  })
  const measure = (text: string, sizePx: number) => metrics.measure(text, style(sizePx))
  const palette = chartPalette(model)
  const seriesColor = (i: number) =>
    model.series[i]?.color ?? palette[(model.series[i]?.paletteIdx ?? i) % palette.length]!

  const allVals = model.series.flatMap((s) => s.values.filter((v): v is number => v != null))
  if (!allVals.length) return null
  const catCount = Math.max(model.categories.length, ...model.series.map((s) => s.values.length), 1)
  const catAbsTotals = Array.from({ length: catCount }, (_, i) =>
    model.series.reduce((a, s) => a + Math.abs(s.values[i] ?? 0), 0),
  )
  const valueAt = (si: number, i: number): number | null => {
    const v = model.series[si]?.values[i]
    if (v == null) return null
    if (grouping !== 'percentStacked') return v
    return (v / (catAbsTotals[i] || 1)) * 100
  }
  let dataMax: number
  let dataMin: number
  if (stacked) {
    const posSums = Array.from({ length: catCount }, (_, i) =>
      model.series.reduce((a, _s, si) => a + Math.max(valueAt(si, i) ?? 0, 0), 0),
    )
    const negSums = Array.from({ length: catCount }, (_, i) =>
      model.series.reduce((a, _s, si) => a + Math.min(valueAt(si, i) ?? 0, 0), 0),
    )
    dataMax = Math.max(...posSums, 0)
    dataMin = Math.min(...negSums, 0)
  } else {
    dataMax = Math.max(...allVals, 0)
    dataMin = Math.min(...allVals, 0)
  }
  // Percent stacked: the value axis is exactly 0-100% (−100% with negative stacks)
  if (grouping === 'percentStacked') {
    dataMax = 100
    dataMin = dataMin < 0 ? -100 : 0
  }
  // Horizontal value ticks also space by physical length: label width bounds the
  // interval count (python-pptx 6-up barH sheets measured: ~330px plots show 2
  // intervals - range 5.25 -> unit 5, 12.9 -> 10, 100% -> 50%)
  const maxIntervalsH = Math.max(2, Math.min(8, Math.floor(box.w / (labelSizePx * 4.5))))
  const { min, max, ticks } = ppTicks(
    model.valAxis?.min ?? dataMin,
    model.valAxis?.max ?? dataMax,
    model.valAxis?.min == null && grouping !== 'percentStacked',
    model.valAxis?.max == null && grouping !== 'percentStacked',
    true,
    maxIntervalsH,
  )

  // Layout: left category-label width + bottom value-tick row + legend
  const pad = Math.max(4, box.w * 0.01)
  const legendPos = model.legendPos
  const legendH = legendPos === 't' || legendPos === 'b' ? labelSizePx * 1.6 : 0
  const catLabelsOff =
    !!model.catAxis?.hidden || !!model.catAxis?.tickLblHidden || !!model.catAxis?.tickLblGarbage
  // garbage labels hide the text but keep their reserved width (same split as the vertical builder)
  const catNoReserve = !!model.catAxis?.hidden || !!model.catAxis?.tickLblHidden
  const catLabelW = catNoReserve
    ? 0
    : Math.max(...model.categories.map((c) => measure(c, catLabelSizePx)), 0)
  const plotX = pad + Math.min(catLabelW, box.w * 0.35) + 8
  const plotY = pad + (legendPos === 't' ? legendH + 4 : 0) + labelSizePx * 0.6
  const plotR = box.w - pad - labelSizePx * 0.7
  const plotB = box.h - pad - labelSizePx * 1.5 - (legendPos === 'b' ? legendH : 0)
  const plot = {
    x: plotX,
    y: plotY,
    w: Math.max(plotR - plotX, 10),
    h: Math.max(plotB - plotY, 10),
  }

  const xOf = (v: number) => plot.x + plot.w * ((v - min) / (max - min || 1))

  // Value ticks (x axis, bottom) + vertical grid
  const gridColor = model.valAxis?.gridColor
  const tickLabels = ticks.map((t) => (grouping === 'percentStacked' ? `${fmtNum(t)}%` : fmtNum(t)))
  ticks.forEach((t, i) => {
    const x = xOf(t)
    if (gridColor && t !== min) {
      node.gridLines.push({
        x1: x,
        y1: plot.y,
        x2: x,
        y2: plot.y + plot.h,
        color: gridColor,
        ...(model.valAxis?.gridDash ? { dash: [4, 4] } : {}),
      })
    }
    const text = tickLabels[i]!
    node.labels.push({
      text,
      x: x - measure(text, labelSizePx) / 2,
      y: plot.y + plot.h + labelSizePx * 0.35,
      fontSizePx: labelSizePx,
      color: labelColor,
    })
  })
  const axisColor = model.valAxis?.lineColor ?? model.catAxis?.lineColor ?? '#888888'
  const axisW = Math.max(1, ptToPx(1, vp.scale))
  node.axisLines.push({
    x1: plot.x,
    y1: plot.y + plot.h,
    x2: plot.x + plot.w,
    y2: plot.y + plot.h,
    color: axisColor,
    widthPx: axisW,
  })
  node.axisLines.push({
    x1: plot.x,
    y1: plot.y,
    x2: plot.x,
    y2: plot.y + plot.h,
    color: axisColor,
    widthPx: axisW,
  })

  // Category labels (y axis, left). PowerPoint default (minMax) puts the first category at the
  // BOTTOM for horizontal bars; c:orientation maxMin flips the first one to the top
  const n = Math.max(model.categories.length, 1)
  const slotH = plot.h / n
  const rowY = (i: number) => {
    const pos = model.catAxis?.reversed ? i : n - 1 - i
    return plot.y + pos * slotH
  }
  if (!catLabelsOff)
    model.categories.forEach((cat, i) => {
      node.labels.push({
        text: cat,
        x: plot.x - 8 - measure(cat, catLabelSizePx),
        y: rowY(i) + slotH / 2 - catLabelSizePx * 0.55,
        fontSizePx: catLabelSizePx,
        color: catLabelColor,
      })
    })

  // Data bars
  const gap = (model.gapWidthPct ?? 150) / 100
  const dlSize = labelSizePx * 0.9
  const dLbl = (
    si: number,
    catIdx: number,
    x: number,
    yMid: number,
    v: number,
    inside: boolean,
  ) => {
    if (!(model.series[si]?.dataLabels ?? model.dataLabels)) return
    const text = composeDataLabel(model, si, catIdx, fmtDataLabel(v, model.dataLabelFmt))
    node.labels.push({
      text,
      x: inside ? x - measure(text, dlSize) / 2 : x,
      y: yMid - dlSize * 0.55,
      fontSizePx: dlSize,
      color: inside ? '#FFFFFF' : '#404040',
    })
  }
  if (stacked) {
    const barH = slotH / (1 + gap)
    for (let i = 0; i < n; i++) {
      const y = rowY(i) + (slotH - barH) / 2
      let posAcc = 0
      let negAcc = 0
      model.series.forEach((ser, si) => {
        const v = valueAt(si, i)
        if (v == null || v === 0) return
        const from = v > 0 ? posAcc : negAcc
        const to = from + v
        if (v > 0) posAcc = to
        else negAcc = to
        const xL = xOf(Math.min(from, to))
        const xR = xOf(Math.max(from, to))
        node.bars.push({
          x: xL,
          y,
          w: Math.max(xR - xL, 0.5),
          h: barH,
          color: ser.pointColors?.[i] ?? seriesColor(si),
        })
        dLbl(si, i, (xL + xR) / 2, y + barH / 2, ser.values[i]!, true)
      })
    }
  } else {
    const sCount = Math.max(model.series.length, 1)
    const barH = slotH / (sCount + gap)
    const groupH = barH * sCount
    const base = Math.max(min, 0)
    model.series.forEach((ser, si) => {
      const color = seriesColor(si)
      ser.values.forEach((v, i) => {
        if (v == null || i >= n) return
        // PowerPoint stacks horizontal-bar series bottom-up: series 1 sits nearest the
        // category axis (the bottom of the group), later series above it
        const y = rowY(i) + (slotH - groupH) / 2 + (sCount - 1 - si) * barH
        const xL = xOf(Math.min(v, base))
        const xR = xOf(Math.max(v, base))
        node.bars.push({
          x: xL,
          y,
          w: Math.max(xR - xL, 0.5),
          h: barH,
          color: ser.pointColors?.[i] ?? color,
        })
        const lblText = composeDataLabel(model, si, i, fmtDataLabel(v, model.dataLabelFmt))
        dLbl(si, i, v >= 0 ? xR + 4 : xL - 4 - measure(lblText, dlSize), y + barH / 2, v, false)
      })
    })
  }

  // legend order mirrors the bottom-up series stacking (Series N listed first)
  addSeriesLegend(node, model, box, plot, labelSizePx, measure, pad, seriesColor, true)
  return node
}

// ── Scatter charts (dual value axes) ───────────────────────────────────

function buildScatterNode(
  id: string,
  sourceId: string,
  model: ChartModel,
  box: PlacedBox,
  vp: Viewport,
  metrics: FontMetricsProvider,
): ChartRenderNode | null {
  const node = emptyChartNode(id, sourceId, box)
  const labelSizePx = ptToPx(model.valAxis?.labelSizePt ?? chartTextPt(model), vp.scale)
  const labelColor = model.valAxis?.labelColor ?? '#666666'
  const style = (sizePx: number): RunStyle => ({
    fontFamily: LABEL_FONT,
    fontSizePx: sizePx,
    bold: false,
    italic: false,
  })
  const measure = (text: string, sizePx: number) => metrics.measure(text, style(sizePx))
  const palette = chartPalette(model)
  const seriesColor = (i: number) =>
    model.series[i]?.color ?? palette[(model.series[i]?.paletteIdx ?? i) % palette.length]!

  // Point sets: x defaults to indices 1..n; i keeps the source row (bubbleSizes align by it)
  const points = model.series.map((s) =>
    s.values
      .map((y, i) => ({ x: s.xValues?.[i] ?? i + 1, y, i }))
      .filter((p): p is { x: number; y: number; i: number } => p.x != null && p.y != null),
  )
  const allX = points.flat().map((p) => p.x)
  const allY = points.flat().map((p) => p.y)
  if (!allY.length) return null

  // Bubbles need axis headroom: PowerPoint widens auto ranges by ~half the largest
  // bubble's diameter (25% of the span × bubbleScale) so bubbles stay inside the plot
  const hasBubbles = model.series.some((s) => s.bubbleSizes?.length)
  const bubblePad = (lo: number, hi: number) =>
    hasBubbles ? ((hi - lo) * 0.25 * ((model.bubbleScale ?? 100) / 100)) / 2 : 0
  const xPad = bubblePad(Math.min(...allX, 0), Math.max(...allX, 0))
  const yPad = bubblePad(Math.min(...allY, 0), Math.max(...allY, 0))

  // The catAxis slot = the x axis (the engine dispatched by axPos)
  const xTicksR = ppTicks(
    model.catAxis?.min ?? Math.min(...allX, 0) - (Math.min(...allX, 0) < 0 ? xPad : 0),
    model.catAxis?.max ?? Math.max(...allX, 0) + xPad,
    model.catAxis?.min == null,
    model.catAxis?.max == null,
    true,
  )
  const yTicksR = ppTicks(
    model.valAxis?.min ?? Math.min(...allY, 0) - (Math.min(...allY, 0) < 0 ? yPad : 0),
    model.valAxis?.max ?? Math.max(...allY, 0) + yPad,
    model.valAxis?.min == null,
    model.valAxis?.max == null,
    true,
  )

  const pad = Math.max(4, box.w * 0.01)
  const legendPos = model.legendPos
  const legendH = legendPos === 't' || legendPos === 'b' ? labelSizePx * 1.6 : 0
  const yLabelW = Math.max(...yTicksR.ticks.map((t) => measure(fmtNum(t), labelSizePx)), 0)
  const plotX = pad + yLabelW + 10
  const plotY = pad + (legendPos === 't' ? legendH + 4 : 0) + labelSizePx * 0.6
  const plotR = box.w - pad - labelSizePx * 0.7
  const plotB = box.h - pad - labelSizePx * 1.5 - (legendPos === 'b' ? legendH : 0)
  const plot = {
    x: plotX,
    y: plotY,
    w: Math.max(plotR - plotX, 10),
    h: Math.max(plotB - plotY, 10),
  }

  const xOf = (v: number) =>
    plot.x + plot.w * ((v - xTicksR.min) / (xTicksR.max - xTicksR.min || 1))
  const yOf = (v: number) =>
    plot.y + plot.h * (1 - (v - yTicksR.min) / (yTicksR.max - yTicksR.min || 1))

  // Grid + ticks: y (horizontal grid) follows valAxis, x (vertical grid) follows the catAxis slot
  const yGrid = model.valAxis?.gridColor
  yTicksR.ticks.forEach((t) => {
    const y = yOf(t)
    if (yGrid && t !== yTicksR.min) {
      node.gridLines.push({
        x1: plot.x,
        y1: y,
        x2: plot.x + plot.w,
        y2: y,
        color: yGrid,
        ...(model.valAxis?.gridDash ? { dash: [4, 4] } : {}),
      })
    }
    const text = fmtNum(t)
    node.labels.push({
      text,
      x: plot.x - 6 - measure(text, labelSizePx),
      y: y - labelSizePx * 0.55,
      fontSizePx: labelSizePx,
      color: labelColor,
    })
  })
  const xGrid = model.catAxis?.gridColor
  xTicksR.ticks.forEach((t) => {
    const x = xOf(t)
    if (xGrid && t !== xTicksR.min) {
      node.gridLines.push({
        x1: x,
        y1: plot.y,
        x2: x,
        y2: plot.y + plot.h,
        color: xGrid,
        ...(model.catAxis?.gridDash ? { dash: [4, 4] } : {}),
      })
    }
    const text = fmtNum(t)
    node.labels.push({
      text,
      x: x - measure(text, labelSizePx) / 2,
      y: plot.y + plot.h + labelSizePx * 0.35,
      fontSizePx: labelSizePx,
      color: model.catAxis?.labelColor ?? labelColor,
    })
  })
  const axisColor = model.valAxis?.lineColor ?? model.catAxis?.lineColor ?? '#888888'
  const axisW = Math.max(1, ptToPx(1, vp.scale))
  node.axisLines.push({
    x1: plot.x,
    y1: plot.y + plot.h,
    x2: plot.x + plot.w,
    y2: plot.y + plot.h,
    color: axisColor,
    widthPx: axisW,
  })
  node.axisLines.push({
    x1: plot.x,
    y1: plot.y,
    x2: plot.x,
    y2: plot.y + plot.h,
    color: axisColor,
    widthPx: axisW,
  })

  // Series: scatterStyle decides the line/marker defaults; an explicit series marker overrides
  const st = model.scatterStyle ?? 'lineMarker'
  const hasLine = st.startsWith('line') || st.startsWith('smooth')
  const smooth = st.startsWith('smooth')
  const defaultMarker = st !== 'line' && st !== 'smooth' && st !== 'none'
  const lineW = Math.max(1.5, ptToPx(1.5, vp.scale))
  const markerR = Math.max(2, ptToPx(3, vp.scale))
  // Bubble: largest bubble diameter = 25% of the smaller plot side × bubbleScale%; radius ∝ √size
  const maxBubbleSize = Math.max(
    ...model.series.flatMap((s) => (s.bubbleSizes ?? []).map((v) => Math.abs(v ?? 0))),
    0,
  )
  const maxBubbleR = (Math.min(plot.w, plot.h) * 0.25 * ((model.bubbleScale ?? 100) / 100)) / 2
  model.series.forEach((ser, si) => {
    const color = seriesColor(si)
    const pts = points[si]!
    const bubbles = ser.bubbleSizes
    const showMarker = ser.marker ?? defaultMarker
    const flat: number[] = []
    pts.forEach((p) => {
      const x = xOf(p.x)
      const y = yOf(p.y)
      flat.push(x, y)
      if (bubbles?.length && maxBubbleSize > 0) {
        const size = Math.abs(bubbles[p.i] ?? 0)
        if (size > 0)
          node.markers.push({
            x,
            y,
            r:
              maxBubbleR *
              (model.bubbleSizeIsWidth ? size / maxBubbleSize : Math.sqrt(size / maxBubbleSize)),
            color,
          })
      } else if (showMarker) node.markers.push({ x, y, r: markerR, color })
      if (ser.dataLabels ?? model.dataLabels) {
        const text = composeDataLabel(model, si, p.i, fmtNum(round12(p.y)))
        node.labels.push({
          text,
          x: x - measure(text, labelSizePx * 0.9) / 2,
          y: y - labelSizePx * 1.3,
          fontSizePx: labelSizePx * 0.9,
          color: '#404040',
        })
      }
    })
    if (hasLine && !bubbles?.length && flat.length >= 4) {
      node.polylines.push({
        points: flat,
        color,
        widthPx: lineW,
        ...(smooth || ser.smooth ? { smooth: true } : {}),
      })
    }
  })

  addSeriesLegend(node, model, box, plot, labelSizePx, measure, pad, seriesColor)
  return node
}

// ── chartEx: funnel / sunburst ─────────────────────────────────────────

function buildFunnelNode(
  id: string,
  sourceId: string,
  model: ChartModel,
  box: PlacedBox,
  vp: Viewport,
  metrics: FontMetricsProvider,
): ChartRenderNode | null {
  const vals = model.series[0]?.values
  if (!vals?.length) return null
  const maxVal = Math.max(...vals.map((v) => Math.abs(v ?? 0)))
  if (maxVal <= 0) return null
  const node = emptyChartNode(id, sourceId, box)

  const labelSizePx = ptToPx(chartTextPt(model), vp.scale) * 0.9
  const style: RunStyle = {
    fontFamily: LABEL_FONT,
    fontSizePx: labelSizePx,
    bold: false,
    italic: false,
  }
  const measure = (text: string) => metrics.measure(text, style)
  const color = model.series[0]!.color ?? chartPalette(model)[0]!
  const pad = Math.max(6, Math.min(box.w, box.h) * 0.03)

  const labelW = Math.max(...model.categories.map((c) => measure(c)), 0)
  const plotX = pad + labelW + labelSizePx
  const plotW = box.w - plotX - pad
  const plotH = box.h - pad * 2
  const n = vals.length
  const slotH = plotH / n
  const gapFrac = (model.gapWidthPct ?? 6) / 100
  const barH = slotH / (1 + gapFrac)
  vals.forEach((v, i) => {
    const y = pad + i * slotH + (slotH - barH) / 2
    const cat = model.categories[i] ?? ''
    if (cat) {
      node.labels.push({
        text: cat,
        x: pad + labelW - measure(cat),
        y: y + barH / 2 - labelSizePx * 0.6,
        fontSizePx: labelSizePx,
        color: '#595959',
      })
    }
    if (v == null || v <= 0) return
    const w = (plotW * v) / maxVal
    node.bars.push({ x: plotX + (plotW - w) / 2, y, w, h: barH, color })
  })
  return node
}

function buildSunburstNode(
  id: string,
  sourceId: string,
  model: ChartModel,
  box: PlacedBox,
  vp: Viewport,
  metrics: FontMetricsProvider,
): ChartRenderNode | null {
  const sb = model.sunburst
  if (!sb) return null
  const node = emptyChartNode(id, sourceId, box)
  node.wedges = []

  // Tree of ancestor paths; every point is a terminal node at its own depth
  interface TNode {
    label: string
    depth: number
    value: number
    children: TNode[]
    point?: number
  }
  const root: TNode = { label: '', depth: 0, value: 0, children: [] }
  const depth = sb.levels.length
  const n = sb.sizes.length
  for (let i = 0; i < n; i++) {
    const size = Math.abs(sb.sizes[i] ?? 0)
    if (size <= 0) continue
    let cur = root
    for (let d = 0; d < depth; d++) {
      const label = sb.levels[depth - 1 - d]?.[i] ?? ''
      if (!label) break
      // Merge by label: a row that ended at this ancestor earlier is the same node
      let child = cur.children.find((c) => c.label === label)
      if (!child) {
        child = { label, depth: d + 1, value: 0, children: [] }
        cur.children.push(child)
      }
      child.value += size
      cur = child
    }
    cur.point ??= i
    cur.value ||= size
  }
  if (root.children.length === 0) return null
  root.value = root.children.reduce((a, c) => a + c.value, 0)
  // PowerPoint orders every node's children by value, largest first
  const sortRec = (t: TNode) => {
    t.children.sort((a, b) => b.value - a.value)
    t.children.forEach(sortRec)
  }
  sortRec(root)

  const labelSizePx = ptToPx(chartTextPt(model), vp.scale) * 0.8
  const style: RunStyle = {
    fontFamily: LABEL_FONT,
    fontSizePx: labelSizePx,
    bold: false,
    italic: false,
  }
  const measure = (text: string) => metrics.measure(text, style)
  const palette = chartPalette(model)
  // Branch colors by data order (not sorted order), matching the legend/palette sequence
  const rootColors = new Map<string, string>()
  const rootLvl = sb.levels[depth - 1] ?? []
  for (const label of rootLvl) {
    if (label && !rootColors.has(label))
      rootColors.set(label, palette[rootColors.size % palette.length]!)
  }

  const pad = Math.max(4, Math.min(box.w, box.h) * 0.02)
  const R = Math.max(Math.min(box.w, box.h) / 2 - pad, 5)
  const cx = box.w / 2
  const cy = box.h / 2
  const maxDepth = (t: TNode): number => Math.max(t.depth, ...t.children.map(maxDepth))
  const rings = Math.max(maxDepth(root), 1)
  const holeR = R * 0.2
  const ringT = (R - holeR) / rings

  const drawNode = (t: TNode, startDeg: number, sweepDeg: number, color: string) => {
    const override = t.point != null ? sb.pointColors?.[t.point] : undefined
    const fill = override ?? color
    // A fully transparent per-point fill (alpha 00) hides the segment entirely
    if (!(override != null && /^#[0-9A-F]{6}00$/i.test(override))) {
      node.wedges!.push({
        cx,
        cy,
        innerR: holeR + (t.depth - 1) * ringT,
        outerR: holeR + t.depth * ringT,
        startDeg,
        sweepDeg,
        color: fill,
      })
      const midR = holeR + (t.depth - 0.5) * ringT
      const midDeg = startDeg + sweepDeg / 2
      const w = measure(t.label)
      const arcLen = (Math.abs(sweepDeg) / 360) * 2 * Math.PI * midR
      if (t.label && w < arcLen * 0.9 && labelSizePx < ringT * 0.9) {
        // Tangential label, kept upright: rotate by the mid angle folded into [-90°, 90°]
        let rot = midDeg + 90
        while (rot > 90) rot -= 180
        while (rot < -90) rot += 180
        const rad = (rot * Math.PI) / 180
        const px = cx + midR * Math.cos((midDeg * Math.PI) / 180)
        const py = cy + midR * Math.sin((midDeg * Math.PI) / 180)
        node.labels.push({
          text: t.label,
          x: px - (w / 2) * Math.cos(rad) + (labelSizePx / 2) * Math.sin(rad),
          y: py - (w / 2) * Math.sin(rad) - (labelSizePx / 2) * Math.cos(rad),
          fontSizePx: labelSizePx,
          color: '#FFFFFF',
          rotationDeg: rot,
        })
      }
    }
    let a = startDeg
    for (const c of t.children) {
      const s = (c.value / t.value) * sweepDeg
      drawNode(c, a, s, t.depth === 0 ? (rootColors.get(c.label) ?? palette[0]!) : fill)
      a += s
    }
  }
  let a = -90
  for (const c of root.children) {
    const s = (c.value / root.value) * 360
    drawNode(c, a, s, rootColors.get(c.label) ?? palette[0]!)
    a += s
  }
  return node
}

// ── Radar charts ───────────────────────────────────────────────────────

function buildRadarNode(
  id: string,
  sourceId: string,
  model: ChartModel,
  box: PlacedBox,
  vp: Viewport,
  metrics: FontMetricsProvider,
): ChartRenderNode | null {
  const n = Math.max(model.categories.length, ...model.series.map((s) => s.values.length))
  if (n < 3) return null
  const allVals = model.series.flatMap((s) => s.values.filter((v): v is number => v != null))
  if (!allVals.length) return null
  const node = emptyChartNode(id, sourceId, box)

  const labelSizePx = ptToPx(model.valAxis?.labelSizePt ?? chartTextPt(model), vp.scale)
  const labelColor = model.valAxis?.labelColor ?? '#666666'
  const catLabelColor = model.catAxis?.labelColor ?? labelColor
  const style = (sizePx: number): RunStyle => ({
    fontFamily: LABEL_FONT,
    fontSizePx: sizePx,
    bold: false,
    italic: false,
  })
  const measure = (text: string, sizePx: number) => metrics.measure(text, style(sizePx))
  const palette = chartPalette(model)
  const seriesColor = (i: number) =>
    model.series[i]?.color ?? palette[(model.series[i]?.paletteIdx ?? i) % palette.length]!

  const { min, max, ticks } = ppTicks(
    model.valAxis?.min ?? Math.min(...allVals, 0),
    model.valAxis?.max ?? Math.max(...allVals, 0),
    model.valAxis?.min == null,
    model.valAxis?.max == null,
    true,
  )

  const pad = Math.max(6, Math.min(box.w, box.h) * 0.03)
  const legendPos = model.legendPos
  const legendH = legendPos === 't' || legendPos === 'b' ? labelSizePx * 1.6 : 0
  const maxCatW = Math.max(...model.categories.map((c) => measure(c, labelSizePx)), 0)
  const sideLegendW =
    legendPos === 'l' || legendPos === 'r' || legendPos === 'tr'
      ? Math.max(...model.series.map((s) => measure(s.name ?? '', labelSizePx)), 0) +
        labelSizePx * 2.2
      : 0
  const plotW = box.w - pad * 2 - sideLegendW - maxCatW * 2
  const plotH = box.h - pad * 2 - legendH - labelSizePx * 2.4
  const R = Math.max(Math.min(plotW, plotH) / 2, 5)
  const cx = pad + maxCatW + plotW / 2 + (legendPos === 'l' ? sideLegendW : 0)
  const cy = pad + labelSizePx * 1.2 + (legendPos === 't' ? legendH : 0) + plotH / 2
  const plot = { x: cx - R, y: cy - R, w: R * 2, h: R * 2 }

  // Vertex directions: from 12 o'clock, clockwise
  const angleOf = (i: number) => -Math.PI / 2 + (i / n) * Math.PI * 2
  const rOf = (v: number) => (R * (v - min)) / (max - min || 1)
  const ptAt = (i: number, v: number): [number, number] => {
    const a = angleOf(i)
    const r = rOf(v)
    return [cx + Math.cos(a) * r, cy + Math.sin(a) * r]
  }

  // Grid rings (one n-gon per tick) + radial spokes
  const gridColor = model.valAxis?.gridColor ?? '#D9D9D9'
  for (const t of ticks) {
    if (t === min) continue
    const ring: number[] = []
    for (let i = 0; i < n; i++) ring.push(...ptAt(i, t))
    node.polylines.push({ points: ring, color: gridColor, widthPx: 1, closed: true })
  }
  for (let i = 0; i < n; i++) {
    const [x, y] = ptAt(i, max)
    node.gridLines.push({ x1: cx, y1: cy, x2: x, y2: y, color: gridColor })
  }

  // Value tick labels (along the 12 o'clock spoke)
  for (const t of ticks) {
    const [x, y] = ptAt(0, t)
    const text = fmtNum(t)
    node.labels.push({
      text,
      x: x - measure(text, labelSizePx) - 4,
      y: y - labelSizePx * 0.55,
      fontSizePx: labelSizePx * 0.9,
      color: labelColor,
    })
  }

  // Category labels (outside the vertices, alignment adjusted by direction)
  model.categories.forEach((cat, i) => {
    const a = angleOf(i)
    const lx = cx + Math.cos(a) * (R + labelSizePx * 0.5)
    const ly = cy + Math.sin(a) * (R + labelSizePx * 0.5)
    const w = measure(cat, labelSizePx)
    const alignX = Math.cos(a) > 0.3 ? lx : Math.cos(a) < -0.3 ? lx - w : lx - w / 2
    const alignY =
      Math.sin(a) > 0.3 ? ly : Math.sin(a) < -0.3 ? ly - labelSizePx : ly - labelSizePx * 0.55
    node.labels.push({
      text: cat,
      x: alignX,
      y: alignY,
      fontSizePx: labelSizePx,
      color: catLabelColor,
    })
  })

  // Series polygons
  const filled = model.radarStyle === 'filled'
  const markerDefault = model.radarStyle === 'marker'
  const lineW = Math.max(1.5, ptToPx(1.5, vp.scale))
  const markerR = Math.max(2, ptToPx(3, vp.scale))
  model.series.forEach((ser, si) => {
    const color = seriesColor(si)
    const flat: number[] = []
    for (let i = 0; i < n; i++) {
      const v = ser.values[i]
      if (v == null) continue
      const [x, y] = ptAt(i, v)
      flat.push(x, y)
      if (ser.marker ?? markerDefault) node.markers.push({ x, y, r: markerR, color })
    }
    if (flat.length >= 6) {
      node.polylines.push({
        points: flat,
        color,
        widthPx: lineW,
        closed: true,
        ...(filled ? { fill: withAlpha(color, 0.4) } : {}),
      })
    }
  })

  addSeriesLegend(node, model, box, plot, labelSizePx, measure, pad, seriesColor)
  return node
}

// ── Shared helpers ─────────────────────────────────────────────────────

function emptyChartNode(id: string, sourceId: string, box: PlacedBox): ChartRenderNode {
  return {
    id,
    type: 'chart',
    box,
    sourceId,
    gridLines: [],
    axisLines: [],
    labels: [],
    bars: [],
    polylines: [],
    markers: [],
    swatches: [],
  }
}

/** Series legend (shared by the newer chart types; t/b centered horizontally, l/r/tr as a top-right column). */
function addSeriesLegend(
  node: ChartRenderNode,
  model: ChartModel,
  box: PlacedBox,
  plot: { x: number; y: number; w: number; h: number },
  labelSizePx: number,
  measure: (text: string, sizePx: number) => number,
  pad: number,
  seriesColor: (i: number) => string,
  reverse = false,
): void {
  const legendPos = model.legendPos
  if (!legendPos || !model.series.some((s) => s.name)) return
  const sw = labelSizePx * 0.5
  const items = model.series.map((s, i) => ({
    label: s.name ?? '',
    color: seriesColor(i),
  }))
  if (reverse) items.reverse()
  const itemWs = items.map((it) => sw + 4 + measure(it.label, labelSizePx) + labelSizePx * 0.5)
  const labelColor = model.valAxis?.labelColor ?? '#666666'
  if (legendPos === 't' || legendPos === 'b') {
    const total = itemWs.reduce((a, b) => a + b, 0)
    let x = Math.max((box.w - total) / 2, pad)
    const y = legendPos === 't' ? pad : box.h - pad - labelSizePx * 1.2
    items.forEach((it, i) => {
      node.swatches.push({
        x,
        y: y + labelSizePx * 0.3,
        w: sw,
        h: labelSizePx * 0.5,
        color: it.color,
      })
      node.labels.push({
        text: it.label,
        x: x + sw + 4,
        y,
        fontSizePx: labelSizePx,
        color: labelColor,
      })
      x += itemWs[i]!
    })
  } else {
    let y = plot.y
    const x = plot.x + plot.w + 8
    items.forEach((it) => {
      node.swatches.push({
        x,
        y: y + labelSizePx * 0.3,
        w: sw,
        h: labelSizePx * 0.5,
        color: it.color,
      })
      node.labels.push({
        text: it.label,
        x: x + sw + 4,
        y,
        fontSizePx: labelSizePx,
        color: labelColor,
      })
      y += labelSizePx * 1.5
    })
  }
}

/** #RRGGBB(AA) → rgba() string (semi-transparent fill for filled radar charts). */
function withAlpha(hex: string, alpha: number): string {
  const m = /^#([0-9a-fA-F]{6})/.exec(hex)
  if (!m) return hex
  const v = parseInt(m[1]!, 16)
  return `rgba(${(v >> 16) & 255}, ${(v >> 8) & 255}, ${v & 255}, ${alpha})`
}

/**
 * Data range → ticks, mimicking PowerPoint/Excel automatic axes:
 * add 5% headroom to the data max, pick a step of 1/2/5×10^k (about 8 ticks max),
 * axis top = headroom value rounded up to a step multiple (e.g. max 24 → 30, 120 → 140, 715 → 800).
 */
function ppTicks(
  rawMin: number,
  rawMax: number,
  autoMin: boolean,
  autoMax: boolean,
  /** Legacy span/8 unit (horizontal value axes & scatter/radar keep PowerPoint's sparser ticks there) */
  legacy = false,
  /** Cap on the interval count (physical-axis-length rule); the unit only steps UP the 1/2/5 ladder */
  maxIntervals?: number,
): { min: number; max: number; ticks: number[] } {
  let lo = autoMin ? Math.min(rawMin, 0) : rawMin
  let hi = rawMax
  if (hi <= lo) hi = lo + 1
  // 5% headroom (auto ends only)
  const span0 = hi - lo
  const hiT = autoMax && hi > 0 ? hi + span0 * 0.05 : hi
  const loT = autoMin && lo < 0 ? lo - span0 * 0.05 : lo
  let step = legacy ? ppUnitLegacy((hiT - loT) / 8) : ppUnit(hiT - loT)
  if (maxIntervals) {
    while (Math.ceil((hiT - loT) / step) > maxIntervals) step = ppUnitUp(step)
  }
  if (autoMin) lo = Math.floor(loT / step) * step
  if (autoMax) hi = Math.ceil(hiT / step) * step
  const ticks: number[] = []
  for (let v = lo; v <= hi + step * 1e-6; v += step) ticks.push(round12(v))
  return { min: lo, max: hi, ticks }
}

/** Legacy 1/2/5 rounding of span/8 (kept for horizontal axes, where PowerPoint places
 *  sparser ticks than the vertical-axis rule below — measured: hbar range 8.6 → unit 2). */
function ppUnitLegacy(x: number): number {
  const exp = Math.floor(Math.log10(Math.max(x, 1e-12)))
  const f = x / 10 ** exp
  const nf = f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10
  return nf * 10 ** exp
}

/** Excel/PowerPoint major tick unit from the axis range (vertical value axes):
 *  p = 10^floor(log10(range)); range/p >= 5 → p, >= 2 → p/2, else p/5
 *  (5-10 intervals; e.g. range 89 → 10, 600 → 100; measured against PowerPoint). */
function ppUnit(range: number): number {
  const p = 10 ** Math.floor(Math.log10(Math.max(range, 1e-12)))
  const ratio = range / p
  return ratio >= 5 ? p : ratio >= 2 ? p / 2 : p / 5
}

/** Next unit up the 1/2/5x10^k ladder. */
function ppUnitUp(s: number): number {
  const exp = Math.floor(Math.log10(Math.max(s, 1e-12)))
  const f = s / 10 ** exp
  return (f < 1.5 ? 2 : f < 3.5 ? 5 : 10) * 10 ** exp
}

function round12(v: number): number {
  return Math.round(v * 1e12) / 1e12
}

/** Compose a data label: optional series/category name parts plus the value (c:showSerName/showCatName). */
function composeDataLabel(
  model: ChartModel,
  si: number,
  catIdx: number,
  valueText: string,
): string {
  const parts: string[] = []
  if (model.dataLabelSerName) parts.push(model.series[si]?.name ?? '')
  if (model.dataLabelCatName) parts.push(model.categories?.[catIdx] ?? '')
  if (!model.dataLabelNoValue) parts.push(valueText)
  return parts.filter(Boolean).join(', ')
}

/** Tick number formatting: integers get thousands separators, decimals keep needed digits. */
/** Data-label number format: the common Excel numeric codes ('0', '#,##0', '0.0', '0%' …);
 *  anything fancier falls back to the plain short form. */
function fmtDataLabel(v: number, code: string | undefined): string {
  if (!code) return fmtNum(round12(v))
  const m = /^([#,0]*0)(?:\.(0+))?(%?)(?:;.*)?$/.exec(code.replace(/"[^"]*"/g, '').trim())
  if (!m) return fmtNum(round12(v))
  const dec = m[2]?.length ?? 0
  const pct = m[3] === '%'
  const scaled = pct ? v * 100 : v
  const text = m[1]!.includes(',')
    ? Number(scaled.toFixed(dec)).toLocaleString('en-US', {
        minimumFractionDigits: dec,
        maximumFractionDigits: dec,
      })
    : scaled.toFixed(dec)
  return pct ? `${text}%` : text
}

function fmtNum(v: number): string {
  if (Number.isInteger(v)) return v.toLocaleString('en-US')
  return String(v)
}
