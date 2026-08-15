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
import { ptToPx, type Viewport } from './coords'
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
  // centered on top. PowerPoint sizes the auto title at 1.2× the chart's default text.
  const titleSizePx = ptToPx(chartTextPt(model) * 1.2, vp.scale)
  const titleH = titleSizePx * 1.7
  const node = buildChartNodeInner(
    id,
    sourceId,
    model,
    { ...box, h: Math.max(box.h - titleH, 10) },
    vp,
    metrics,
  )
  if (!node) return null
  shiftChartNode(node, titleH)
  extrudeBars(node, model)
  node.box = box
  const tw = metrics.measure(model.title, {
    fontFamily: LABEL_FONT,
    fontSizePx: titleSizePx,
    bold: true,
    italic: false,
  })
  node.labels.push({
    text: model.title,
    x: Math.max((box.w - tw) / 2, 4),
    y: titleSizePx * 0.3,
    fontSizePx: titleSizePx,
    color: '#333333',
    bold: true,
  })
  return node
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
  if (model.kind === 'bar' && model.barDir === 'bar') {
    return buildHBarNode(id, sourceId, model, box, vp, metrics)
  }
  if (model.kind !== 'line' && model.kind !== 'bar' && model.kind !== 'area') return null

  const grouping = model.kind === 'bar' ? (model.grouping ?? 'clustered') : 'standard'
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
  const seriesColor = (i: number) => model.series[i]?.color ?? palette[i % palette.length]!

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
    const posSums = Array.from({ length: catCount }, (_, i) =>
      barSeriesIdx.reduce((a, si) => a + Math.max(valueAt(si, i) ?? 0, 0), 0),
    )
    const negSums = Array.from({ length: catCount }, (_, i) =>
      barSeriesIdx.reduce((a, si) => a + Math.min(valueAt(si, i) ?? 0, 0), 0),
    )
    // Overlaid line/area series on the secondary axis don't feed into the primary range
    const overlayVals = model.series
      .filter((s) => serKind(s) !== 'bar' && !onSecAxis(s))
      .flatMap(numVals)
    dataMax = Math.max(...posSums, ...overlayVals, 0)
    dataMin = Math.min(...negSums, ...overlayVals, 0)
  } else {
    dataMax = Math.max(...priVals, 0)
    dataMin = Math.min(...priVals, 0)
  }
  const { min, max, ticks } = ppTicks(
    model.valAxis?.min ?? dataMin,
    model.valAxis?.max ?? dataMax,
    model.valAxis?.min == null,
    model.valAxis?.max == null,
  )
  // Secondary value axis ticks (right side): range fully independent of the primary (e.g. left axis revenue 0-250, right axis growth% 0-30)
  const sec = secVals.length
    ? ppTicks(
        model.valAxis2?.min ?? Math.min(...secVals, 0),
        model.valAxis2?.max ?? Math.max(...secVals, 0),
        model.valAxis2?.min == null,
        model.valAxis2?.max == null,
      )
    : undefined

  // ── Layout frame ────────────────────────────────────────────────
  const pad = Math.max(4, box.w * 0.01)
  const legendPos = model.legendPos
  const legendH = legendPos === 't' || legendPos === 'b' ? labelSizePx * 1.6 : 0
  const tickLabels = ticks.map((t) => fmtNum(t))
  const yLabelW = Math.max(...tickLabels.map((t) => measure(t, labelSizePx)), 0)
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
  const plotR = box.w - pad - (sec ? y2LabelW + axisTitle2W + 10 : labelSizePx * 0.7)
  const plotB = box.h - pad - catLabelSizePx * 1.5 - (legendPos === 'b' ? legendH : 0)
  const plot = {
    x: plotX,
    y: plotY,
    w: Math.max(plotR - plotX, 10),
    h: Math.max(plotB - plotY, 10),
  }

  const yOf = (v: number) => plot.y + plot.h * (1 - (v - min) / (max - min || 1))
  // Secondary axis mapping: same plot-area height, independent range
  const yOf2 = (v: number) =>
    plot.y + plot.h * (1 - (v - (sec?.min ?? 0)) / ((sec?.max ?? 1) - (sec?.min ?? 0) || 1))

  // ── Gridlines + tick labels + axes ──────────────────────────────
  const gridColor = model.valAxis?.gridColor
  for (let i = 0; i < ticks.length; i++) {
    const t = ticks[i]!
    const y = yOf(t)
    if (gridColor && t !== min) {
      node.gridLines.push({
        x1: plot.x,
        y1: y,
        x2: plot.x + plot.w,
        y2: y,
        color: gridColor,
        ...(model.valAxis?.gridDash ? { dash: [4, 4] } : {}),
      })
    }
    const text = tickLabels[i]!
    node.labels.push({
      text,
      x: plot.x - 6 - measure(text, labelSizePx),
      y: y - labelSizePx * 0.55,
      fontSizePx: labelSizePx,
      color: labelColor,
    })
  }
  const axisColor = model.valAxis?.lineColor ?? model.catAxis?.lineColor ?? '#888888'
  const axisW = Math.max(1, ptToPx(1, vp.scale))
  // x axis (bottom) + y axis (left)
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

  // ── Category labels (PowerPoint shows all if possible, thinning only when they truly don't fit) ──
  const n = Math.max(model.categories.length, 1)
  const slotW = plot.w / n
  const maxCatW = Math.max(...model.categories.map((c) => measure(c, catLabelSizePx)), 1)
  // Slight overflow (<30%) isn't thinned — PowerPoint prefers showing all (tolerating crowding/slight overlap)
  const skip = maxCatW <= slotW * 1.3 ? 1 : Math.max(1, Math.ceil(maxCatW / slotW))
  model.categories.forEach((cat, i) => {
    if (i % skip !== 0 && i !== n - 1) return
    const cx = plot.x + (i + 0.5) * slotW
    node.labels.push({
      text: cat,
      x: cx - measure(cat, catLabelSizePx) / 2,
      y: plot.y + plot.h + catLabelSizePx * 0.35,
      fontSizePx: catLabelSizePx,
      color: catLabelColor,
    })
  })

  // ── Data series (bar series + line/area series can coexist: combo charts) ──
  const dlSize = labelSizePx * 0.9
  // Data labels (c:dLbls showVal): white centered inside stacked segments, dark gray outside points/bars otherwise
  const dLbl = (cx: number, y: number, v: number, inside: boolean) => {
    if (!model.dataLabels) return
    const text = fmtNum(round12(v))
    node.labels.push({
      text,
      x: cx - measure(text, dlSize) / 2,
      y,
      fontSizePx: dlSize,
      color: inside ? '#FFFFFF' : '#404040',
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
        dLbl(x + barW / 2, (yTop + yBot) / 2 - dlSize * 0.55, ser.values[i]!, true)
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
        dLbl(x + barW / 2, v >= 0 ? yTop - dlSize * 1.15 : yBot + dlSize * 0.15, v, false)
      })
    })
  }
  {
    const lineW = Math.max(1.5, ptToPx(1.5, vp.scale))
    const markerR = Math.max(2, ptToPx(3, vp.scale))
    model.series.forEach((ser, si) => {
      const k = serKind(ser)
      if (k === 'bar') return
      // Dual axis: series on the secondary axis map y via the right-axis range
      const secSer = onSecAxis(ser)
      const yOfSer = secSer ? yOf2 : yOf
      const color = seriesColor(si)
      const pts: number[] = []
      ser.values.forEach((v, i) => {
        if (v == null || i >= n) return
        const x = plot.x + (i + 0.5) * slotW
        const y = yOfSer(v)
        pts.push(x, y)
        if (k === 'line' && ser.marker) node.markers.push({ x, y, r: markerR, color })
        dLbl(x, y - dlSize * 1.3, v, false)
      })
      if (pts.length < 4) return
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
        node.polylines.push({
          points: pts,
          color,
          widthPx: lineW,
          ...(ser.smooth ? { smooth: true } : {}),
        })
      }
    })
  }

  // ── Legend ──────────────────────────────────────────────────────
  if (legendPos && model.series.some((s) => s.name)) {
    const sw = labelSizePx * 1.1
    const items = model.series.map((s, i) => ({
      label: s.name ?? '',
      color: seriesColor(i),
    }))
    const itemWs = items.map((it) => sw + 4 + measure(it.label, labelSizePx) + labelSizePx)
    if (legendPos === 't' || legendPos === 'b') {
      const total = itemWs.reduce((a, b) => a + b, 0)
      let x = Math.max((box.w - total) / 2, pad)
      const y = legendPos === 't' ? pad : box.h - pad - labelSizePx * 1.2
      items.forEach((it, i) => {
        node.swatches.push({
          x,
          y: y + labelSizePx * 0.25,
          w: sw,
          h: labelSizePx * 0.6,
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
      // Right side (l/r/tr all laid out as a top-right column); with a secondary axis, shift right to clear its tick labels
      let y = plot.y
      const x = plot.x + plot.w + 8 + (sec ? y2LabelW + axisTitle2W + 8 : 0)
      items.forEach((it) => {
        node.swatches.push({
          x,
          y: y + labelSizePx * 0.25,
          w: sw,
          h: labelSizePx * 0.6,
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
  const sliceColor = (i: number) => ser.pointColors?.[i] ?? palette[i % palette.length]!
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

  const outerR = Math.max(Math.min(plotW, plotH) / 2, 5)
  const cx = plotX + plotW / 2
  const cy = plotY + plotH / 2
  const innerR = (outerR * Math.min(Math.max(model.holePct ?? 0, 0), 90)) / 100

  // Pseudo-3D pie: the tilted disc projects to an ellipse (ry = sin(rotX) · rx) with a darker
  // side rim of depth below the front-facing arcs; wedges become elliptical sector paths.
  const p3d = !!model.pseudo3D && innerR === 0
  const kY = p3d
    ? Math.min(Math.max(Math.sin(((model.rotXDeg ?? 30) * Math.PI) / 180), 0.35), 0.9)
    : 1
  const rx = p3d ? Math.max(Math.min(plotW / 2, plotH / 2 / (kY + 0.3)), 5) : outerR
  const ry = rx * kY
  const depth = p3d ? ry * 0.3 : 0
  const ptAt = (deg: number) => {
    const t = (deg * Math.PI) / 180
    return { x: cx + Math.cos(t) * rx, y: cy + Math.sin(t) * ry }
  }
  if (p3d) {
    node.paths = node.paths ?? []
    // Rim: the front half is the parametric range [0°, 180°) (screen lower half)
    let a = -90 + (model.firstSliceAngDeg ?? 0)
    vals.forEach((v, i) => {
      if (v <= 0) return
      const sweep = (v / total) * 360
      // normalize wedge interval into [-180, 180) then clamp to the front range [0, 180]
      for (const off of [-360, 0, 360]) {
        const b1 = Math.max(a + off, 0)
        const b2 = Math.min(a + off + sweep, 180)
        if (b2 <= b1) continue
        const p1 = ptAt(b1)
        const p2 = ptAt(b2)
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
    if (p3d) {
      const p1 = ptAt(angle)
      if (sweep >= 359.999) {
        // Full circle: an SVG arc with coincident endpoints renders nothing, so use two half arcs
        const pm = ptAt(angle + 180)
        node.paths!.push({
          d:
            `M ${p1.x} ${p1.y} A ${rx} ${ry} 0 1 1 ${pm.x} ${pm.y} ` +
            `A ${rx} ${ry} 0 1 1 ${p1.x} ${p1.y} Z`,
          fill: sliceColor(i),
          stroke: '#ffffff',
        })
      } else {
        const p2 = ptAt(angle + sweep)
        const large = sweep > 180 ? 1 : 0
        node.paths!.push({
          d: `M ${cx} ${cy} L ${p1.x} ${p1.y} A ${rx} ${ry} 0 ${large} 1 ${p2.x} ${p2.y} Z`,
          fill: sliceColor(i),
          stroke: '#ffffff',
        })
      }
    } else {
      node.wedges!.push({
        cx,
        cy,
        outerR,
        innerR,
        startDeg: angle,
        sweepDeg: sweep,
        color: sliceColor(i),
      })
    }
    if (model.dataLabels) {
      // At the wedge midline radius: doughnut uses the ring-band midpoint, pie uses 2/3 radius
      const midRad = ((angle + sweep / 2) * Math.PI) / 180
      const r = innerR > 0 ? (innerR + outerR) / 2 : outerR * 0.66
      const text = model.dataLabelsPct ? `${Math.round((v / total) * 100)}%` : fmtNum(v)
      node.labels.push({
        text,
        x: cx + Math.cos(midRad) * (p3d ? rx * 0.66 : r) - measure(text) / 2,
        y: cy + Math.sin(midRad) * (p3d ? ry * 0.66 : r) - labelSizePx * 0.55,
        fontSizePx: labelSizePx * 0.9,
        color: '#FFFFFF',
      })
    }
    angle += sweep
  })

  // Legend
  if (legendPos) {
    const sw = labelSizePx * 1.1
    if (legendPos === 't' || legendPos === 'b') {
      const itemWs = legendItems.map((it) => sw + 4 + measure(it.label) + labelSizePx)
      const totalW = itemWs.reduce((a, b) => a + b, 0)
      let x = Math.max((box.w - totalW) / 2, pad)
      const y = legendPos === 't' ? pad * 0.5 : box.h - pad * 0.5 - labelSizePx * 1.2
      legendItems.forEach((it, i) => {
        node.swatches.push({
          x,
          y: y + labelSizePx * 0.25,
          w: sw,
          h: labelSizePx * 0.6,
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
          h: labelSizePx * 0.6,
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
  const seriesColor = (i: number) => model.series[i]?.color ?? palette[i % palette.length]!

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
  const { min, max, ticks } = ppTicks(
    model.valAxis?.min ?? dataMin,
    model.valAxis?.max ?? dataMax,
    model.valAxis?.min == null,
    model.valAxis?.max == null,
  )

  // Layout: left category-label width + bottom value-tick row + legend
  const pad = Math.max(4, box.w * 0.01)
  const legendPos = model.legendPos
  const legendH = legendPos === 't' || legendPos === 'b' ? labelSizePx * 1.6 : 0
  const catLabelW = Math.max(...model.categories.map((c) => measure(c, catLabelSizePx)), 0)
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
  const tickLabels = ticks.map((t) => fmtNum(t))
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
  const dLbl = (x: number, yMid: number, v: number, inside: boolean) => {
    if (!model.dataLabels) return
    const text = fmtNum(round12(v))
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
        dLbl((xL + xR) / 2, y + barH / 2, ser.values[i]!, true)
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
        const y = rowY(i) + (slotH - groupH) / 2 + si * barH
        const xL = xOf(Math.min(v, base))
        const xR = xOf(Math.max(v, base))
        node.bars.push({
          x: xL,
          y,
          w: Math.max(xR - xL, 0.5),
          h: barH,
          color: ser.pointColors?.[i] ?? color,
        })
        dLbl(v >= 0 ? xR + 4 : xL - 4 - measure(fmtNum(round12(v)), dlSize), y + barH / 2, v, false)
      })
    })
  }

  addSeriesLegend(node, model, box, plot, labelSizePx, measure, pad, seriesColor)
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
  const seriesColor = (i: number) => model.series[i]?.color ?? palette[i % palette.length]!

  // Point sets: x defaults to indices 1..n
  const points = model.series.map((s) =>
    s.values
      .map((y, i) => ({ x: s.xValues?.[i] ?? i + 1, y }))
      .filter((p): p is { x: number; y: number } => p.x != null && p.y != null),
  )
  const allX = points.flat().map((p) => p.x)
  const allY = points.flat().map((p) => p.y)
  if (!allY.length) return null

  // The catAxis slot = the x axis (the engine dispatched by axPos)
  const xTicksR = ppTicks(
    model.catAxis?.min ?? Math.min(...allX, 0),
    model.catAxis?.max ?? Math.max(...allX, 0),
    model.catAxis?.min == null,
    model.catAxis?.max == null,
  )
  const yTicksR = ppTicks(
    model.valAxis?.min ?? Math.min(...allY, 0),
    model.valAxis?.max ?? Math.max(...allY, 0),
    model.valAxis?.min == null,
    model.valAxis?.max == null,
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
  model.series.forEach((ser, si) => {
    const color = seriesColor(si)
    const pts = points[si]!
    const showMarker = ser.marker ?? defaultMarker
    const flat: number[] = []
    for (const p of pts) {
      const x = xOf(p.x)
      const y = yOf(p.y)
      flat.push(x, y)
      if (showMarker) node.markers.push({ x, y, r: markerR, color })
      if (model.dataLabels) {
        const text = fmtNum(round12(p.y))
        node.labels.push({
          text,
          x: x - measure(text, labelSizePx * 0.9) / 2,
          y: y - labelSizePx * 1.3,
          fontSizePx: labelSizePx * 0.9,
          color: '#404040',
        })
      }
    }
    if (hasLine && flat.length >= 4) {
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
  const seriesColor = (i: number) => model.series[i]?.color ?? palette[i % palette.length]!

  const { min, max, ticks } = ppTicks(
    model.valAxis?.min ?? Math.min(...allVals, 0),
    model.valAxis?.max ?? Math.max(...allVals, 0),
    model.valAxis?.min == null,
    model.valAxis?.max == null,
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
): void {
  const legendPos = model.legendPos
  if (!legendPos || !model.series.some((s) => s.name)) return
  const sw = labelSizePx * 1.1
  const items = model.series.map((s, i) => ({
    label: s.name ?? '',
    color: seriesColor(i),
  }))
  const itemWs = items.map((it) => sw + 4 + measure(it.label, labelSizePx) + labelSizePx)
  const labelColor = model.valAxis?.labelColor ?? '#666666'
  if (legendPos === 't' || legendPos === 'b') {
    const total = itemWs.reduce((a, b) => a + b, 0)
    let x = Math.max((box.w - total) / 2, pad)
    const y = legendPos === 't' ? pad : box.h - pad - labelSizePx * 1.2
    items.forEach((it, i) => {
      node.swatches.push({
        x,
        y: y + labelSizePx * 0.25,
        w: sw,
        h: labelSizePx * 0.6,
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
        y: y + labelSizePx * 0.25,
        w: sw,
        h: labelSizePx * 0.6,
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
): { min: number; max: number; ticks: number[] } {
  let lo = autoMin ? Math.min(rawMin, 0) : rawMin
  let hi = rawMax
  if (hi <= lo) hi = lo + 1
  // 5% headroom (auto ends only)
  const span0 = hi - lo
  const hiT = autoMax && hi > 0 ? hi + span0 * 0.05 : hi
  const loT = autoMin && lo < 0 ? lo - span0 * 0.05 : lo
  const step = ppUnit((hiT - loT) / 8)
  if (autoMin) lo = Math.floor(loT / step) * step
  if (autoMax) hi = Math.ceil(hiT / step) * step
  const ticks: number[] = []
  for (let v = lo; v <= hi + step * 1e-6; v += step) ticks.push(round12(v))
  return { min: lo, max: hi, ticks }
}

/** Excel major tick unit: 1/2/5 × 10^k. */
function ppUnit(x: number): number {
  const exp = Math.floor(Math.log10(Math.max(x, 1e-12)))
  const f = x / 10 ** exp
  const nf = f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10
  return nf * 10 ** exp
}

function round12(v: number): number {
  return Math.round(v * 1e12) / 1e12
}

/** Tick number formatting: integers get thousands separators, decimals keep needed digits. */
function fmtNum(v: number): string {
  if (Number.isInteger(v)) return v.toLocaleString('en-US')
  return String(v)
}
