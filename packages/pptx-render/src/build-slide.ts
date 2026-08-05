/**
 * Slide → RenderSlide builder — turns the pptx-engine Slide element tree into a
 * RenderTree with pixel geometry, resolved styles, and laid-out text (core of approach A).
 */
import type {
  Slide,
  SlideDeck,
  SlideElement,
  TextElement,
  PictureElement,
  GroupElement,
  TableElement,
  ChartElement,
  PassthroughElement,
  ArrowEnd,
} from '@genoffice/pptx-engine'
// Subpath import: the renderer bundles this package, so pulling the engine's index
// (Node-only imports like node:crypto) would break the browser build
import { tableRowGridCols } from '@genoffice/pptx-engine/table-grid'
import { isBackgroundLikeElement } from '@genoffice/pptx-engine/background-promote'
import { buildChartNode } from './build-chart'
import type {
  RenderSlide,
  RenderNode,
  ShapeRenderNode,
  PictureRenderNode,
  GroupRenderNode,
  TableRenderNode,
  TableCellRender,
  ChipRenderNode,
  ArrowEndRender,
} from './render-tree'
import {
  makeViewport,
  placeTransform,
  emuToPx,
  type Viewport,
  type PlacedBox,
  type ParentPlacement,
} from './coords'
import { resolveFill, resolveStroke, resolveShadow, resolveGlow, type MediaResolver } from './fill'
import { layoutText } from './text-layout'
import { HeuristicMetrics, type FontMetricsProvider } from './metrics'
import {
  isConnectorPreset,
  isPillPreset,
  connectorPoints,
  connectorBezier,
  presetPolygon,
  presetPath,
} from './preset-geometry'

export interface BuildOptions {
  /** Target canvas width (px); height follows the slide aspect ratio. */
  fitWidthPx: number
  /** Font metrics provider (heuristic by default; the frontend injects exact opentype metrics). */
  metrics?: FontMetricsProvider
  /** Image/fill mediaRef → dataUrl. */
  media?: MediaResolver
  /**
   * Current slide number (1-based): slidenum fields display it (model bytes untouched).
   * If omitted, the cached text in <a:fld> is used (which may be stale from before slides were moved).
   */
  slideNo?: number
}

/** Refresh slidenum field display to the current slide number (returns a replaced copy; returns the input as-is on no match, never mutates the model). */
function withSlideNum(el: SlideElement, num: number): SlideElement {
  if (el.type === 'group') {
    const g = el as GroupElement
    const children = g.children.map((c) => withSlideNum(c, num))
    return children.some((c, i) => c !== g.children[i]) ? { ...g, children } : el
  }
  const text = (el as TextElement).text
  const hit = text?.paragraphs.some((p) =>
    p.runs.some((r) => r.field === 'slidenum' && r.text !== String(num)),
  )
  if (!hit) return el
  return {
    ...el,
    text: {
      ...text!,
      paragraphs: text!.paragraphs.map((p) => ({
        ...p,
        runs: p.runs.map((r) => (r.field === 'slidenum' ? { ...r, text: String(num) } : r)),
      })),
    },
  } as SlideElement
}

const CHIP_LABEL: Record<string, string> = {
  chart: 'Chart',
  table: 'Table',
  smartart: 'SmartArt',
  ole: 'Embedded object',
  connector: 'Connector',
  media: 'Media',
  unknown: 'Unsupported element',
}

export function buildRenderSlide(
  slide: Slide,
  size: SlideDeck['size'],
  opts: BuildOptions,
): RenderSlide {
  const vp = makeViewport(size, opts.fitWidthPx)
  const metrics = opts.metrics ?? new HeuristicMetrics()
  const sub =
    opts.slideNo != null
      ? (el: SlideElement) => withSlideNum(el, opts.slideNo!)
      : (el: SlideElement) => el
  const nodes: RenderNode[] = []
  // master/layout decoration layer: drawn at the bottom (above the background, below content), read-only
  for (const el of slide.decorations ?? []) {
    const node = buildNode(sub(el), vp, metrics, opts.media, { x: 0, y: 0 })
    if (node && node.type !== 'placeholder-chip') {
      node.decoration = true
      nodes.push(node)
    }
  }
  // Bottom-of-z-order contiguous full-page backgrounds (cloud-generated decks encode
  // the page background as ordinary shapes): marked so the interaction layer can
  // let marquee drags start on them without making them uneditable.
  let bgLeading = 0
  while (
    bgLeading < slide.elements.length &&
    isBackgroundLikeElement(slide.elements[bgLeading]!, size)
  ) {
    bgLeading += 1
  }
  for (const [i, el] of slide.elements.entries()) {
    const node = buildNode(sub(el), vp, metrics, opts.media, { x: 0, y: 0 })
    if (node) {
      if (i < bgLeading) node.background = true
      nodes.push(node)
    }
  }
  // Hidden-slide flag (<p:sld show="0">). Same logic as the engine's readSlideHiddenXml,
  // inlined here as a small regex so we only depend on pptx-engine types (safe to bundle in the renderer).
  const sldOpen = /<p:sld\b[^>]*>/.exec(slide.bodyPrefix)?.[0]
  const hidden = !!sldOpen && /\sshow="0"/.test(sldOpen)
  return {
    widthPx: vp.widthPx,
    heightPx: vp.heightPx,
    scale: vp.scale,
    background: resolveFill(slide.background, vp, opts.media),
    nodes,
    ...(hidden ? { hidden: true } : {}),
  }
}

function buildNode(
  el: SlideElement,
  vp: Viewport,
  metrics: FontMetricsProvider,
  media: MediaResolver | undefined,
  parentOffset: ParentPlacement,
): RenderNode | null {
  const box = placeTransform(el.transform, vp, parentOffset)
  switch (el.type) {
    case 'text':
    case 'shape':
      return buildShape(el as TextElement, box, vp, metrics, media)
    case 'picture':
      return buildPicture(el as PictureElement, box, vp, media)
    case 'group':
      return buildGroup(el as GroupElement, box, vp, metrics, media)
    case 'table':
      return buildTable(el as TableElement, box, vp, metrics, media)
    case 'chart': {
      const chartEl = el as ChartElement
      const appCreated = chartEl.descr === 'aislides-chart'
      // Unsupported chart types fall back to a placeholder chip
      const node =
        buildChartNode(`r_${el.id}`, el.id, chartEl.chart, box, vp, metrics) ??
        chipNode(el.id, box, 'chart', CHIP_LABEL['chart']!)
      if (appCreated && node.type === 'chart') {
        ;(node as import('./render-tree').ChartRenderNode).appCreated = true
      }
      if (node.type === 'chart') {
        ;(node as import('./render-tree').ChartRenderNode).styleInfo = chartStyleInfo(chartEl.chart)
      }
      return node
    }
    case 'passthrough': {
      const pt = el as PassthroughElement
      // SmartArt read-only preview: pre-rendered drawing shapes rendered with group semantics
      // (diagram canvas coordinate size = frame ext, equivalent to chOff 0,0 / chExt ext)
      if (pt.previewShapes?.length) {
        const pseudo: GroupElement = {
          id: pt.id,
          type: 'group',
          anchor: pt.anchor,
          transform: pt.transform,
          children: pt.previewShapes,
          childOffset: { x: 0, y: 0, cx: pt.transform.offset.cx, cy: pt.transform.offset.cy },
        }
        return buildGroup(pseudo, box, vp, metrics, media)
      }
      // OLE read-only preview: the embedded preview image stretches to fill the frame (ignoring the pic's own xfrm)
      if (pt.previewPicture) {
        return buildPicture({ ...pt.previewPicture, id: pt.id }, box, vp, media)
      }
      return buildChip(pt, box)
    }
    default:
      return null
  }
}

/**
 * ArrowEnd (OOXML Stroke model) → ArrowEndRender (px sizes).
 * PowerPoint arrow sizing rules (ECMA-376 §20.1.8.27):
 *   sm / med / lg map to approximate multiples of line width (empirically):
 *   width:  sm≈2×lw, med≈3×lw, lg≈5×lw
 *   length: sm≈2×lw, med≈3×lw, lg≈5×lw
 * A 4px minimum keeps arrows visible.
 */
function resolveArrowEnd(end: ArrowEnd, strokeWidthEmu: number, scale: number): ArrowEndRender {
  const lw = Math.max(emuToPx(strokeWidthEmu, scale), 1)
  const sizeMultW = end.w === 'sm' ? 2 : end.w === 'lg' ? 5 : 3
  const sizeMultL = end.len === 'sm' ? 2 : end.len === 'lg' ? 5 : 3
  return {
    type: end.type as ArrowEndRender['type'],
    widthPx: Math.max(lw * sizeMultW, 4),
    lengthPx: Math.max(lw * sizeMultL, 4),
  }
}

/**
 * Normalized path (0..1) → local px: numeric tokens alternate x/y, multiplied by w/h
 * (M/L/C/Q all have an even coordinate count, so the alternation always holds).
 */
function scaleUnitPath(d: string, w: number, h: number): string {
  let i = 0
  return d
    .split(' ')
    .map((tok) => {
      const n = Number(tok)
      if (!Number.isFinite(n)) return tok
      return String(Math.round(n * (i++ % 2 === 0 ? w : h) * 100) / 100)
    })
    .join(' ')
}

function buildShape(
  el: TextElement,
  box: PlacedBox,
  vp: Viewport,
  metrics: FontMetricsProvider,
  media: MediaResolver | undefined,
): ShapeRenderNode {
  const node: ShapeRenderNode = {
    id: `r_${el.id}`,
    type: el.type,
    box,
    sourceId: el.id,
    fill: resolveFill(el.fill, vp, media),
    ...(el.placeholder ? { placeholder: el.placeholder } : {}),
    ...(el.presetGeometry ? { presetGeometry: el.presetGeometry } : {}),
  }
  // custGeom: the normalized path is scaled to local px by the element box (mutually exclusive with preset, takes priority)
  if (el.customGeometry) {
    const g = el.customGeometry
    if (g.path) node.pathData = scaleUnitPath(g.path, box.w, box.h)
    if (g.fillPath) node.fillPathData = scaleUnitPath(g.fillPath, box.w, box.h)
    if (g.strokePath) node.strokePathData = scaleUnitPath(g.strokePath, box.w, box.h)
  } else if (el.presetGeometry === 'roundRect') {
    // roundRect corner radius: avLst adj (1/1000 %, default 16667≈16.7%) × short side; 50000 = pill
    const adj = el.adjust?.adj ?? 16667
    node.cornerRadiusPx = (Math.min(box.w, box.h) * Math.min(Math.max(adj, 0), 50000)) / 100000
  } else if (isPillPreset(el.presetGeometry)) {
    node.cornerRadiusPx = Math.min(box.w, box.h) / 2
  } else if (isConnectorPreset(el.presetGeometry)) {
    // Connectors: after baking flips into the points, clear the box flip flags (the render container no longer mirrors)
    const pts = connectorPoints(el.presetGeometry!, box.w, box.h, box.flipH, box.flipV, el.adjust)
    const isCurved = /^curvedConnector/.test(el.presetGeometry ?? '')
    const bezier = isCurved ? connectorBezier(pts) : undefined
    const strokeWidth = el.stroke?.width ?? 12700
    node.line = {
      points: pts,
      ...(bezier?.length ? { bezier } : {}),
      ...(el.stroke?.headEnd
        ? { headEnd: resolveArrowEnd(el.stroke.headEnd, strokeWidth, vp.scale) }
        : {}),
      ...(el.stroke?.tailEnd
        ? { tailEnd: resolveArrowEnd(el.stroke.tailEnd, strokeWidth, vp.scale) }
        : {}),
    }
    node.box = { ...box, flipH: false, flipV: false }
  } else if (el.presetGeometry && el.presetGeometry !== 'rect') {
    const poly = presetPolygon(el.presetGeometry, box.w, box.h, el.adjust)
    if (poly) node.polygonPoints = poly
    else {
      const p = presetPath(el.presetGeometry, box.w, box.h, el.adjust)
      if (p) {
        if (p.path) node.pathData = p.path
        if (p.fillPath) node.fillPathData = p.fillPath
        if (p.strokePath) node.strokePathData = p.strokePath
      }
    }
  }
  const stroke = resolveStroke(el.stroke, vp)
  if (stroke) node.stroke = stroke
  const shadow = resolveShadow(el.shadow, vp)
  if (shadow) node.shadow = shadow
  const glow = resolveGlow(el.glow, vp)
  if (glow) node.glow = glow
  if (el.text && el.text.paragraphs.length) {
    node.text = layoutText({
      body: el.text,
      boxWidthPx: box.w,
      boxHeightPx: box.h,
      metrics,
      vp,
    })
  }
  return node
}

/** Picture shape geometry -> three clip channels (same preset implementation as shape geometry). */
function pictureClip(
  preset: string | undefined,
  box: PlacedBox,
  adjust?: Record<string, number>,
): PictureRenderNode['clip'] {
  if (!preset || preset === 'rect') return undefined
  if (preset === 'roundRect') {
    const adj = adjust?.adj ?? 16667
    return { cornerRadiusPx: (Math.min(box.w, box.h) * Math.min(Math.max(adj, 0), 50000)) / 100000 }
  }
  if (isPillPreset(preset)) return { cornerRadiusPx: Math.min(box.w, box.h) / 2 }
  if (preset === 'ellipse') {
    // ellipse goes through the Konva primitive channel in the render layer and produces no path; clip with an SVG two-arc path
    const rx = box.w / 2
    const ry = box.h / 2
    return {
      pathData: `M 0 ${ry} A ${rx} ${ry} 0 1 0 ${box.w} ${ry} A ${rx} ${ry} 0 1 0 0 ${ry} Z`,
    }
  }
  const poly = presetPolygon(preset, box.w, box.h, adjust)
  if (poly) return { polygonPoints: poly }
  const p = presetPath(preset, box.w, box.h, adjust)
  if (p?.path) return { pathData: p.path }
  return undefined
}

function buildPicture(
  el: PictureElement,
  box: PlacedBox,
  vp: Viewport,
  media: MediaResolver | undefined,
): PictureRenderNode {
  const dataUrl = el.dataUrl ?? (el.mediaRef ? media?.(el.mediaRef) : undefined)
  const clip = pictureClip(el.presetGeometry, box, el.adjust)
  const node: PictureRenderNode = {
    id: `r_${el.id}`,
    type: 'picture',
    box,
    sourceId: el.id,
    ...(dataUrl ? { dataUrl } : {}),
    ...(clip ? { clip } : {}),
    ...(el.srcRect ? { srcRect: el.srcRect } : {}),
    ...(el.opacity != null ? { opacity: el.opacity } : {}),
    ...(el.softEdge ? { softEdgePx: emuToPx(el.softEdge, vp.scale) } : {}),
    ...(el.media ? { media: el.media.kind } : {}),
    ...(el.name ? { name: el.name } : {}),
    ...(el.descr ? { descr: el.descr } : {}),
  }
  const stroke = resolveStroke(el.stroke, vp)
  if (stroke) node.stroke = stroke
  const shadow = resolveShadow(el.shadow, vp)
  if (shadow) node.shadow = shadow
  const glow = resolveGlow(el.glow, vp)
  if (glow) node.glow = glow
  return node
}

function buildGroup(
  el: GroupElement,
  box: PlacedBox,
  vp: Viewport,
  metrics: FontMetricsProvider,
  media: MediaResolver | undefined,
): GroupRenderNode {
  // Child coordinate system: child element EMU coords are based on <a:chOff>/<a:chExt>.
  // The size difference between the group box and the child coordinate system (the group
  // was scaled as a whole; ext/chExt may be non-uniform) is **baked into child geometry**
  // at placement time (position/size converted to group-local px by scale), while font
  // size and stroke width keep their declared values unscaled -- group scaling only
  // affects geometry, and text lays out in the scaled box at the
  // original font size (MS-OE376 §2.1.1360/2.1.1364). The render container no longer
  // stretches as a whole (which would make line-wrap widths wrong and distort glyphs
  // non-uniformly). childScaleX/Y are still written on the node so the edit pipeline can
  // convert between group-local px and child EMU.
  const ch = el.childOffset
  const chX = ch?.x ?? el.transform.offset.x
  const chY = ch?.y ?? el.transform.offset.y
  const chCx = ch?.cx || el.transform.offset.cx
  const chCy = ch?.cy || el.transform.offset.cy
  const chWpx = emuToPx(chCx, vp.scale)
  const chHpx = emuToPx(chCy, vp.scale)
  const childScaleX = chWpx > 0 ? box.w / chWpx : 1
  const childScaleY = chHpx > 0 ? box.h / chHpx : 1
  const parentOffset: ParentPlacement = {
    x: -emuToPx(chX, vp.scale) * childScaleX,
    y: -emuToPx(chY, vp.scale) * childScaleY,
    scaleX: childScaleX,
    scaleY: childScaleY,
  }

  const children: RenderNode[] = []
  for (const child of el.children ?? []) {
    const c = buildNode(child, vp, metrics, media, parentOffset)
    if (c) children.push(c)
  }
  return {
    id: `r_${el.id}`,
    type: 'group',
    box,
    sourceId: el.id,
    children,
    ...(Math.abs(childScaleX - 1) > 1e-6 ? { childScaleX } : {}),
    ...(Math.abs(childScaleY - 1) > 1e-6 ? { childScaleY } : {}),
  }
}

/**
 * Table → TableRenderNode. Cell coordinates are relative to the table's top-left (px).
 * Column widths/row heights are distributed by EMU ratio over the frame's actual px size.
 * Row height is a minimum in PPT: rows whose wrapped cell text needs more space grow to
 * fit it, and the node box grows with the total so nothing is clipped.
 */
function buildTable(
  el: TableElement,
  box: PlacedBox,
  vp: Viewport,
  metrics: FontMetricsProvider,
  media: MediaResolver | undefined,
): TableRenderNode {
  const sumW = el.colWidths.reduce((a, b) => a + b, 0) || 1
  const sumH = el.rowHeights.reduce((a, b) => a + b, 0) || 1
  const colPx = el.colWidths.map((w) => (w / sumW) * box.w)
  const rowPx = el.rowHeights.map((h) => (h / sumH) * box.h)
  // Prefix sums → start of each column
  const colX: number[] = [0]
  for (const w of colPx) colX.push(colX[colX.length - 1]! + w)

  // Measure pass: grow any row whose cell content wraps taller than the stored height
  // (rowSpan cells are skipped — their height is ambiguous to attribute to one row).
  el.rows.forEach((row, r) => {
    const gridCols = tableRowGridCols(row)
    row.forEach((cell, tcIdx) => {
      const cIdx = gridCols[tcIdx]!
      if (cell.merged || (cell.rowSpan ?? 1) > 1) return
      if (!cell.text || !cell.text.paragraphs.length) return
      const x = colX[cIdx] ?? 0
      const w = (colX[Math.min(cIdx + (cell.gridSpan ?? 1), colX.length - 1)] ?? x) - x
      const probe = layoutText({
        body: cell.text,
        boxWidthPx: w,
        boxHeightPx: rowPx[r] ?? 0,
        metrics,
        vp,
      })
      const needed = probe.contentHeight + probe.insets.t + probe.insets.b
      if (needed > (rowPx[r] ?? 0)) rowPx[r] = needed
    })
  })

  const rowY: number[] = [0]
  for (const h of rowPx) rowY.push(rowY[rowY.length - 1]! + h)
  const totalH = rowY[rowY.length - 1]!
  if (totalH > box.h + 0.5) box = { ...box, h: totalH }

  const cells: TableCellRender[] = []
  el.rows.forEach((row, r) => {
    const gridCols = tableRowGridCols(row)
    row.forEach((cell, tcIdx) => {
      const cIdx = gridCols[tcIdx]!
      const gridSpan = cell.gridSpan ?? 1
      if (cell.merged) return
      const rowSpan = cell.rowSpan ?? 1
      const x = colX[cIdx] ?? 0
      const y = rowY[r] ?? 0
      const w = (colX[Math.min(cIdx + gridSpan, colX.length - 1)] ?? x) - x
      const h = (rowY[Math.min(r + rowSpan, rowY.length - 1)] ?? y) - y
      const out: TableCellRender = {
        x,
        y,
        w,
        h,
        row: r,
        col: tcIdx,
        ...(gridSpan > 1 ? { gridSpan } : {}),
        ...(rowSpan > 1 ? { rowSpan } : {}),
        fill: resolveFill(cell.fill, vp, media),
      }
      const borders: NonNullable<TableCellRender['borders']> = {}
      for (const k of ['l', 'r', 't', 'b'] as const) {
        const s = resolveStroke(cell.borders?.[k], vp)
        if (s) borders[k] = s
      }
      if (Object.keys(borders).length) out.borders = borders
      if (cell.text && cell.text.paragraphs.length) {
        out.text = layoutText({
          body: cell.text,
          boxWidthPx: w,
          boxHeightPx: h,
          metrics,
          vp,
        })
      }
      cells.push(out)
    })
  })

  return {
    id: `r_${el.id}`,
    type: 'table',
    box,
    sourceId: el.id,
    cells,
    gridX: colX,
    gridY: rowY,
    ...(el.styleFlags ? { styleFlags: el.styleFlags } : {}),
  }
}

/** ChartModel → style summary echoed to the Ribbon (kind maps to EditChartOp semantics; horizontal bars count as bar). */
function chartStyleInfo(m: ChartElement['chart']): import('./render-tree').ChartStyleInfo {
  const kind =
    m.kind === 'bar'
      ? m.series.some((s) => s.plotKind === 'line')
        ? 'comboBarLine'
        : m.grouping === 'stacked' || m.grouping === 'percentStacked'
          ? 'barStacked'
          : 'bar'
      : m.kind === 'pie'
        ? (m.holePct ?? 0) > 0
          ? 'doughnut'
          : 'pie'
        : m.kind
  return {
    kind,
    legendPos: m.legendPos == null ? 'none' : m.legendPos === 'tr' ? 'r' : m.legendPos,
    dataLabels: !!m.dataLabels,
    gridlines: !!m.valAxis?.gridColor,
    ...(m.title ? { title: m.title } : {}),
    ...(m.catAxis?.title ? { catAxisTitle: m.catAxis.title } : {}),
    ...(m.valAxis?.title ? { valAxisTitle: m.valAxis.title } : {}),
    ...(m.gapWidthPct != null ? { gapWidthPct: m.gapWidthPct } : {}),
  }
}

function buildChip(el: PassthroughElement, box: PlacedBox): ChipRenderNode {
  return chipNode(el.id, box, el.kind, CHIP_LABEL[el.kind] ?? el.kind)
}

function chipNode(sourceId: string, box: PlacedBox, kind: string, label: string): ChipRenderNode {
  return {
    id: `r_${sourceId}`,
    type: 'placeholder-chip',
    box,
    sourceId,
    kind,
    label,
  }
}
