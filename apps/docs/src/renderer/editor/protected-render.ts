import type { Node as PmNode } from '@tiptap/pm/model'
import {} from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import {} from '@tiptap/pm/tables'
import { WORDART_PRESETS, wordArtStrokePx } from '@genoffice/ui'
import {
  autospaceBoundaries,
  autospacePadBetween,
  cjkDeclaredLineFactor,
  cssAutoLineMult,
  cssCsFontFamily,
  cssFontFamily,
  cssRunFontFamily,
  cssGridLineBase,
  cssGridSpacingPt,
  cssLineHeight,
  isCjk,
  isCjkFontName,
  lineHeightFactor,
  paraLineFactorCss,
  runLetterSpacingCss,
  strutFontCss,
  textHasCjk,
  textHasComplexScript,
  textHasHangul,
  WORD_AUTO_SPACING_PT,
} from '../line-metrics'
import { custGeomBackgroundCss, shapeBackgroundCss, shapeTextInsetsPx } from './shape-svg'
import { t } from '../i18n/locale'
import {
  lumHex,
  type ChartDisplay,
  type FieldDisplay,
  type FormulaDisplay,
  type Run,
  type TableModel,
  type TableParagraph,
  type TextboxDisplay,
} from '@genoffice/docx-engine'

/**
 * Custom schema mirroring the docx-engine Block model 1:1.
 * Every top-level node carries `docxIndex` (patch anchor, null = new) and
 * `aiChanged` (diff highlighting for AI edits).
 */

import {
  DomSpec,
  ProtectedContentEditor,
  TableBordersAttr,
  borderLineCss,
  cellClipStyle,
  cellPadCss,
  cellVAlignGridCss,
  cellWritingMode,
  preventProtectedLineBreak,
  protectedText,
  tableBordersCss,
} from './extensions'
import { cellClipTwips, inferredBidi } from './convert'

// Word: links and TOC entries jump on modifier+click only
const jumpHint = () =>
  navigator.platform.toLowerCase().includes('mac') ? t('editorJumpHintMac') : t('editorJumpHintWin')

/** Rendering of a field result; safe visible text becomes editable on double-click. */
export function renderFieldSpec(field: FieldDisplay): DomSpec | null {
  if (field.kind === 'tocLine') {
    const attrs: Record<string, string> = {
      class: `doc-toc-line doc-toc-l${Math.min(field.level ?? 1, 4)}`,
      'data-toc-title': field.left ?? '',
      title: jumpHint(),
    }
    if (field.anchor) attrs['data-toc-anchor'] = field.anchor
    // direct pPr/run metrics of the entry paragraph beat the inherited body size
    const tocStyles: string[] = []
    if (field.szHalfPoints) tocStyles.push(`font-size:${field.szHalfPoints / 2}pt`)
    const tocLh = cssLineHeight(field.lineRule, field.lineRawTwips, field.lineSpacing)
    if (tocLh) tocStyles.push(`line-height:${tocLh}`)
    // --doc-line-max reads the multiple from this var on the same element
    const tocMult = cssAutoLineMult(field.lineRule, field.lineRawTwips, field.lineSpacing)
    if (tocMult && tocMult !== 1) tocStyles.push(`--doc-line-mult:${tocMult}`)
    if (tocStyles.length > 0) attrs.style = tocStyles.join(';')
    const num: DomSpec[] = field.num
      ? [['span', { class: 'doc-toc-num', contenteditable: 'false' }, field.num]]
      : []
    return [
      'div',
      attrs,
      ...num,
      ['span', { class: 'doc-toc-title', contenteditable: 'false' }, field.left || '\u00a0'],
      // real dot glyphs (clipped to the free width), not a border decoration:
      // Word/LO leader dots are text, and exported-PDF text comparison sees them
      ['span', { class: 'doc-toc-dots', contenteditable: 'false' }, '.'.repeat(220)],
      ['span', { class: 'doc-toc-page', contenteditable: 'false' }, field.right ?? ''],
    ]
  }
  if (field.kind === 'pageBreak') {
    return ['div', { class: 'doc-field-pagebreak' }, ['span', {}, t('editorPageBreak')]]
  }
  if (field.kind === 'text' && field.left) {
    // carry the result runs' face/size: the passthrough div otherwise inherits
    // the document default and mis-snaps on a typed line grid. Spaces must
    // survive inline: prosemirror-view injects a higher-specificity
    // `.ProseMirror [contenteditable=false] { white-space: normal }`
    const styles: string[] = ['white-space:pre-wrap']
    if (field.szHalfPoints) styles.push(`font-size:${field.szHalfPoints / 2}pt`)
    if (field.align) styles.push(`text-align:${field.align}`)
    if (field.fontFamily) {
      styles.push(
        `--doc-line-factor:${lineHeightFactor(field.fontFamily)}`,
        `font-family:${cssFontFamily(field.fontFamily)}`,
      )
    }
    if (field.szHalfPoints || field.fontFamily || field.lineRule) {
      // explicit w:spacing beats the single-spacing snap; either way the line
      // strut must recompute from the field's own size/face, not the
      // wrapper's document-default computed box
      styles.push(
        `line-height:${cssLineHeight(field.lineRule, field.lineRawTwips, field.lineSpacing) ?? cssGridLineBase()}`,
      )
      const mult = cssAutoLineMult(field.lineRule, field.lineRawTwips, field.lineSpacing)
      if (mult && mult !== 1) styles.push(`--doc-line-mult:${mult}`)
    }
    const attrs: Record<string, string> = {
      class: 'doc-field-text',
      contenteditable: 'false',
      style: styles.join(';'),
    }
    // mixed-size result (a manual drop-cap letter before body text): per-run
    // sized spans; the wrapper keeps the dominant size for the line strut
    if (field.runs) {
      return [
        'div',
        attrs,
        ...field.runs.map((r): DomSpec =>
          r.szHalfPoints
            ? ['span', { style: `font-size:${r.szHalfPoints / 2}pt` }, r.text]
            : ['span', {}, r.text],
        ),
      ]
    }
    return ['div', attrs, field.left]
  }
  return null
}

export function renderFormulaSpec(formula: FormulaDisplay): DomSpec {
  const tokenStrip: DomSpec = [
    'span',
    { class: 'doc-formula' + (formula.mathml ? ' doc-formula-has-math' : '') },
    ...formula.tokens.map((token, index): DomSpec => [
      'span',
      {
        class: 'doc-formula-token',
        'data-token-index': String(index),
        contenteditable: 'false',
      },
      token || '\u00a0',
    ]),
  ]
  if (!formula.mathml) return tokenStrip
  // the MathML host is empty in the spec; buildProtectedDom injects the markup
  // (renderSpec cannot emit raw MathML)
  return [
    'span',
    { class: 'doc-formula-wrap' },
    ['span', { class: 'doc-formula-math', contenteditable: 'false' }],
    tokenStrip,
  ]
}

// ---- embedded charts: SVG preview + editable data grid ----

/** Office theme default accent colors, used for new charts / theme-less docs */
const CHART_PALETTE = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']

/** series/point cycle color; repeat rounds darken like Word (accentN lumMod 60%) */
function chartColor(chart: ChartDisplay, i: number): string {
  const palette = chart.palette?.length ? chart.palette : CHART_PALETTE
  const base = palette[i % palette.length]
  const round = Math.floor(i / palette.length)
  return `#${round > 0 ? lumHex(base, 0.6 ** round, 0) : base}`
}

/** explicit c:ser fill beats the style cycle */
function seriesColor(chart: ChartDisplay, s: number): string {
  const explicit = chart.series[s]?.color
  return explicit ? `#${explicit}` : chartColor(chart, s)
}

/** per-point fill (pie slices): c:dPt beats the cycle */
function pointColor(chart: ChartDisplay, s: number, i: number): string {
  const explicit = chart.series[s]?.pointColors?.[i]
  return explicit ? `#${explicit}` : chartColor(chart, i)
}

/**
 * Chart preview + data grid. The grid is a chart data sheet
 * (rows = series, columns = categories); its cells and the title become
 * editable on double-click, everything else stays protected.
 */
export function renderChartSpec(chart: ChartDisplay): DomSpec {
  const children: DomSpec[] = []
  if (chart.title !== undefined) {
    children.push([
      'div',
      { class: 'doc-chart-title', contenteditable: 'false' },
      chart.title || '\u00a0',
    ])
  }
  // SVG preview drawn imperatively after mount (renderSpec has no SVG namespace)
  children.push(['div', { class: 'doc-chart-canvas' }])

  const colCount = Math.max(chart.categories.length, ...chart.series.map((s) => s.values.length))
  const headCells: DomSpec[] = [['th', { class: 'doc-chart-corner' }, '\u00a0']]
  for (let c = 0; c < colCount; c++) {
    headCells.push([
      'th',
      { class: 'doc-chart-cell doc-chart-cat', 'data-cat': String(c), contenteditable: 'false' },
      chart.categories[c] || '\u00a0',
    ])
  }
  const rows: DomSpec[] = [['tr', {}, ...headCells]]
  chart.series.forEach((ser, s) => {
    const cells: DomSpec[] = [
      [
        'th',
        {
          class: `doc-chart-name${ser.name !== undefined ? ' doc-chart-cell' : ''}`,
          'data-ser': String(s),
          contenteditable: 'false',
          style: `border-left-color:${seriesColor(chart, s)}`,
        },
        ser.name ?? t('editorChartSeries', { num: s + 1 }),
      ],
    ]
    for (let c = 0; c < colCount; c++) {
      const value = ser.values[c]
      cells.push([
        'td',
        {
          // cache gaps have no pt to patch; they stay read-only
          class: value === null ? 'doc-chart-gap' : 'doc-chart-cell doc-chart-val',
          'data-ser': String(s),
          'data-val': String(c),
          contenteditable: 'false',
        },
        value === null || value === undefined ? '' : String(value),
      ])
    }
    rows.push(['tr', {}, ...cells])
  })
  children.push(['table', { class: 'doc-chart-data' }, ...rows])

  return ['div', { class: 'doc-chart' }, ...children]
}

const SVG_NS = 'http://www.w3.org/2000/svg'

/** round up to a 1/2/5×10ⁿ "nice" axis step */
function niceStep(target: number): number {
  const pow = 10 ** Math.floor(Math.log10(Math.max(target, 1e-9)))
  const n = target / pow
  return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10) * pow
}

interface ChartGeom {
  width: number
  height: number
  left: number
  right: number
  top: number
  bottom: number
  /** side-legend gutters; pies center in the remaining box (axes use left/right) */
  sideLeft?: number
  sideRight?: number
}

/** widest a chart draws at; the resize handle clamps to this too so mouseup never snaps back */
export const CHART_MAX_WIDTH_PX = 660

/** height of the title row above the plot SVG; heightPx = title row + plot,
 * so the resize write-back must add it before storing the measured SVG height */
export const CHART_TITLE_ROW_PX = 22

/** draw the read-only SVG preview into the node's .doc-chart-canvas */
export function drawChartSvg(dom: HTMLElement, chart: ChartDisplay | null): void {
  const canvas = dom.querySelector<HTMLElement>('.doc-chart-canvas')
  if (!canvas || !chart?.series.length) return
  // series-name legend inside the SVG: the data grid is an editing affordance
  // (hidden unless the block is selected), so the printed chart must carry the
  // legend itself, like Word/LibreOffice output. Pie legends list the
  // categories (one colored slice each), not the series.
  const isPie = chart.kind === 'pie'
  const legendNames = isPie
    ? chart.categories
    : chart.series.map((s, i) => s.name ?? t('editorChartSeries', { num: i + 1 }))
  const legendColor = (i: number) => (isPie ? pointColor(chart, 0, i) : seriesColor(chart, i))
  const showLegend = isPie
    ? legendNames.length > 0
    : chart.series.length > 1 || chart.series.some((s) => s.name)
  // the title row renders above the SVG but Word draws the title inside the
  // drawing extent; shrink the plot so title + plot together fill heightPx,
  // or pagination gains ~22px per titled chart and drifts
  const titleRowPx = chart.title !== undefined ? CHART_TITLE_ROW_PX : 0
  // horizontal bars put category labels on the y axis; reserve room for the
  // longest one (10px axis font: CJK glyphs are ~1em wide, Latin ~0.7em)
  const catLabelPx = (s: string) =>
    [...s].reduce((w, ch) => w + (isCjk(ch.codePointAt(0) ?? 0) ? 10 : 7), 0)
  const maxCatPx = Math.max(0, ...chart.categories.map(catLabelPx))
  const width = Math.min(chart.widthPx ?? 560, CHART_MAX_WIDTH_PX)
  // l/r/tr legends stack vertically in a side gutter, like Word; other
  // positions (and legacy models without legendPos) keep the bottom row
  const sideLegend =
    showLegend && (chart.legendPos === 'l' || chart.legendPos === 'r' || chart.legendPos === 'tr')
  const legendW = sideLegend
    ? Math.min(Math.round(width * 0.35), 20 + Math.max(0, ...legendNames.map(catLabelPx)))
    : 0
  const legendLeft = sideLegend && chart.legendPos === 'l'
  const geom: ChartGeom = {
    width,
    height: (chart.heightPx ?? 240) - titleRowPx,
    // also capped against the chart's own width: resize allows 120px-wide
    // charts, and a gutter wider than the plot would flip plotW negative
    left:
      (chart.kind === 'bar' && chart.horizontal ? Math.min(140, width * 0.4, 16 + maxCatPx) : 46) +
      (legendLeft ? legendW : 0),
    right: 12 + (sideLegend && !legendLeft ? legendW : 0),
    top: 12,
    bottom: 26 + (showLegend && !sideLegend ? 18 : 0),
    ...(sideLegend ? (legendLeft ? { sideLeft: legendW } : { sideRight: legendW }) : {}),
  }
  const svg = document.createElementNS(SVG_NS, 'svg')
  svg.setAttribute('viewBox', `0 0 ${geom.width} ${geom.height}`)
  svg.setAttribute('class', 'doc-chart-svg')
  svg.style.width = `${geom.width}px`
  svg.style.height = `${geom.height}px`

  if (chart.kind === 'pie') drawPie(svg, chart, geom)
  else if (chart.kind === 'scatter' || chart.kind === 'bubble') drawScatter(svg, chart, geom)
  else if (chart.kind === 'bar' && chart.horizontal) drawAxesHorizontal(svg, chart, geom)
  else drawAxes(svg, chart, geom)

  if (sideLegend) {
    const entryH = 16
    const x0 = legendLeft ? 4 : geom.width - legendW + 4
    const y0 =
      chart.legendPos === 'tr'
        ? geom.top
        : Math.max(geom.top, (geom.height - legendNames.length * entryH) / 2)
    legendNames.forEach((name, i) => {
      const y = y0 + entryH * i
      svgEl(svg, 'rect', {
        x: String(x0),
        y: String(y),
        width: '8',
        height: '8',
        fill: legendColor(i),
      })
      svgEl(
        svg,
        'text',
        { x: String(x0 + 12), y: String(y + 8), class: 'doc-chart-axis-label' },
        name,
      )
    })
  } else if (showLegend) {
    const slot = geom.width / legendNames.length
    legendNames.forEach((name, i) => {
      const cx = slot * i + slot / 2
      svgEl(svg, 'rect', {
        x: String(cx - Math.min(name.length * 3.2, slot / 2 - 14) - 12),
        y: String(geom.height - 15),
        width: '8',
        height: '8',
        fill: legendColor(i),
      })
      svgEl(
        svg,
        'text',
        {
          x: String(cx),
          y: String(geom.height - 7),
          class: 'doc-chart-axis-label',
          'text-anchor': 'middle',
        },
        name,
      )
    })
  }

  canvas.replaceChildren(svg)
}

function svgEl(
  parent: Element,
  tag: string,
  attrs: Record<string, string>,
  text?: string,
): Element {
  const el = document.createElementNS(SVG_NS, tag)
  for (const [key, value] of Object.entries(attrs)) el.setAttribute(key, value)
  if (text !== undefined) el.textContent = text
  parent.appendChild(el)
  return el
}

/** per-category positive/negative stack sums plus the absolute total (percentStacked base) */
function stackSums(
  chart: ChartDisplay,
  cols: number,
): { pos: number[]; neg: number[]; abs: number[] } {
  const pos = new Array<number>(cols).fill(0)
  const neg = new Array<number>(cols).fill(0)
  const abs = new Array<number>(cols).fill(0)
  for (const ser of chart.series) {
    ser.values.forEach((v, c) => {
      if (v === null || c >= cols) return
      if (v >= 0) pos[c] += v
      else neg[c] += v
      abs[c] += Math.abs(v)
    })
  }
  return { pos, neg, abs }
}

/** bar / line / area charts share the same axes and scale */
function drawAxes(svg: SVGElement, chart: ChartDisplay, geom: ChartGeom): void {
  const cols = Math.max(chart.categories.length, ...chart.series.map((s) => s.values.length), 1)
  const stacked = (chart.kind === 'bar' || chart.kind === 'area') && chart.grouping !== undefined
  const pct = stacked && chart.grouping === 'percentStacked'
  const sums = stacked ? stackSums(chart, cols) : null
  // percentStacked: each value becomes its share of the category's absolute total
  const norm = (value: number, c: number) =>
    pct ? (sums!.abs[c] > 0 ? (value / sums!.abs[c]) * 100 : 0) : value
  const values = stacked
    ? [...sums!.pos.map(norm), ...sums!.neg.map(norm)]
    : chart.series.flatMap((s) => s.values).filter((v): v is number => v !== null)
  // "nice" axis bounds (1/2/5×10ⁿ step, integer-friendly labels like Word/LO)
  const rawMax = Math.max(0, ...values)
  const rawMin = Math.min(0, ...values)
  const step = niceStep((rawMax - rawMin) / 5 || 1)
  const min = Math.floor(rawMin / step) * step
  // Word/LO leave headroom: the top tick sits strictly above the data maximum
  // (percent axes stop at 100%)
  let max = Math.ceil(rawMax / step) * step || step
  if (!pct && rawMax > 0 && max <= rawMax + 1e-9) max += step
  const span = max - min || 1
  const plotW = geom.width - geom.left - geom.right
  const plotH = geom.height - geom.top - geom.bottom
  const yOf = (v: number) => geom.top + plotH - ((v - min) / span) * plotH
  const slotW = plotW / cols

  // horizontal gridlines with value labels
  const steps = Math.max(1, Math.round(span / step))
  for (let i = 0; i <= steps; i++) {
    const v = min + step * i
    const y = yOf(v)
    svgEl(svg, 'line', {
      x1: String(geom.left),
      y1: String(y),
      x2: String(geom.width - geom.right),
      y2: String(y),
      class: 'doc-chart-grid',
    })
    svgEl(
      svg,
      'text',
      {
        x: String(geom.left - 6),
        y: String(y + 3),
        class: 'doc-chart-axis-label',
        'text-anchor': 'end',
      },
      formatAxisValue(v) + (pct ? '%' : ''),
    )
  }

  // category labels
  chart.categories.forEach((cat, c) => {
    if (c >= cols) return
    svgEl(
      svg,
      'text',
      {
        x: String(geom.left + slotW * c + slotW / 2),
        y: String(geom.height - geom.bottom + 14),
        class: 'doc-chart-axis-label',
        'text-anchor': 'middle',
      },
      cat,
    )
  })

  if (chart.kind === 'bar' && stacked) {
    // one bar per category, series segments cumulated up (down for negatives)
    const posBase = new Array<number>(cols).fill(0)
    const negBase = new Array<number>(cols).fill(0)
    const barPad = slotW * 0.2
    chart.series.forEach((ser, s) => {
      ser.values.forEach((value, c) => {
        if (value === null || c >= cols) return
        const v = norm(value, c)
        const from = v >= 0 ? posBase[c] : negBase[c]
        const to = from + v
        const y0 = yOf(from)
        const y1 = yOf(to)
        svgEl(svg, 'rect', {
          x: String(geom.left + slotW * c + barPad),
          y: String(Math.min(y0, y1)),
          width: String(slotW - barPad * 2),
          height: String(Math.max(1, Math.abs(y0 - y1))),
          fill: seriesColor(chart, s),
        })
        if (v >= 0) posBase[c] = to
        else negBase[c] = to
      })
    })
  } else if (chart.kind === 'bar') {
    const groupPad = slotW * 0.15
    const barW = (slotW - groupPad * 2) / chart.series.length
    chart.series.forEach((ser, s) => {
      ser.values.forEach((value, c) => {
        if (value === null) return
        const x = geom.left + slotW * c + groupPad + barW * s
        const y0 = yOf(0)
        const y1 = yOf(value)
        svgEl(svg, 'rect', {
          x: String(x + barW * 0.08),
          y: String(Math.min(y0, y1)),
          width: String(barW * 0.84),
          height: String(Math.max(1, Math.abs(y0 - y1))),
          fill: seriesColor(chart, s),
        })
      })
    })
  } else if (chart.kind === 'area' && stacked) {
    // stacked / 100% stacked area: each band fills between the running total
    // below it and its own cumulated top (empty cells count as 0, like Excel)
    const bases = new Array<number>(cols).fill(0)
    const xAt = (c: number) => geom.left + slotW * c + slotW / 2
    chart.series.forEach((ser, s) => {
      const tops = bases.map((b, c) => b + norm(ser.values[c] ?? 0, c))
      const upper = tops.map((v, c) => `${xAt(c)},${yOf(v)}`)
      const lower = bases.map((v, c) => `${xAt(c)},${yOf(v)}`).reverse()
      svgEl(svg, 'polygon', {
        points: [...upper, ...lower].join(' '),
        fill: seriesColor(chart, s),
        'fill-opacity': '0.35',
        stroke: 'none',
      })
      svgEl(svg, 'polyline', {
        points: upper.join(' '),
        fill: 'none',
        stroke: seriesColor(chart, s),
        'stroke-width': '2',
      })
      tops.forEach((v, c) => (bases[c] = v))
    })
  } else {
    // line / area / other: one polyline per series through slot centers
    chart.series.forEach((ser, s) => {
      const points = ser.values
        .map((value, c) =>
          value === null ? null : `${geom.left + slotW * c + slotW / 2},${yOf(value)}`,
        )
        .filter((p): p is string => p !== null)
      if (points.length === 0) return
      if (chart.kind === 'area' && points.length > 1) {
        const first = points[0].split(',')[0]
        const last = points[points.length - 1].split(',')[0]
        svgEl(svg, 'polygon', {
          points: `${first},${yOf(0)} ${points.join(' ')} ${last},${yOf(0)}`,
          fill: seriesColor(chart, s),
          'fill-opacity': '0.35',
          stroke: 'none',
        })
      }
      svgEl(svg, 'polyline', {
        points: points.join(' '),
        fill: 'none',
        stroke: seriesColor(chart, s),
        'stroke-width': '2',
      })
      if (chart.markers) {
        for (const p of points) {
          const [x, y] = p.split(',')
          svgEl(svg, 'circle', { cx: x, cy: y, r: '2.5', fill: seriesColor(chart, s) })
        }
      }
    })
  }
}

/** scatter / bubble: numeric x-y plot; bubble radius scales by point size (area-true) */
function drawScatter(svg: SVGElement, chart: ChartDisplay, geom: ChartGeom): void {
  interface Pt {
    x: number
    y: number
    size: number | null
    c: number
  }
  const pts: Pt[][] = chart.series.map((ser) =>
    ser.values
      .map((y, c) => {
        // no xVal cache at all → 1-based index axis; a null inside the cache is a gap
        const x = ser.xValues ? (ser.xValues[c] ?? null) : c + 1
        return y === null || x === null ? null : { x, y, size: ser.sizes?.[c] ?? null, c }
      })
      .filter((p): p is Pt => p !== null),
  )
  const all = pts.flat()
  if (all.length === 0) return
  const axis = (vals: number[]) => {
    let rawMax = Math.max(...vals)
    let rawMin = Math.min(...vals)
    // Excel-style auto minimum: anchor at 0 unless the data sits far above it
    // (date-serial x values must not squash the points against the right edge)
    if (rawMin > 0 && rawMin <= rawMax - rawMin) rawMin = 0
    if (rawMax < 0) rawMax = 0
    const step = niceStep((rawMax - rawMin) / 5 || 1)
    const min = Math.floor(rawMin / step) * step
    let max = Math.ceil(rawMax / step) * step || step
    if (rawMax > 0 && max <= rawMax + 1e-9) max += step
    return { min, max, step, span: max - min || 1 }
  }
  const ax = axis(all.map((p) => p.x))
  const ay = axis(all.map((p) => p.y))
  const plotW = geom.width - geom.left - geom.right
  const plotH = geom.height - geom.top - geom.bottom
  const xOf = (v: number) => geom.left + ((v - ax.min) / ax.span) * plotW
  const yOf = (v: number) => geom.top + plotH - ((v - ay.min) / ay.span) * plotH

  // horizontal gridlines with value labels, like the category charts
  const ySteps = Math.max(1, Math.round(ay.span / ay.step))
  for (let i = 0; i <= ySteps; i++) {
    const v = ay.min + ay.step * i
    const y = yOf(v)
    svgEl(svg, 'line', {
      x1: String(geom.left),
      y1: String(y),
      x2: String(geom.width - geom.right),
      y2: String(y),
      class: 'doc-chart-grid',
    })
    svgEl(
      svg,
      'text',
      {
        x: String(geom.left - 6),
        y: String(y + 3),
        class: 'doc-chart-axis-label',
        'text-anchor': 'end',
      },
      formatAxisValue(v),
    )
  }

  // x labels: date/text categories sit under their own data points (Word puts
  // the cached texts along the axis); numeric x gets plain scale ticks
  const catNumeric =
    chart.categories.length === 0 ||
    chart.categories.every((c) => c === '' || Number.isFinite(Number(c)))
  const labelY = String(geom.height - geom.bottom + 14)
  if (catNumeric) {
    const xSteps = Math.max(1, Math.round(ax.span / ax.step))
    for (let i = 0; i <= xSteps; i++) {
      const v = ax.min + ax.step * i
      svgEl(
        svg,
        'text',
        { x: String(xOf(v)), y: labelY, class: 'doc-chart-axis-label', 'text-anchor': 'middle' },
        formatAxisValue(v),
      )
    }
  } else {
    const anchor = pts[0] ?? []
    let lastEnd = -Infinity
    chart.categories.forEach((cat, c) => {
      const p = anchor.find((pt) => pt.c === c)
      if (!p) return
      // greedy label thinning: drop labels that would overlap the previous one
      const x = xOf(p.x)
      const half = cat.length * 2.6
      if (x - half < lastEnd + 4) return
      lastEnd = x + half
      svgEl(
        svg,
        'text',
        { x: String(x), y: labelY, class: 'doc-chart-axis-label', 'text-anchor': 'middle' },
        cat,
      )
    })
  }

  const maxSize = Math.max(0, ...all.map((p) => p.size ?? 0))
  const rMax = Math.min(plotW, plotH) * 0.125
  pts.forEach((points, s) => {
    const color = seriesColor(chart, s)
    if (chart.series[s]?.line && points.length > 1) {
      svgEl(svg, 'polyline', {
        points: points.map((p) => `${xOf(p.x)},${yOf(p.y)}`).join(' '),
        fill: 'none',
        stroke: color,
        'stroke-width': '2',
      })
    }
    // scatterStyle "line"/"smooth" draws no point markers on lined series
    if (chart.kind === 'scatter' && !chart.markers && chart.series[s]?.line) return
    for (const p of points) {
      const bubble = chart.kind === 'bubble' && p.size !== null && maxSize > 0
      const r = bubble ? Math.max(3, rMax * Math.sqrt(Math.abs(p.size!) / maxSize)) : 3
      svgEl(svg, 'circle', {
        cx: String(xOf(p.x)),
        cy: String(yOf(p.y)),
        r: String(r),
        fill: color,
        ...(bubble ? { 'fill-opacity': '0.85', stroke: '#fff', 'stroke-width': '1' } : {}),
      })
    }
  })
}

/** horizontal bar charts (c:barDir="bar"): value axis on x, categories on y,
 * first category in the bottom row and series 1 at the bottom of each group, like Word */
function drawAxesHorizontal(svg: SVGElement, chart: ChartDisplay, geom: ChartGeom): void {
  const rows = Math.max(chart.categories.length, ...chart.series.map((s) => s.values.length), 1)
  const stacked = chart.grouping !== undefined
  const pct = chart.grouping === 'percentStacked'
  const sums = stacked ? stackSums(chart, rows) : null
  const norm = (value: number, c: number) =>
    pct ? (sums!.abs[c] > 0 ? (value / sums!.abs[c]) * 100 : 0) : value
  const values = stacked
    ? [...sums!.pos.map(norm), ...sums!.neg.map(norm)]
    : chart.series.flatMap((s) => s.values).filter((v): v is number => v !== null)
  const rawMax = Math.max(0, ...values)
  const rawMin = Math.min(0, ...values)
  const step = niceStep((rawMax - rawMin) / 5 || 1)
  const min = Math.floor(rawMin / step) * step
  let max = Math.ceil(rawMax / step) * step || step
  if (!pct && rawMax > 0 && max <= rawMax + 1e-9) max += step
  const span = max - min || 1
  const plotW = geom.width - geom.left - geom.right
  const plotH = geom.height - geom.top - geom.bottom
  const xOf = (v: number) => geom.left + ((v - min) / span) * plotW
  const slotH = plotH / rows

  // vertical gridlines with value labels along the bottom
  const steps = Math.max(1, Math.round(span / step))
  for (let i = 0; i <= steps; i++) {
    const v = min + step * i
    const x = xOf(v)
    svgEl(svg, 'line', {
      x1: String(x),
      y1: String(geom.top),
      x2: String(x),
      y2: String(geom.height - geom.bottom),
      class: 'doc-chart-grid',
    })
    svgEl(
      svg,
      'text',
      {
        x: String(x),
        y: String(geom.height - geom.bottom + 14),
        class: 'doc-chart-axis-label',
        'text-anchor': 'middle',
      },
      formatAxisValue(v) + (pct ? '%' : ''),
    )
  }

  chart.categories.forEach((cat, c) => {
    if (c >= rows) return
    svgEl(
      svg,
      'text',
      {
        x: String(geom.left - 6),
        y: String(geom.top + plotH - slotH * c - slotH / 2 + 3),
        class: 'doc-chart-axis-label',
        'text-anchor': 'end',
      },
      cat,
    )
  })

  if (stacked) {
    const posBase = new Array<number>(rows).fill(0)
    const negBase = new Array<number>(rows).fill(0)
    const barPad = slotH * 0.2
    chart.series.forEach((ser, s) => {
      ser.values.forEach((value, c) => {
        if (value === null || c >= rows) return
        const v = norm(value, c)
        const from = v >= 0 ? posBase[c] : negBase[c]
        const to = from + v
        const x0 = xOf(from)
        const x1 = xOf(to)
        svgEl(svg, 'rect', {
          x: String(Math.min(x0, x1)),
          y: String(geom.top + plotH - slotH * (c + 1) + barPad),
          width: String(Math.max(1, Math.abs(x0 - x1))),
          height: String(slotH - barPad * 2),
          fill: seriesColor(chart, s),
        })
        if (v >= 0) posBase[c] = to
        else negBase[c] = to
      })
    })
    return
  }

  const groupPad = slotH * 0.15
  const barH = (slotH - groupPad * 2) / chart.series.length
  chart.series.forEach((ser, s) => {
    ser.values.forEach((value, c) => {
      if (value === null) return
      const y = geom.top + plotH - slotH * c - groupPad - barH * (s + 1)
      const x0 = xOf(0)
      const x1 = xOf(value)
      svgEl(svg, 'rect', {
        x: String(Math.min(x0, x1)),
        y: String(y + barH * 0.08),
        width: String(Math.max(1, Math.abs(x0 - x1))),
        height: String(barH * 0.84),
        fill: seriesColor(chart, s),
      })
    })
  })
}

/** pie / doughnut preview renders the first series only */
function drawPie(svg: SVGElement, chart: ChartDisplay, geom: ChartGeom): void {
  const values = chart.series[0].values.map((v) => (v === null || v < 0 ? 0 : v))
  const total = values.reduce((a, b) => a + b, 0)
  if (total <= 0) return
  const availW = geom.width - (geom.sideLeft ?? 0) - (geom.sideRight ?? 0)
  const cx = (geom.sideLeft ?? 0) + availW / 2
  const cy = geom.height / 2
  const r = Math.min(availW, geom.height) / 2 - 16
  const ri = (r * Math.min(Math.max(chart.holePct ?? 0, 0), 90)) / 100
  const single = values.filter((v) => v > 0).length === 1
  let angle = -Math.PI / 2
  values.forEach((value, i) => {
    if (value === 0) return
    const sweep = (value / total) * Math.PI * 2
    const x1 = cx + r * Math.cos(angle)
    const y1 = cy + r * Math.sin(angle)
    const ix1 = cx + ri * Math.cos(angle)
    const iy1 = cy + ri * Math.sin(angle)
    angle += sweep
    const x2 = cx + r * Math.cos(angle)
    const y2 = cy + r * Math.sin(angle)
    const ix2 = cx + ri * Math.cos(angle)
    const iy2 = cy + ri * Math.sin(angle)
    const large = sweep > Math.PI ? 1 : 0
    let d: string
    if (single) {
      d = `M ${cx - r} ${cy} A ${r} ${r} 0 1 1 ${cx - r} ${cy - 0.01} Z`
      // reverse-sweep inner circle cuts the hole under the nonzero fill rule
      if (ri > 0) d += ` M ${cx - ri} ${cy} A ${ri} ${ri} 0 1 0 ${cx - ri} ${cy - 0.01} Z`
    } else if (ri > 0) {
      d =
        `M ${ix1} ${iy1} L ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} ` +
        `L ${ix2} ${iy2} A ${ri} ${ri} 0 ${large} 0 ${ix1} ${iy1} Z`
    } else {
      d = `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} Z`
    }
    svgEl(svg, 'path', { d, fill: pointColor(chart, 0, i), stroke: '#fff', 'stroke-width': '1' })
  })
}

function formatAxisValue(v: number): string {
  const rounded = Math.round(v * 100) / 100
  return Math.abs(rounded) >= 1000 ? String(Math.round(rounded)) : String(rounded)
}

/** Edit chart title / series names / category labels / cached values in place. */
export function wireChartEditing(
  dom: HTMLElement,
  getNode: () => PmNode,
  getPos: () => number | undefined,
  view: EditorView,
): (ProtectedContentEditor & { cleanup(): void }) | null {
  const chart = getNode().attrs.chartDisplay as ChartDisplay | null
  if (!chart) return null
  const targets = Array.from(dom.querySelectorAll<HTMLElement>('.doc-chart-cell, .doc-chart-title'))
  if (targets.length === 0) return null

  const setEditable = (editable: boolean) => {
    for (const target of targets)
      target.setAttribute('contenteditable', editable ? 'true' : 'false')
  }
  const commit = () => {
    const current = getNode()
    const model = current.attrs.chartDisplay as ChartDisplay | null
    if (!model) return
    const next: ChartDisplay = {
      ...model,
      categories: [...model.categories],
      series: model.series.map((s) => ({ ...s, values: [...s.values] })),
    }
    const title = dom.querySelector<HTMLElement>('.doc-chart-title')
    if (title && next.title !== undefined) next.title = protectedText(title).trim()
    for (const cat of Array.from(dom.querySelectorAll<HTMLElement>('.doc-chart-cat'))) {
      const c = parseInt(cat.dataset.cat ?? '', 10)
      if (c >= 0 && c < next.categories.length) next.categories[c] = protectedText(cat).trim()
    }
    for (const name of Array.from(
      dom.querySelectorAll<HTMLElement>('.doc-chart-name.doc-chart-cell'),
    )) {
      const s = parseInt(name.dataset.ser ?? '', 10)
      if (next.series[s]) next.series[s].name = protectedText(name).trim()
    }
    for (const cell of Array.from(dom.querySelectorAll<HTMLElement>('.doc-chart-val'))) {
      const s = parseInt(cell.dataset.ser ?? '', 10)
      const c = parseInt(cell.dataset.val ?? '', 10)
      const ser = next.series[s]
      if (!ser || c < 0 || c >= ser.values.length) continue
      const parsed = Number(protectedText(cell).trim().replace(/,/g, ''))
      // unparseable input keeps the original number instead of corrupting the cache
      if (Number.isFinite(parsed)) ser.values[c] = parsed
    }
    if (JSON.stringify(next) === JSON.stringify(model)) return
    const pos = getPos()
    if (typeof pos !== 'number') return
    view.dispatch(
      view.state.tr.setNodeMarkup(pos, undefined, { ...current.attrs, chartDisplay: next }),
    )
  }
  for (const target of targets) target.addEventListener('keydown', preventProtectedLineBreak)
  window.addEventListener('ai-docs-commit-tables', commit)
  return {
    setEditable,
    commit,
    cleanup: () => {
      for (const target of targets) target.removeEventListener('keydown', preventProtectedLineBreak)
      window.removeEventListener('ai-docs-commit-tables', commit)
    },
  }
}

/** rendering of an anchored textbox (code box / callout card); text is editable in place */
/**
 * WordArt preset CSS approximation applied to the entire textbox container.
 * color: used as -webkit-text-fill-color; stroke: optional -webkit-text-stroke.
 * Derived from the shared cross-app presets; the wordArt-N entries keep
 * sessions with blocks inserted by the old docs-only gallery rendering.
 */
const WORDART_CSS: Record<string, { color: string; stroke?: string; textShadow?: string }> = {
  'wordArt-1': { color: '#4472C4' },
  'wordArt-2': { color: '#7B2FBE', stroke: '1px #4472C4' },
  'wordArt-3': { color: 'transparent', stroke: '2px #4472C4' },
  'wordArt-4': { color: '#1F3864', textShadow: '2px 2px 4px rgba(0,0,0,0.5)' },
  'wordArt-5': { color: '#ED7D31', textShadow: '0 0 8px #ED7D31, 0 0 16px #ED7D31' },
  'wordArt-6': { color: '#C00000' },
}
for (const p of WORDART_PRESETS) {
  WORDART_CSS[p.id] = {
    color: p.fill,
    stroke: p.outline ? `${wordArtStrokePx(p.outline.widthEmu)}px ${p.outline.color}` : undefined,
  }
}

export function textboxBoxStyle(box: TextboxDisplay): string {
  const boxW = box.widthPx ?? 189
  const boxH = box.heightPx ?? 113
  // Word keeps shape text inside the preset's text rectangle (e.g. the
  // ellipse's inscribed rect); bodyPr insets apply within that rect
  const geomInset = box.prst ? shapeTextInsetsPx(box.prst, boxW, boxH) : null
  const pad = (v: number, geom: number): number => Math.round((v + geom) * 100) / 100
  const insetTop = pad(box.insetTopPx ?? 4.8, geomInset?.t ?? 0)
  const insetRight = pad(box.insetRightPx ?? 9.6, geomInset?.r ?? 0)
  const insetBottom = pad(box.insetBottomPx ?? 4.8, geomInset?.b ?? 0)
  const insetLeft = pad(box.insetLeftPx ?? 9.6, geomInset?.l ?? 0)
  // preset/custom geometry renders as an SVG background so the border follows
  // the outline (a clip-path would clip a CSS outline away with the box corners)
  const geomCss = box.pathData
    ? custGeomBackgroundCss(box.pathData, boxW, boxH, box.fill, box.borderColor)
    : box.prst
      ? shapeBackgroundCss(box.prst, boxW, boxH, box.fill, box.borderColor, {
          diag: box.lineDiag,
          flipH: box.flipH,
          flipV: box.flipV,
        })
      : null
  const waStyle = box.wordArtId ? WORDART_CSS[box.wordArtId] : undefined
  // picture fill (photo boxes / a:blipFill): tiles repeat at natural size,
  // stretch fills cover the whole box. Document data, hence inline.
  const fillImage = box.fillImageDataUrl
    ? `background-image:url("${box.fillImageDataUrl}");` +
      (box.fillTile
        ? 'background-repeat:repeat'
        : 'background-repeat:no-repeat;background-size:100% 100%')
    : ''
  const transforms = [box.rotDeg ? `rotate(${box.rotDeg}deg)` : '']
  const floatPos = box.floating
    ? `position:absolute;left:${((box.offsetXEmu ?? 0) / 9525).toFixed(1)}px;` +
      `top:${((box.offsetYEmu ?? 0) / 9525).toFixed(1)}px`
    : ''
  return [
    geomCss ?? '',
    !geomCss && box.fill ? `background-color:#${box.fill}` : '',
    !geomCss && box.borderColor ? `border-color:#${box.borderColor}` : '',
    !geomCss && box.borderColor && box.borderWidthPx ? `border-width:${box.borderWidthPx}px` : '',
    !geomCss && box.borderColor && box.borderDash ? `border-style:${box.borderDash}` : '',
    fillImage,
    // shape-style fontRef color: the box default, so runs carrying their own
    // w:color still override it through the run spans
    box.textColor ? `color:#${box.textColor}` : '',
    floatPos,
    box.widthPx ? `width:${box.widthPx}px` : '',
    // Word clips fixed-height (noAutofit) boxes instead of growing them
    box.heightPx ? `height:${box.heightPx}px` : '',
    `padding:${insetTop}px ${insetRight}px ${insetBottom}px ${insetLeft}px`,
    transforms.filter(Boolean).length > 0
      ? `transform:${transforms.filter(Boolean).join(' ')}`
      : '',
    waStyle?.color ? `-webkit-text-fill-color:${waStyle.color}` : '',
    waStyle?.stroke ? `-webkit-text-stroke:${waStyle.stroke}` : '',
    // imported VML WordArt outline (document data, not chrome)
    !waStyle?.stroke && box.textOutline
      ? `-webkit-text-stroke:${box.textOutline.widthPx}px #${box.textOutline.colorHex}`
      : '',
    // behindDoc anchor: under the body text (per box, so paragraphs mixing
    // behind and front drawings keep the split); the relativeHeight rank
    // orders overlapping floats within each band like the image z bands
    box.behind
      ? `z-index:${Math.min(-1, -1000 + (box.z ?? 0))}`
      : box.floating && box.z !== undefined
        ? `z-index:${Math.max(1, 2 + box.z)}`
        : '',
    // WordArt strings never wrap; spill instead of clipping when the
    // font-size approximation runs slightly wide
    box.nowrap ? 'white-space:nowrap;overflow:visible' : '',
    // wps:bodyPr anchor="ctr|b": text body hugs the box middle/bottom
    box.vAlign && box.heightPx
      ? `display:flex;flex-direction:column;justify-content:${
          box.vAlign === 'bottom' ? 'flex-end' : 'center'
        }`
      : '',
    waStyle?.textShadow ? `text-shadow:${waStyle.textShadow}` : '',
  ]
    .filter(Boolean)
    .join(';')
}

const AUTOSPACE_PAD_SPEC: DomSpec = ['span', { class: 'doc-autospace-pad' }]

/** static-DOM counterpart of the editor's autospace pad decorations */
function padSegments(text: string): unknown[] {
  const cuts = autospaceBoundaries(text)
  if (cuts.length === 0) return [text]
  const out: unknown[] = []
  let start = 0
  for (const cut of cuts) {
    out.push(text.slice(start, cut), AUTOSPACE_PAD_SPEC)
    start = cut
  }
  out.push(text.slice(start))
  return out
}

/** run → styled <span> (+ inline <img>) specs, shared by textbox and table-cell rendering */
export function runSpanSpecs(run: Run, autoSpace?: boolean): DomSpec[] {
  const out: DomSpec[] = []
  // a run can carry both w:t text and a w:drawing; the text renders before the
  // image (generate.ts / the editable path keep the same order)
  if (run.text !== '' || !run.image) out.push(textSpanSpec(run, autoSpace))
  if (run.image) {
    const attrs: Record<string, string> = { class: 'doc-inline-img', src: run.image.dataUrl }
    const { widthPx, heightPx } = run.image
    if (widthPx)
      attrs.style = `width:${widthPx}px;${heightPx ? `height:${heightPx}px` : 'height:auto'}`
    out.push(['img', attrs])
  }
  return out
}

function textSpanSpec(run: Run, autoSpace?: boolean): DomSpec {
  const cs = run.csFont && textHasComplexScript(run.text) ? run.csFont : undefined
  const letterSpacing = runLetterSpacingCss(run)
  const runStyle = [
    run.color ? `color:#${run.color}` : '',
    run.bold ? 'font-weight:700' : run.bold === false ? 'font-weight:normal' : '',
    run.italic ? 'font-style:italic' : run.italic === false ? 'font-style:normal' : '',
    run.underline ? 'text-decoration:underline' : '',
    run.font || run.fontAscii || cs
      ? `font-family:${
          cs
            ? cssCsFontFamily(cs, run.fontAscii, run.font)
            : cssRunFontFamily(run.fontAscii, run.font)
        }`
      : '',
    run.sizeHalfPoints ? `font-size:${run.sizeHalfPoints / 2}pt` : '',
    letterSpacing ? `letter-spacing:${letterSpacing}` : '',
    run.caps === 'all' ? 'text-transform:uppercase' : '',
    run.caps === 'small' ? 'font-variant-caps:small-caps' : '',
    run.caps === 'none' ? 'text-transform:none;font-variant-caps:normal' : '',
    // explicit autoSpaceDE/DN off also disables the browser's native gap (same as the editor path)
    autoSpace === false ? 'text-autospace:no-autospace' : '',
  ]
    .filter(Boolean)
    .join(';')
  const content = autoSpace === false ? [run.text] : padSegments(run.text)
  // hyperlink runs keep the editable path's look (.doc-link) and real href;
  // App-level click handling prevents in-place navigation (jump on mod+click)
  if (run.link?.href) {
    const attrs: Record<string, string> = { class: 'doc-link', href: run.link.href }
    if (run.link.tooltip) attrs.title = run.link.tooltip
    if (runStyle) attrs.style = runStyle
    return ['a', attrs, ...content]
  }
  return runStyle ? ['span', { style: runStyle }, ...content] : ['span', {}, ...content]
}

/** run spans with pads at run-boundary CJK-Latin seams (empty runs keep their span, no pad) */
function runSpansWithPads(runs: Run[], autoSpace?: boolean): DomSpec[] {
  const out: DomSpec[] = []
  let prevText = ''
  for (const run of runs) {
    if (run.text !== '') {
      if (autoSpace !== false && autospacePadBetween(prevText, run.text)) {
        out.push(AUTOSPACE_PAD_SPEC)
      }
      prevText = run.text
    }
    out.push(...runSpanSpecs(run, autoSpace))
  }
  return out
}

export function renderTextboxSpec(box: TextboxDisplay): DomSpec {
  const style = textboxBoxStyle(box)
  const boxAttrs: Record<string, string> = { class: 'doc-textbox' }
  if (style) boxAttrs.style = style
  // page-absolute V rendered from the anchor: syncFloatShifts re-pins it
  if (box.floating && box.pageRelV) boxAttrs['data-page-rel-v'] = '1'
  // page-absolute X: the column-layout counter-translate keys on this
  if (box.floating && box.pageRelX) boxAttrs['data-page-rel-x'] = '1'

  const paras: DomSpec[] = box.paras.map((para) => {
    const spans: DomSpec[] = runSpansWithPads(para.runs, para.autoSpace)
    const text = para.runs.map((r) => r.text).join('')
    // same line strut as body/cell paragraphs (factor by run fonts + run-size
    // strut + grid snapping); inheriting the page's computed line-height
    // instead inflated every CJK textbox line to the body's pixel value
    const fontStyles: string[] = []
    if (text) {
      fontStyles.push(`--doc-line-factor:${runsLineFactor(para.runs, text)}`)
      const fam = runsDeclaredFontFamily(para.runs)
      if (fam) fontStyles.push(`font-family:${fam}`)
      const strut = runStrutHalfPoints(para.runs)
      if (strut) fontStyles.push(...strutFontCss(strut))
    }
    const lineMult = cssAutoLineMult(para.lineRule, para.lineRawTwips, para.lineSpacing)
    const lh = cssLineHeight(para.lineRule, para.lineRawTwips, para.lineSpacing)
    const pStyles = [
      para.align ? `text-align:${para.align}` : '',
      // .doc-textbox-para's pre-wrap would still wrap a nowrap (WordArt) box
      box.nowrap ? 'white-space:nowrap' : '',
      ...fontStyles,
      // undeclared spacing inherits the document default via the
      // .doc-textbox-para stylesheet rule (an inline base would override it)
      lh ? `line-height:${lh}` : '',
      // explicit single (mult 1) still overrides an inherited style/doc multiple
      lineMult ? `--doc-line-mult:${lineMult}` : '',
      // w:snapToGrid=0 opts the paragraph out of docGrid snapping (Word applies
      // the typed line grid inside textboxes too)
      para.snapToGrid === false ? '--doc-grid-pitch:0.0001px' : '',
      para.indentLeft ? `margin-left:${para.indentLeft / 20}pt` : '',
      para.indentRight ? `margin-right:${para.indentRight / 20}pt` : '',
      para.indentFirstLine ? `text-indent:${para.indentFirstLine / 20}pt` : '',
      // != null: an explicit 0 twips must still emit (matches the sub-editor)
      para.spaceBeforeAuto
        ? `margin-top:${WORD_AUTO_SPACING_PT}pt`
        : para.spaceBefore != null
          ? `margin-top:${para.spaceBefore / 20}pt`
          : '',
      para.spaceAfterAuto
        ? `margin-bottom:${WORD_AUTO_SPACING_PT}pt`
        : para.spaceAfter != null
          ? `margin-bottom:${para.spaceAfter / 20}pt`
          : '',
    ]
      .filter(Boolean)
      .join(';')
    const fixedLh = (para.lineRule === 'exact' && para.lineRawTwips) || para.lineRule === 'atLeast'
    const pAttrs: Record<string, string> = {
      class: `doc-textbox-para${spans.length === 0 ? ' doc-textbox-para-empty' : ''}${
        fixedLh ? ' doc-lh-fixed' : ''
      }${para.snapToGrid === false ? ' doc-nosnap' : ''}`,
    }
    if (para.styleId) pAttrs['data-style'] = para.styleId
    if (pStyles) pAttrs.style = pStyles
    // empty paragraphs hold a <br> so the caret can land on them while editing
    return spans.length > 0 ? ['div', pAttrs, ...spans] : ['div', pAttrs, ['br', {}]]
  })

  return ['div', boxAttrs, ...paras]
}

/** Max explicit run size (half-points, plus uniformity) when every run declares one (blockAttrs' strut rule). */
function runStrutHalfPoints(runs: Run[]): { halfPoints: number; uniform: boolean } | null {
  let max: number | null = null
  let min: number | null = null
  for (const run of runs) {
    if (run.sizeHalfPoints == null) return null
    max = Math.max(max ?? 0, run.sizeHalfPoints)
    min = Math.min(min ?? Infinity, run.sizeHalfPoints)
  }
  return max === null ? null : { halfPoints: max, uniform: max === min }
}

/** Run[] port of extensions' latinParaFactor (declared run fonts override the doc factor). */
function latinRunsFactor(runs: Run[], scriptVar: string): string {
  let declaredMax = 0
  let undeclared = false
  for (const run of runs) {
    // ascii slot only, like extensions' latinParaFactor: an eastAsia-only
    // declaration must not set a Latin line's factor
    const family = run.fontAscii
    if (family) declaredMax = Math.max(declaredMax, lineHeightFactor(family))
    else undeclared = true
  }
  if (declaredMax <= 0) return scriptVar
  return undeclared ? `max(${scriptVar}, ${declaredMax})` : String(declaredMax)
}

/** Run[] port of extensions' paraDeclaredFontFamily (strut face follows the runs). */
function runsDeclaredFontFamily(runs: Run[]): string | null {
  let first: string | null = null
  for (const run of runs) {
    const ea = run.font || null
    const ascii = run.fontAscii || null
    if (!ea && !ascii) return null
    first ??= cssRunFontFamily(ascii, ea)
  }
  return first
}

/** Per-paragraph --doc-line-factor from runs (Run[] port of extensions' paraLineFactor). */
function runsLineFactor(runs: Run[], text: string): string {
  const scriptVar = paraLineFactorCss(text)
  if (!textHasCjk(text)) return latinRunsFactor(runs, scriptVar)
  let declaredMax = 0
  let undeclaredCjk = false
  for (const run of runs) {
    if (!textHasCjk(run.text)) continue
    const family = run.eaSlotEmpty === true ? null : (run.font ?? run.fontAscii)
    if (family && isCjkFontName(family)) {
      declaredMax = Math.max(declaredMax, cjkDeclaredLineFactor(family) ?? lineHeightFactor(family))
    } else undeclaredCjk = true
  }
  if (declaredMax <= 0) return scriptVar
  return undeclaredCjk ? `max(${scriptVar}, ${declaredMax})` : String(declaredMax)
}

/**
 * Cell paragraph block: run-size strut + line factor + explicit line spacing,
 * mirroring the main renderer's blockAttrs. Without it the cell inherits
 * .doc-page's line height as a computed px value (body font size), inflating
 * every line whose runs are smaller. Block divs keep innerText's
 * one-\n-per-paragraph semantics that cell edit write-back depends on.
 */
function cellParaSpec(
  content: unknown[],
  text: string,
  runs: Run[] | null,
  fmt?: TableParagraph,
): DomSpec {
  const styles: string[] = []
  // explicit per-paragraph direction (same rule as the editor's blockAttrs):
  // a bidiVisual table's dir="rtl" mirrors columns but must not reorder cell text
  styles.push(fmt?.bidi || inferredBidi(fmt, runs ?? undefined) ? 'direction:rtl' : 'direction:ltr')
  if (text) {
    // Korean cells break at spaces like Word (same rule as the editor's blockAttrs)
    if (textHasHangul(text)) styles.push('word-break:keep-all', 'overflow-wrap:anywhere')
    styles.push(`--doc-line-factor:${runs ? runsLineFactor(runs, text) : paraLineFactorCss(text)}`)
    const fam = runs ? runsDeclaredFontFamily(runs) : null
    if (fam) styles.push(`font-family:${fam}`)
    const strut = runs ? runStrutHalfPoints(runs) : null
    if (strut) styles.push(...strutFontCss(strut))
  }
  styles.push(
    `line-height:${cssLineHeight(fmt?.lineRule, fmt?.lineRawTwips, fmt?.lineSpacing) ?? cssGridLineBase()}`,
  )
  const cellMult = cssAutoLineMult(fmt?.lineRule, fmt?.lineRawTwips, fmt?.lineSpacing)
  if (cellMult && cellMult !== 1) styles.push(`--doc-line-mult:${cellMult}`)
  if (fmt?.spaceBeforeAuto) styles.push(`margin-top:${cssGridSpacingPt(WORD_AUTO_SPACING_PT)}`)
  else if (fmt?.spaceBefore) styles.push(`margin-top:${cssGridSpacingPt(fmt.spaceBefore / 20)}`)
  if (fmt?.spaceAfterAuto) styles.push(`margin-bottom:${cssGridSpacingPt(WORD_AUTO_SPACING_PT)}`)
  else if (fmt?.spaceAfter) styles.push(`margin-bottom:${cssGridSpacingPt(fmt.spaceAfter / 20)}`)
  // Word sizes an empty line by the paragraph mark / empty run (same as blockAttrs)
  if (!text && fmt?.emptyRunSizeHalfPoints) {
    styles.push(`font-size:${fmt.emptyRunSizeHalfPoints / 2}pt`)
  }
  // Western mark faces only (same scoping as blockAttrs): empty cells keep
  // the Latin-factor rule instead of growing to a CJK mark face
  if (!text && fmt?.emptyRunFontFamily && !isCjkFontName(fmt.emptyRunFontFamily)) {
    const fam = fmt.emptyRunFontFamily
    styles.push(`--doc-line-factor:${lineHeightFactor(fam)}`, `font-family:${cssFontFamily(fam)}`)
  }
  const attrs: Record<string, string> = { style: styles.join(';') }
  // empty paragraphs get the Latin factor (.doc-table .doc-p-empty) and a <br> line box
  if (!text) attrs.class = 'doc-p-empty'
  return content.length > 0 ? ['div', attrs, ...content] : ['div', attrs, ['br', {}]]
}

/** read-only <table> DOM spec from the display model (vMerge -> rowSpan);
 *  nested = rendered inside a cell, capped at the cell instead of spilling into page margins */
export function renderTableSpec(model: TableModel, nested = false): DomSpec {
  // grid positions per row (accounting for colSpan) so vertical merges line up
  const positions: number[][] = model.rows.map((row) => {
    let col = 0
    return row.map((cell) => {
      const at = col
      col += cell.colSpan ?? 1
      return at
    })
  })

  const bodyRows: DomSpec[] = model.rows.map((row, ri) => {
    const tds: DomSpec[] = []
    row.forEach((cell, ci) => {
      if (cell.vMerge === 'continue') return
      let rowSpan = 1
      if (cell.vMerge === 'restart') {
        const gridCol = positions[ri][ci]
        for (let r = ri + 1; r < model.rows.length; r++) {
          const idx = positions[r].indexOf(gridCol)
          if (idx === -1 || model.rows[r][idx].vMerge !== 'continue') break
          rowSpan++
        }
      }
      const style = [
        // gridBefore/gridAfter placeholder: bare grid space, never bordered/filled
        cell.gridGap ? 'border:none;background:none' : '',
        // vertical-text cells: writing-mode rides the .cell-vert/.cell-clip wrapper below
        cell.textDirection ? 'position:relative' : '',
        cell.color ? `color:#${cell.color}` : '',
        cell.bold ? 'font-weight:600' : '',
        cell.fill ? `background:#${cell.fill}` : '',
        cell.align ? `text-align:${cell.align}` : '',
        cell.vAlign && cell.vAlign !== 'top'
          ? `vertical-align:${cell.vAlign === 'center' ? 'middle' : 'bottom'}`
          : '',
        // w:tcBorders — nested/read-only tables get no default gridlines, so
        // per-cell borders are the only line source for style-less documents
        ...(['top', 'left', 'bottom', 'right'] as const).map((side) => {
          const v = borderLineCss(cell.borders?.[side])
          return v ? `border-${side}:${v}` : ''
        }),
        ...(['top', 'left', 'bottom', 'right'] as const).map((side) =>
          cell.cellMarTwips?.[side] !== undefined
            ? `padding-${side}:${(cell.cellMarTwips[side]! / 15).toFixed(1)}px`
            : '',
        ),
      ]
        .filter(Boolean)
        .join(';')
      const tdAttrs: Record<string, string> = {}
      if (style) tdAttrs.style = style
      // not in-flow evidence for the .cell-vert switch (same as tableCellHtml)
      if (cell.gridGap) tdAttrs['data-grid-gap'] = '1'
      if (cell.colSpan && cell.colSpan > 1) tdAttrs.colspan = String(cell.colSpan)
      if (rowSpan > 1) tdAttrs.rowspan = String(rowSpan)
      const paraBlocks: DomSpec[] = cell.richParas?.length
        ? cell.richParas.map((p) => {
            const runs = p.runs.filter((run) => run.text !== '' || run.image)
            return cellParaSpec(
              runSpansWithPads(runs, p.autoSpace),
              runs.map((r) => r.text).join(''),
              runs,
              p,
            )
          })
        : cell.paras.map((p) => cellParaSpec(p === '' ? [] : [...padSegments(p)], p, null))
      // nested tables spliced in at their paragraph anchors (cells with them are never editable)
      const nested = cell.nestedTables ?? []
      const anchorOf = (i: number) =>
        Math.min(cell.nestedTableAnchors?.[i] ?? paraBlocks.length, paraBlocks.length)
      const content: unknown[] = []
      let ni = 0
      paraBlocks.forEach((blk, pi) => {
        while (ni < nested.length && anchorOf(ni) <= pi)
          content.push(renderTableSpec(nested[ni++], true))
        content.push(blk)
      })
      while (ni < nested.length) content.push(renderTableSpec(nested[ni++], true))
      if (content.length === 0) content.push('\u00a0')
      const clip = cellClipTwips(model, ri, cell, rowSpan)
      const wm = cellWritingMode(cell.textDirection ?? null)
      if (clip !== null) {
        const clipStyle = cellClipStyle(cell.vAlign ?? null, clip)
        tds.push([
          'td',
          tdAttrs,
          ['div', { class: 'cell-clip', style: wm ? `${clipStyle};${wm}` : clipStyle }, ...content],
        ])
      } else if (wm) {
        // rotated text wraps into the row height the horizontal cells produce
        // instead of stretching the row (same rule as tableCellSpec); rows with
        // no other height source drop the wrapper back into flow via the
        // cell-vert-host :has() switch in styles.css
        tdAttrs.class = tdAttrs.class ? `${tdAttrs.class} cell-vert-host` : 'cell-vert-host'
        const av = cellVAlignGridCss(cell.vAlign ?? null)
        tds.push([
          'td',
          tdAttrs,
          ['div', { class: 'cell-vert', style: av ? `${wm};${av}` : wm }, ...content],
        ])
      } else {
        tds.push(['td', tdAttrs, ...content])
      }
    })
    const trAttrs: Record<string, string> = {}
    const rh = model.rowHeightsTwips?.[ri]
    if (rh) trAttrs.style = `height:${((rh / 1440) * 96).toFixed(1)}px`
    return ['tr', trAttrs, ...tds]
  })

  const tableChildren: unknown[] = []
  const colPx = !model.widthPct
    ? model.colWidthsTwips?.map((w) => Math.max(24, Math.round(w / 15)))
    : undefined
  if (colPx) {
    tableChildren.push(['colgroup', {}, ...colPx.map((w) => ['col', { style: `width:${w}px` }])])
  } else if (model.colWidthsPct) {
    tableChildren.push([
      'colgroup',
      {},
      ...model.colWidthsPct.map((w) => ['col', { style: `width:${w.toFixed(2)}%` }]),
    ])
  }
  tableChildren.push(['tbody', {}, ...bodyRows])
  const tableAttrs: Record<string, string> = { class: 'doc-table' }
  if (model.bidiVisual) tableAttrs.dir = 'rtl'
  const tableStyles: string[] = []
  let centerMargin: string | null = null
  if (model.widthPct)
    // 'pct' table widths are a share of the section's TEXT COLUMN (see
    // DocTable.renderHTML): differing-margin/width sections resolve through the
    // per-block --doc-content-w; nested tables stay relative to their cell
    tableStyles.push(
      nested
        ? `width:${model.widthPct}%`
        : `width:calc(var(--doc-content-w,100%) * ${Number(model.widthPct) / 100})`,
    )
  else if (colPx) {
    // nested tables are capped by their cell; top-level ones spill into the page
    // margins like Word (centered: both sides via negative-margin centering,
    // left-aligned: right up to the paper edge — see DocTable.renderHTML);
    // indent shifts the table right, so it comes out of the budget
    const widthPx = colPx.reduce((sum, w) => sum + w, 0)
    // --doc-content-w: per-block section content width (differing-width sections); defaults to the page content box
    const contentW = 'var(--doc-content-w,100%)'
    // w:tblLayout fixed holds the declared widths even past the paper edge (see DocTable.renderHTML)
    if (!nested && model.fixedLayout) {
      tableStyles.push(`width:${widthPx}px`, 'max-width:none')
      if (model.align === 'center')
        centerMargin = `margin-left:calc((${contentW} - ${widthPx}px)/2)`
    } else if (!nested && model.align === 'center') {
      const paper = `calc(${contentW} + var(--doc-margin-left,var(--doc-margin-right,0px)) + var(--doc-margin-right,0px))`
      tableStyles.push(`width:min(${widthPx}px,${paper})`)
      centerMargin = `margin-left:calc((${contentW} - min(${widthPx}px,${paper}))/2)`
    } else {
      const indented =
        model.align !== 'center' && model.align !== 'right' && (model.indentTwips ?? 0) > 0
      const indentPx = indented ? model.indentTwips! / 15 : 0
      const base = nested ? '100%' : `calc(${contentW} + var(--doc-margin-right,0px))`
      const avail = indentPx
        ? nested
          ? `calc(100% - ${indentPx.toFixed(1)}px)`
          : `calc(${contentW} + var(--doc-margin-right,0px) - ${indentPx.toFixed(1)}px)`
        : base
      tableStyles.push(`width:min(${widthPx}px,${avail})`)
    }
  }
  const pad = cellPadCss(model.cellMarTwips ?? null)
  if (pad) tableStyles.push(`--doc-cell-pad:${pad}`)
  // w:tblCellSpacing: the CSS gap is twice the per-cell-side value (see DocTable.renderHTML)
  if (model.cellSpacingTwips) {
    const gapPx = ((model.cellSpacingTwips * 2) / 15).toFixed(1)
    tableStyles.push('border-collapse:separate', `border-spacing:${gapPx}px`)
  }
  if (model.fill) tableStyles.push(`background-color:#${model.fill}`)
  tableStyles.push(...tableBordersCss((model.borders as TableBordersAttr | undefined) ?? null))
  if (model.align === 'center') {
    if (centerMargin) tableStyles.push(centerMargin)
    else tableStyles.push('margin-left:auto', 'margin-right:auto')
  } else if (model.align === 'right') tableStyles.push('margin-left:auto')
  else if (model.indentTwips)
    tableStyles.push(`margin-left:${(model.indentTwips / 15).toFixed(1)}px`)
  if (tableStyles.length > 0) tableAttrs.style = tableStyles.join(';')
  return ['table', tableAttrs, ...tableChildren]
}

// ---- marks ----
