import type { Node as PmNode } from '@tiptap/pm/model'
import {} from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import {} from '@tiptap/pm/tables'
import { WORDART_PRESETS, wordArtStrokePx } from '@genoffice/ui'
import { cssDualFontFamily, cssFontFamily } from '../line-metrics'
import { shapeBackgroundCss } from './shape-svg'
import { t } from '../i18n/locale'
import {
  type ChartDisplay,
  type FieldDisplay,
  type FormulaDisplay,
  type TableModel,
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
  cellPadCss,
  preventProtectedLineBreak,
  protectedText,
  tableBordersCss,
} from './extensions'

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
    const num: DomSpec[] = field.num
      ? [['span', { class: 'doc-toc-num', contenteditable: 'false' }, field.num]]
      : []
    return [
      'div',
      attrs,
      ...num,
      ['span', { class: 'doc-toc-title', contenteditable: 'false' }, field.left || '\u00a0'],
      ['span', { class: 'doc-toc-dots' }],
      ['span', { class: 'doc-toc-page', contenteditable: 'false' }, field.right ?? ''],
    ]
  }
  if (field.kind === 'pageBreak') {
    return ['div', { class: 'doc-field-pagebreak' }, ['span', {}, t('editorPageBreak')]]
  }
  if (field.kind === 'text' && field.left) {
    return ['div', { class: 'doc-field-text', contenteditable: 'false' }, field.left]
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

/** Office theme default accent colors, used for new charts */
const CHART_PALETTE = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']

const chartColor = (i: number) => `#${CHART_PALETTE[i % CHART_PALETTE.length]}`

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
          style: `border-left-color:${chartColor(s)}`,
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

interface ChartGeom {
  width: number
  height: number
  left: number
  right: number
  top: number
  bottom: number
}

/** widest a chart draws at; the resize handle clamps to this too so mouseup never snaps back */
export const CHART_MAX_WIDTH_PX = 660

/** draw the read-only SVG preview into the node's .doc-chart-canvas */
export function drawChartSvg(dom: HTMLElement, chart: ChartDisplay | null): void {
  const canvas = dom.querySelector<HTMLElement>('.doc-chart-canvas')
  if (!canvas || !chart?.series.length) return
  const geom: ChartGeom = {
    width: Math.min(chart.widthPx ?? 560, CHART_MAX_WIDTH_PX),
    height: chart.heightPx ?? 240,
    left: 46,
    right: 12,
    top: 12,
    bottom: 26,
  }
  const svg = document.createElementNS(SVG_NS, 'svg')
  svg.setAttribute('viewBox', `0 0 ${geom.width} ${geom.height}`)
  svg.setAttribute('class', 'doc-chart-svg')
  svg.style.width = `${geom.width}px`
  svg.style.height = `${geom.height}px`

  if (chart.kind === 'pie') drawPie(svg, chart, geom)
  else drawAxes(svg, chart, geom)

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

/** bar / line / area charts share the same axes and scale */
function drawAxes(svg: SVGElement, chart: ChartDisplay, geom: ChartGeom): void {
  const values = chart.series.flatMap((s) => s.values).filter((v): v is number => v !== null)
  const max = Math.max(0, ...values)
  const min = Math.min(0, ...values)
  const span = max - min || 1
  const plotW = geom.width - geom.left - geom.right
  const plotH = geom.height - geom.top - geom.bottom
  const cols = Math.max(chart.categories.length, ...chart.series.map((s) => s.values.length), 1)
  const yOf = (v: number) => geom.top + plotH - ((v - min) / span) * plotH
  const slotW = plotW / cols

  // horizontal gridlines with value labels
  const steps = 4
  for (let i = 0; i <= steps; i++) {
    const v = min + (span / steps) * i
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

  if (chart.kind === 'bar') {
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
          fill: chartColor(s),
        })
      })
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
          fill: chartColor(s),
          'fill-opacity': '0.35',
          stroke: 'none',
        })
      }
      svgEl(svg, 'polyline', {
        points: points.join(' '),
        fill: 'none',
        stroke: chartColor(s),
        'stroke-width': '2',
      })
    })
  }
}

/** pie preview renders the first series only */
function drawPie(svg: SVGElement, chart: ChartDisplay, geom: ChartGeom): void {
  const values = chart.series[0].values.map((v) => (v === null || v < 0 ? 0 : v))
  const total = values.reduce((a, b) => a + b, 0)
  if (total <= 0) return
  const cx = geom.width / 2
  const cy = geom.height / 2
  const r = Math.min(geom.width, geom.height) / 2 - 16
  let angle = -Math.PI / 2
  values.forEach((value, i) => {
    if (value === 0) return
    const sweep = (value / total) * Math.PI * 2
    const x1 = cx + r * Math.cos(angle)
    const y1 = cy + r * Math.sin(angle)
    angle += sweep
    const x2 = cx + r * Math.cos(angle)
    const y2 = cy + r * Math.sin(angle)
    const large = sweep > Math.PI ? 1 : 0
    const d =
      values.filter((v) => v > 0).length === 1
        ? `M ${cx - r} ${cy} A ${r} ${r} 0 1 1 ${cx - r} ${cy - 0.01} Z`
        : `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} Z`
    svgEl(svg, 'path', { d, fill: chartColor(i), stroke: '#fff', 'stroke-width': '1' })
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
  const insetTop = box.insetTopPx ?? 4.8
  const insetRight = box.insetRightPx ?? 9.6
  const insetBottom = box.insetBottomPx ?? 4.8
  const insetLeft = box.insetLeftPx ?? 9.6
  // preset geometry renders as an SVG background so the border follows the
  // outline (a clip-path would clip a CSS outline away with the box corners)
  const geomCss = box.prst
    ? shapeBackgroundCss(
        box.prst,
        box.widthPx ?? 189,
        box.heightPx ?? 113,
        box.fill,
        box.borderColor,
      )
    : null
  const waStyle = box.wordArtId ? WORDART_CSS[box.wordArtId] : undefined
  return [
    geomCss ?? '',
    !geomCss && box.fill ? `background:#${box.fill}` : '',
    !geomCss && box.borderColor ? `border-color:#${box.borderColor}` : '',
    box.widthPx ? `width:${box.widthPx}px` : '',
    // Word clips fixed-height (noAutofit) boxes instead of growing them
    box.heightPx ? `height:${box.heightPx}px` : '',
    `padding:${insetTop}px ${insetRight}px ${insetBottom}px ${insetLeft}px`,
    waStyle?.color ? `-webkit-text-fill-color:${waStyle.color}` : '',
    waStyle?.stroke ? `-webkit-text-stroke:${waStyle.stroke}` : '',
    waStyle?.textShadow ? `text-shadow:${waStyle.textShadow}` : '',
  ]
    .filter(Boolean)
    .join(';')
}

export function renderTextboxSpec(box: TextboxDisplay): DomSpec {
  const style = textboxBoxStyle(box)
  const boxAttrs: Record<string, string> = { class: 'doc-textbox' }
  if (style) boxAttrs.style = style

  const paras: DomSpec[] = box.paras.map((para) => {
    const spans: DomSpec[] = para.runs.map((run) => {
      const runStyle = [
        run.color ? `color:#${run.color}` : '',
        run.bold ? 'font-weight:700' : '',
        run.italic ? 'font-style:italic' : '',
        run.underline ? 'text-decoration:underline' : '',
        run.font || run.fontAscii
          ? `font-family:${
              run.font && run.fontAscii
                ? cssDualFontFamily(run.fontAscii, run.font)
                : cssFontFamily((run.font ?? run.fontAscii)!)
            }`
          : '',
        run.sizeHalfPoints ? `font-size:${run.sizeHalfPoints / 2}pt` : '',
      ]
        .filter(Boolean)
        .join(';')
      return runStyle ? ['span', { style: runStyle }, run.text] : ['span', {}, run.text]
    })
    const pStyles = [
      para.align ? `text-align:${para.align}` : '',
      para.lineSpacing ? `line-height:${para.lineSpacing * 1.2}` : '',
      para.indentLeft ? `margin-left:${para.indentLeft / 20}pt` : '',
      para.indentRight ? `margin-right:${para.indentRight / 20}pt` : '',
      para.indentFirstLine ? `text-indent:${para.indentFirstLine / 20}pt` : '',
      para.spaceBefore ? `margin-top:${para.spaceBefore / 20}pt` : '',
      para.spaceAfter ? `margin-bottom:${para.spaceAfter / 20}pt` : '',
    ]
      .filter(Boolean)
      .join(';')
    const pAttrs: Record<string, string> = {
      class: `doc-textbox-para${spans.length === 0 ? ' doc-textbox-para-empty' : ''}`,
    }
    if (pStyles) pAttrs.style = pStyles
    // empty paragraphs hold a <br> so the caret can land on them while editing
    return spans.length > 0 ? ['div', pAttrs, ...spans] : ['div', pAttrs, ['br', {}]]
  })

  return ['div', boxAttrs, ...paras]
}

/** read-only <table> DOM spec from the display model (vMerge -> rowSpan) */
export function renderTableSpec(model: TableModel): DomSpec {
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
        cell.textDirection === 'tbRl'
          ? 'writing-mode:vertical-rl'
          : cell.textDirection === 'btLr'
            ? 'writing-mode:sideways-lr'
            : '',
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
      if (cell.colSpan && cell.colSpan > 1) tdAttrs.colspan = String(cell.colSpan)
      if (rowSpan > 1) tdAttrs.rowspan = String(rowSpan)
      const content: unknown[] = []
      cell.paras.forEach((p, i) => {
        if (i > 0) content.push(['br', {}])
        if (p !== '') content.push(p)
      })
      for (const nt of cell.nestedTables ?? []) content.push(renderTableSpec(nt))
      if (content.length === 0) content.push('\u00a0')
      tds.push(['td', tdAttrs, ...content])
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
  if (model.widthPct) tableStyles.push(`width:${model.widthPct}%`)
  else if (colPx) tableStyles.push(`width:${colPx.reduce((sum, w) => sum + w, 0)}px`)
  const pad = cellPadCss(model.cellMarTwips ?? null)
  if (pad) tableStyles.push(`--doc-cell-pad:${pad}`)
  tableStyles.push(...tableBordersCss((model.borders as TableBordersAttr | undefined) ?? null))
  if (model.align === 'center') tableStyles.push('margin-left:auto', 'margin-right:auto')
  else if (model.align === 'right') tableStyles.push('margin-left:auto')
  else if (model.indentTwips)
    tableStyles.push(`margin-left:${(model.indentTwips / 15).toFixed(1)}px`)
  if (tableStyles.length > 0) tableAttrs.style = tableStyles.join(';')
  return ['table', tableAttrs, ...tableChildren]
}

// ---- marks ----
