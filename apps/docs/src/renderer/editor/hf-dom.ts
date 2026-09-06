import {
  PAGE_MARK,
  TOTAL_PAGES_MARK,
  type HeaderFooter,
  type HfCellParaProps,
  type HfImage,
  type HfParagraph,
  type HfTableCell,
  type HfTableRow,
  type HfTextBox,
  type Run,
  type SectionSettings,
} from '@genoffice/docx-engine'
import {
  cssLineHeight,
  cssRunFontFamily,
  estimateHfHeight,
  runLetterSpacingCss,
} from '../line-metrics'
import { setDkBackground, setDkBorder, setDkColor } from './dark-page'
import { INLINE_RULE_CLASS, inlineRuleDecls } from './inline-rule'

/**
 * Plain-DOM header/footer rendering for the canvas page gaps (M4 always-on
 * pagination). Mirrors HeaderFooterArea's read-only display: same classes,
 * same PAGE_MARK / NUMPAGES substitution — but built imperatively because page gaps
 * are ProseMirror widget decorations, not React children.
 */

/** w:pBdr line CSS: declared color/width when present, legacy 1px #444 otherwise (document content color) */
export function paraBorderCss(line?: { color?: string; szPt?: number }): string {
  const widthPx = line?.szPt ? Math.max(1, Math.round((line.szPt * 96) / 72)) : 1
  return `${widthPx}px solid #${line?.color ?? '444'}`
}

/** One OOXML border → CSS border value; 'none' means explicitly borderless */
export function borderLineCss(
  b: { style: string; szEighths?: number; color?: string } | undefined | null,
): string | null {
  if (!b) return null
  if (b.style === 'none' || b.style === 'nil') return 'none'
  const px = Math.max(1, Math.round(((b.szEighths ?? 4) / 8 / 72) * 96))
  const dash =
    b.style === 'dashed'
      ? 'dashed'
      : b.style === 'dotted'
        ? 'dotted'
        : b.style === 'double'
          ? 'double'
          : 'solid'
  const color = b.color && b.color !== 'auto' ? `#${b.color}` : '#000'
  return `${px}px ${dash} ${color}`
}

const px = (twips: number) => `${(twips / TWIPS_PER_PX).toFixed(1)}px`

/** Word's default cell margins (twips) for sides the table / cell leaves undeclared */
const HF_CELL_MAR = { top: 0, right: 108, bottom: 0, left: 108 } as const

/** cell box declarations (camelCase), its border CSS per side (callers add the
 *  dark-page twins) and the text-area width tabs lay out against */
export interface HfCellGeometry {
  style: Record<string, string>
  borders: Partial<Record<'t' | 'b' | 'l' | 'r', string>>
  textWidthPx?: number
}

/** Layout-table cell box: declared column width (or its share of the row),
 *  cell margins as padding, vertical alignment, resolved borders. */
export function hfCellGeometry(cell: HfTableCell): HfCellGeometry {
  const style: Record<string, string> = {}
  const mar = { ...HF_CELL_MAR, ...cell.marTwips }
  let textWidthPx: number | undefined
  if (cell.widthTwips) {
    style.width = px(cell.widthTwips)
    style.flex = '0 0 auto'
    textWidthPx = Math.max(0, cell.widthTwips - mar.left - mar.right) / TWIPS_PER_PX
  } else if (cell.widthPct) style.width = `${cell.widthPct}%`
  if (cell.marTwips || cell.widthTwips || cell.borders) {
    style.padding = `${px(mar.top)} ${px(mar.right)} ${px(mar.bottom)} ${px(mar.left)}`
  }
  if (cell.vAlign === 'center') style.justifyContent = 'center'
  else if (cell.vAlign === 'bottom') style.justifyContent = 'flex-end'
  if (cell.align) {
    style.textAlign =
      cell.align === 'left' || cell.align === 'center' || cell.align === 'right'
        ? cell.align
        : 'justify'
  }
  const borders: HfCellGeometry['borders'] = {}
  const sides = [
    ['top', 't', 'borderTop'],
    ['right', 'r', 'borderRight'],
    ['bottom', 'b', 'borderBottom'],
    ['left', 'l', 'borderLeft'],
  ] as const
  for (const [side, key, prop] of sides) {
    const css = borderLineCss(cell.borders?.[side])
    if (!css || css === 'none') continue
    style[prop] = css
    borders[key] = css
  }
  return { style, borders, ...(textWidthPx != null ? { textWidthPx } : {}) }
}

/** Layout-table row box: left offset of the table edge and the declared w:trHeight */
export function hfRowStyle(row: HfTableRow | undefined): Record<string, string> {
  const style: Record<string, string> = {}
  if (!row) return style
  // the strip clips at its edge, so a legacy table hanging into the margin
  // (tblInd below the cell margin) starts flush with it instead
  if (row.indentTwips && row.indentTwips > 0) {
    style[row.bidiVisual ? 'marginRight' : 'marginLeft'] = px(row.indentTwips)
  }
  if (row.bidiVisual) style.flexDirection = 'row-reverse'
  if (row.heightTwips) {
    if (row.heightRule === 'exact') {
      style.height = px(row.heightTwips)
      style.overflow = 'hidden'
    } else style.minHeight = px(row.heightTwips)
  }
  return style
}

/** cell paragraph box; a tab-wrapped paragraph spans several line boxes, so
 *  first-line-only (indent, space before) and last-line-only (space after)
 *  props land on the matching line */
export function hfCellParaStyle(
  props: HfCellParaProps | undefined,
  line: { first: boolean; last: boolean } = { first: true, last: true },
): Record<string, string> {
  const style: Record<string, string> = {}
  if (!props) return style
  if (props.align) {
    style.textAlign =
      props.align === 'left' || props.align === 'center' || props.align === 'right'
        ? props.align
        : 'justify'
  }
  if (props.indentLeft) style.paddingLeft = px(props.indentLeft)
  if (line.first && props.indentFirstLine) style.textIndent = px(props.indentFirstLine)
  if (line.first && props.spaceBefore) style.marginTop = px(props.spaceBefore)
  if (line.last && props.spaceAfter) style.marginBottom = px(props.spaceAfter)
  return style
}

function applyRunStyle(span: HTMLElement, run: Run): void {
  if (run.bold) span.style.fontWeight = '600'
  else if (run.bold === false) span.style.fontWeight = 'normal'
  if (run.italic) span.style.fontStyle = 'italic'
  else if (run.italic === false) span.style.fontStyle = 'normal'
  const deco = [run.underline && 'underline', run.strike && 'line-through'].filter(Boolean)
  if (deco.length > 0) span.style.textDecoration = deco.join(' ')
  if (run.color) {
    span.style.color = `#${run.color}`
    setDkColor(span, run.color) // dark-page twin; the authored color stays the declaration
  }
  if (run.sizeHalfPoints) span.style.fontSize = `${run.sizeHalfPoints / 2}pt`
  const letterSpacing = runLetterSpacingCss(run)
  if (letterSpacing) span.style.letterSpacing = letterSpacing
  if (run.font || run.fontAscii) span.style.fontFamily = cssRunFontFamily(run.fontAscii, run.font)
  if (run.caps === 'all') span.style.textTransform = 'uppercase'
  else if (run.caps === 'small') span.style.fontVariantCaps = 'small-caps'
  else if (run.caps === 'none') {
    span.style.textTransform = 'none'
    span.style.fontVariantCaps = 'normal'
  }
}

/** one tab-delimited chunk of a paragraph: the runs after the k-th tab, laid out at its stop */
export interface HfTabSegment {
  runs: Run[]
  /** offset from the text-column left edge; pct = margin-relative (w:ptab / implicit stops) */
  left: { px: number } | { pct: number }
  anchor: 'left' | 'center' | 'right'
}

export interface HfTabLayout {
  lead: Run[]
  segments: HfTabSegment[]
  minHeightPt?: number
  /** w:jc of the tab-laid line: Word lays tabs out in left-aligned space, then
   *  shifts the whole line (right: line end at the column edge; center: line
   *  centered). lineEndPx = laid-out line end in column space. */
  shift?: { align: 'center' | 'right'; lineEndPx: number }
  /** runs a cell-edge tab pushed onto the next line (Word 2013+ wrap) */
  rest?: Run[]
}

/** tab layout inside a table cell: stops measure from the cell text edge,
 *  the lead text starts at the paragraph indent, tabs reaching widthPx overflow */
export interface HfCellTabOpts {
  startPx: number
  widthPx?: number
  overflow?: HfTableRow['tabOverflow']
}

const TWIPS_PER_PX = 15
/** default tab grid past the last explicit stop (720 twips, matches Word) */
const HF_DEFAULT_GRID_PX = 48
const HF_DEFAULT_FONT_PT = 10.5

/** effective strip base size (pt): the mounted --hf-default-fs (docStyleCss,
 *  shrink-only) when present, so tab measures match what the strip paints */
function hfBaseFontPt(): number {
  if (typeof document === 'undefined') return HF_DEFAULT_FONT_PT
  const host = document.querySelector('.page-wrap, .doc-page, .pv-page')
  if (!host) return HF_DEFAULT_FONT_PT
  const pt = parseFloat(getComputedStyle(host).getPropertyValue('--hf-default-fs'))
  return Number.isFinite(pt) && pt > 0 ? pt : HF_DEFAULT_FONT_PT
}

let hfMeasureCtx: CanvasRenderingContext2D | null | undefined
/** approximate rendered width of hf runs (px): canvas mirror of applyRunStyle's font mapping.
 *  `display` is the caller's field substitution (PAGE_MARK -> the page digits): the raw
 *  private-use marks measure as notdef boxes and would skew stop placement. */
function hfRunsWidthPx(runs: Run[], display: (text: string) => string): number {
  if (hfMeasureCtx === undefined) {
    hfMeasureCtx =
      typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null
  }
  let w = 0
  const basePt = hfBaseFontPt()
  for (const run of runs) {
    if (run.image?.widthPx) w += run.image.widthPx
    if (!run.text) continue
    const shown = display(run.text)
    const text = run.caps === 'all' ? shown.toUpperCase() : shown
    const sizePt = run.sizeHalfPoints ? run.sizeHalfPoints / 2 : basePt
    if (hfMeasureCtx) {
      const family =
        run.font || run.fontAscii ? cssRunFontFamily(run.fontAscii, run.font) : 'sans-serif'
      // Detached-canvas font parsing cannot resolve var() (no element style context —
      // the whole assignment would be ignored): inline each var's fallback chain
      const canvasFamily = family.replace(/var\(--[\w-]+,([^)]*)\)/g, '$1')
      hfMeasureCtx.font = `${run.italic ? 'italic ' : ''}${run.bold ? '600 ' : ''}${(sizePt * 4) / 3}px ${canvasFamily}`
      w += hfMeasureCtx.measureText(text).width
    } else {
      w += text.length * sizePt * 0.5 * (4 / 3)
    }
    if (run.charSpacingTwips) w += (((run.charSpacingTwips / 20) * 4) / 3) * text.length
  }
  return w
}

/**
 * Tab layout of a header/footer paragraph, Word semantics: style-chain and
 * direct stops are pre-merged by the parser; each tab advances from the
 * current position to the next stop past it (center/right stops align their
 * segment at the stop), the default grid takes over past the last stop, and
 * w:jc then shifts the whole laid-out line. Returns null when the paragraph
 * has no tab. Without any stops the implicit header stops apply (center at
 * half the column, then its right edge).
 */
export function hfTabSegments(
  para: Pick<HfParagraph, 'runs' | 'tabStops' | 'ptabAligns' | 'align'>,
  display: (text: string) => string = (t) => t,
  cell?: HfCellTabOpts,
): HfTabLayout | null {
  if (!para.runs.some((r) => r.text.includes('\t'))) return null
  const chunks: Run[][] = [[]]
  for (const run of para.runs) {
    const pieces = run.text.split('\t')
    chunks[chunks.length - 1].push({ ...run, text: pieces[0] })
    for (const piece of pieces.slice(1)) chunks.push([{ ...run, text: piece }])
  }
  const stops = (para.tabStops ?? [])
    .filter(
      (s) =>
        (s.val === 'left' || s.val === 'center' || s.val === 'right' || s.val === 'decimal') &&
        Number.isFinite(s.pos),
    )
    .map((s) => ({ x: s.pos / TWIPS_PER_PX, val: s.val }))
    .sort((a, b) => a.x - b.x)

  const segments: HfTabSegment[] = []
  let usedPct = false
  let rest: Run[] | undefined
  let x = (cell?.startPx ?? 0) + hfRunsWidthPx(chunks[0], display)
  for (let k = 1; k < chunks.length; k++) {
    const runs = chunks[k].filter((r) => r.text !== '')
    // w:ptab carries its own margin-relative alignment and ignores tab stops
    const ptab = para.ptabAligns?.[k - 1]
    if (ptab) {
      usedPct = true
      segments.push({
        runs,
        left: { pct: ptab === 'center' ? 50 : ptab === 'right' ? 100 : 0 },
        anchor: ptab,
      })
      continue
    }
    if (stops.length === 0 && !cell) {
      // implicit Word header/footer stops: first tab to the column center, next to its right edge
      usedPct = true
      segments.push(
        k === 1
          ? { runs, left: { pct: 50 }, anchor: 'center' }
          : { runs, left: { pct: 100 }, anchor: 'right' },
      )
      continue
    }
    const segW = hfRunsWidthPx(runs, display)
    const stop = stops.find((s) => s.x > x + 0.5)
    const target = stop
      ? stop.x
      : (Math.floor((x + 0.5) / HF_DEFAULT_GRID_PX) + 1) * HF_DEFAULT_GRID_PX
    // a tab reaching the cell edge: Word 2013+ wraps what follows onto the next
    // line (starting at the indent); legacy layout leaves it beyond the edge,
    // where the cell clips it
    if (cell?.widthPx != null && cell.overflow === 'wrap' && target >= cell.widthPx - 0.5) {
      if (runs.length > 0 || k + 1 < chunks.length) {
        rest = chunks
          .slice(k)
          .flatMap((c, i) =>
            i === 0 ? c : c.map((r, j) => (j === 0 ? { ...r, text: `\t${r.text}` } : r)),
          )
        const first = rest.findIndex((r) => r.text !== '')
        if (first >= 0) rest[first] = { ...rest[first], text: rest[first].text.replace(/^\s+/, '') }
        rest = rest.filter((r) => r.text !== '')
      }
      break
    }
    const val = stop?.val ?? 'left'
    const placed = Math.max(
      x,
      val === 'center' ? target - segW / 2 : val === 'left' ? target : target - segW,
    )
    if (runs.length > 0) segments.push({ runs, left: { px: placed }, anchor: 'left' })
    x = placed + segW
  }
  // absolutely positioned segments add no flow height: an oversized run after a
  // tab (or an empty lead) would collapse to the strip's min-height and clip
  const maxHalfPoints = Math.max(0, ...para.runs.map((r) => r.sizeHalfPoints ?? 0))
  const align = para.align
  return {
    lead: chunks[0].filter((r) => r.text !== ''),
    segments,
    ...(maxHalfPoints > 0 ? { minHeightPt: (maxHalfPoints / 2) * 1.3 } : {}),
    ...(!usedPct && !cell && (align === 'center' || align === 'right')
      ? { shift: { align, lineEndPx: x } }
      : {}),
    ...(rest?.length ? { rest } : {}),
  }
}

/** Tab layout of one cell paragraph as display lines (null without tabs).
 *  Stops measure from the cell text edge; an overflowing tab wraps the rest
 *  onto further lines under Word 2013+ layout. */
export function hfCellTabLines(
  runs: Run[],
  props: HfCellParaProps | undefined,
  geom: HfCellGeometry,
  row: HfTableRow | undefined,
  display: (text: string) => string = (t) => t,
): HfTabLayout[] | null {
  if (!runs.some((r) => r.text.includes('\t'))) return null
  const opts: HfCellTabOpts = {
    startPx: Math.max(0, (props?.indentLeft ?? 0) + (props?.indentFirstLine ?? 0)) / TWIPS_PER_PX,
    ...(geom.textWidthPx != null ? { widthPx: geom.textWidthPx } : {}),
    ...(row?.tabOverflow ? { overflow: row.tabOverflow } : {}),
  }
  const lines: HfTabLayout[] = []
  let current: Run[] | undefined = runs
  while (current) {
    const line = hfTabSegments({ runs: current, tabStops: props?.tabStops }, display, opts)
    if (!line) {
      lines.push({ lead: current, segments: [] })
      break
    }
    // a segment starting past the cell edge is invisible; its box must not
    // exist either (Chromium's print fit-to-paper measures clipped boxes and
    // would shrink the whole export)
    if (opts.widthPx != null) {
      const w = opts.widthPx
      line.segments = line.segments.filter((s) => !('px' in s.left) || s.left.px < w - 0.5)
    }
    lines.push(line)
    current = line.rest
  }
  return lines
}

/** cell segment placement: the box ends at the cell text edge (see hfCellTabLines) */
export function hfCellSegStyle(seg: HfTabSegment): Record<string, string> {
  const left = 'px' in seg.left ? seg.left.px : 0
  return { left: `${left.toFixed(1)}px`, maxWidth: `calc(100% - ${left.toFixed(1)}px)` }
}

/** CSS left of a tab segment, including the line's w:jc shift (never left of its laid-out spot) */
export function hfSegLeftCss(seg: HfTabSegment, layout: HfTabLayout): string {
  if ('pct' in seg.left) return `${seg.left.pct}%`
  const px = seg.left.px
  const s = layout.shift
  if (!s) return `${px.toFixed(1)}px`
  return s.align === 'right'
    ? `max(${px.toFixed(1)}px, calc(100% - ${(s.lineEndPx - px).toFixed(1)}px))`
    : `max(${px.toFixed(1)}px, calc(50% + ${(px - s.lineEndPx / 2).toFixed(1)}px))`
}

/** text-indent that applies the line's w:jc shift to the in-flow lead runs */
export function hfLeadIndentCss(layout: HfTabLayout): string | null {
  const s = layout.shift
  if (!s) return null
  return s.align === 'right'
    ? `max(0px, calc(100% - ${s.lineEndPx.toFixed(1)}px))`
    : `max(0px, calc(50% - ${(s.lineEndPx / 2).toFixed(1)}px))`
}

/** effective paragraphs: rich paras when present, else the legacy single line (mirrors HeaderFooterArea) */
function parasOf(value: HeaderFooter): HfParagraph[] {
  if (value.paras?.length) return value.paras
  const runs: Run[] = value.text ? [{ text: value.text }] : []
  if (value.pageNumber && !value.text.includes('#') && !value.text.includes(PAGE_MARK)) {
    runs.push({ text: runs.length > 0 ? ` ${PAGE_MARK}` : PAGE_MARK })
  }
  return [{ align: 'center', runs }]
}

/**
 * Largest declared run size (pt) when every text-bearing run of the strip
 * declares one, else null (some run inherits the strip base). Whitespace-only
 * runs don't size lines in Word and only count when nothing else does (an
 * empty paragraph is sized by its mark).
 */
/** strip paragraph line-height from its (style-resolved) w:spacing; strips sit
 *  outside the body grid, so multiples stay plain unitless factors */
export function hfParaLineHeightCss(para: {
  lineRule?: 'auto' | 'atLeast' | 'exact'
  lineRawTwips?: number
  lineSpacing?: number
}): string | null {
  if (para.lineRule === 'exact' || para.lineRule === 'atLeast') {
    return cssLineHeight(para.lineRule, para.lineRawTwips, para.lineSpacing)
  }
  const m =
    para.lineSpacing ??
    (para.lineRule === 'auto' && para.lineRawTwips ? para.lineRawTwips / 240 : undefined)
  return m ? `calc(var(--doc-line-factor,1.2) * ${m})` : null
}

export function hfDeclaredStrutPt(paras: HfParagraph[]): number | null {
  let text: number | null = null
  let blank: number | null = null
  const scan = (runs: Run[]): boolean => {
    for (const run of runs) {
      if (!run.text && run.image) continue
      const sz = run.sizeHalfPoints
      if (!(run.text ?? '').trim()) {
        if (sz) blank = Math.max(blank ?? 0, sz)
        continue
      }
      if (!sz) return false
      text = Math.max(text ?? 0, sz)
    }
    return true
  }
  for (const p of paras) {
    if (p.boxAnchored) continue
    if (p.cells) {
      for (const c of p.cells) for (const runs of c.paras) if (!scan(runs)) return null
    } else if (!scan(p.runs)) return null
  }
  const half = text ?? blank
  return half ? half / 2 : null
}

/** parsed parts carry PAGE fields as PAGE_MARK; only mark-free values fall back to the user-typed '#' (mirrors HeaderFooterArea) */
function paraHasPageMark(p: HfParagraph): boolean {
  return [p.runs, ...(p.cells?.flatMap((c) => c.paras) ?? [])].some((rs) =>
    rs.some((r) => r.text.includes(PAGE_MARK)),
  )
}

export function hfUsesLegacyHash(value: HeaderFooter): boolean {
  if (!value.pageNumber) return false
  if (value.text.includes(PAGE_MARK)) return false
  return !value.paras?.some(paraHasPageMark)
}

export function hfHasPageField(value: HeaderFooter | null | undefined): boolean {
  return Boolean(
    value &&
    (value.pageNumber || value.text.includes(PAGE_MARK) || value.paras?.some(paraHasPageMark)),
  )
}

/** Remove Page Numbers: strip PAGE_MARK fields (legacy parts: only the first user-typed '#'), keeping literal '#' text and rich formatting.
 *  Layout-table rows (`cells`) pass through untouched: they are display-only and saving keeps
 *  the part's original w:tbl bytes, so a table-held PAGE field cannot be removed. */
export function hfWithoutPageMarks(value: HeaderFooter): HeaderFooter {
  let legacyHash = hfUsesLegacyHash(value)
  const strip = (t: string) => {
    const out = t.replaceAll(PAGE_MARK, '')
    if (legacyHash && out.includes('#')) {
      legacyHash = false
      return out.replace('#', '')
    }
    return out
  }
  if (!value.paras?.length) return { ...value, text: strip(value.text), pageNumber: false }
  const paras = value.paras
    .map((p) =>
      p.cells
        ? p
        : { ...p, runs: p.runs.map((r) => ({ ...r, text: strip(r.text) })).filter((r) => r.text) },
    )
    // dedicated page-number paragraphs go away entirely; table rows and user-typed blank lines stay
    .filter((p, i) => p.cells != null || p.runs.length > 0 || value.paras![i].runs.length === 0)
  const text = paras
    .map((p) =>
      [...p.runs, ...(p.cells?.flatMap((c) => c.paras.flat()) ?? [])].map((r) => r.text).join(''),
    )
    .join('')
  if (!text) return { ...value, text: '', paras: undefined, pageNumber: false }
  return { ...value, text, paras, pageNumber: false }
}

export function hfHasVisibleContent(
  value: HeaderFooter | null | undefined,
  images?: HfImage[],
): boolean {
  if (images?.length) return true
  if (!value) return false
  return Boolean(
    value.text ||
    value.pageNumber ||
    value.paras?.some((p) => p.runs.length > 0 || p.cells?.length),
  )
}

/** rendered strip heights keyed by content signature (pagination recomputes per page) */
const hfHeightCache = new Map<string, number>()
// strips probed before an @font-face finished loading wrapped in the fallback
// face (prod-sas 047: a footer measured two lines, one once its font arrived)
let hfProbeFontEpoch = 0
export function bumpHfProbeFontEpoch(): void {
  hfProbeFontEpoch++
}

function hfProbeHost(): HTMLElement | null {
  if (typeof document === 'undefined' || !document.body) return null
  let host = document.getElementById('hf-strip-probe')
  if (!host) {
    host = document.createElement('div')
    host.id = 'hf-strip-probe'
    // .doc-page: the probe inherits the same document-default font/line-height
    // CSS the real gap strips get inside the editor root
    host.className = 'doc-page'
    host.style.cssText =
      'position:absolute;left:-99999px;top:0;visibility:hidden;pointer-events:none'
    // editor-chrome floors (dashed separator, clickable-strip min-height) are
    // not Word geometry: with them the measure over-reserves ~1 line per strip
    // paragraph, costing body lines on every page of every footered document.
    // In document.head because the probe host replaces its children per measure
    const neutralize = document.createElement('style')
    neutralize.textContent =
      '#hf-strip-probe .page-hf { padding: 0; border: none; }' +
      '#hf-strip-probe .page-hf-para { min-height: 0; }' +
      // floating-box content draws at its anchor, not in the strip flow
      '#hf-strip-probe .page-hf-box-anchored { display: none; }'
    document.head.appendChild(neutralize)
    document.body.appendChild(host)
  }
  return host
}

/** djb2 over the mounted doc-scoped stylesheets (memoized on the concatenated string) */
let mountedCssMemo: { text: string; hash: string } | null = null
function mountedDocCssSig(): string {
  const text = Array.from(document.querySelectorAll('style[data-doc-css]'))
    .map((s) => s.textContent ?? '')
    .join(' ')
  if (mountedCssMemo?.text === text) return mountedCssMemo.hash
  let h = 5381
  for (let i = 0; i < text.length; i++) h = ((h << 5) + h + text.charCodeAt(i)) | 0
  const hash = `${text.length}:${h}`
  mountedCssMemo = { text, hash }
  return hash
}

/**
 * Reserved height (px) of a header/footer strip for body push-down. The
 * line-box estimate decides alone outside a DOM (tests); in the renderer the
 * strip is laid out in a hidden .doc-page probe — the same makeGapHfEl markup
 * and CSS the canvas draws — and that measure decides: real wraps, borders and
 * line heights, in both directions (the estimate's width model wrapped a line
 * Word keeps whole on SAS-2 sample 017 and reserved a body line per page).
 * Only floating-image reservations are estimate-only and floor the DOM value.
 * The probe measures under whatever
 * doc styles are mounted right now and the cache keys on their content, so a
 * value computed before a new document's CSS commits is re-measured on the
 * caller's post-commit pass (App's hfMeasureEpoch) instead of going stale.
 */
export function hfReservedHeightPx(
  kind: 'header' | 'footer',
  value: HeaderFooter | null,
  contentWidthPx: number,
  images?: HfImage[] | null,
  geom?: { marginTopPx: number; headerDistPx: number },
): number {
  const est = estimateHfHeight(value, contentWidthPx, images, geom)
  const inline = (images ?? []).filter((img) => !img.floating)
  if (!hfHasVisibleContent(value, inline) || contentWidthPx <= 0) return est
  if (typeof document === 'undefined' || !document.body) return est
  // cell-run images can carry megabytes of base64: key on the dataUrl length only
  const noDataUrl = (k: string, v: unknown) => (k === 'dataUrl' ? String(v).length : v)
  const key =
    `${kind}|${Math.round(contentWidthPx)}|${mountedDocCssSig()}|${hfProbeFontEpoch}|` +
    `${JSON.stringify(value, noDataUrl)}|` +
    inline.map((im) => `${im.widthPx ?? ''}x${im.heightPx ?? ''}:${im.dataUrl.length}`).join(',')
  let dom = hfHeightCache.get(key)
  if (dom === undefined) {
    const host = hfProbeHost()
    if (!host) return est
    const el = makeGapHfEl({
      kind,
      value: value ?? { text: '' },
      images: inline,
      pageNo: 88,
      pageTotal: 88,
    })
    el.style.position = 'static'
    el.style.left = 'auto'
    el.style.transform = 'none'
    el.style.width = `${contentWidthPx}px`
    // an empty paragraph still occupies its line in Word; with the chrome
    // min-height neutralized it needs a real line box to measure that
    for (const p of el.querySelectorAll('.page-hf-para')) {
      if (!(p.textContent ?? '').length && p.childElementCount === 0) {
        p.appendChild(document.createTextNode('\u200b'))
      }
    }
    host.replaceChildren(el)
    dom = el.getBoundingClientRect().height
    host.replaceChildren()
    if (hfHeightCache.size > 300) hfHeightCache.clear()
    hfHeightCache.set(key, dom)
  }
  if (dom <= 0) return est
  return Math.max(dom, estimateHfHeight(null, contentWidthPx, images, geom))
}

/** page (paper) geometry a floating header image positions against, px */
export interface HfFloatBox {
  pageW: number
  pageH: number
  marginLeft: number
  marginRight: number
  /** effective top margin after header push-down (mirrors the strip layout) */
  marginTop: number
  marginBottom: number
  /** header strip top (w:headerDist, px): origin of paragraph-relative vertical offsets */
  headerDist: number
  /** raw sectPr top margin (px, before push-down): origin wrapped margin-relative images reserve from */
  sectMarginTop: number
}

/**
 * Position of a floating header image in page coordinates (px from the page's
 * top-left corner), with the translate that resolves center/right/bottom
 * anchors without knowing the image's natural size. wp:posOffset offsets win;
 * alignment fields reproduce the legacy VML placement (margin-box corners).
 */
export function hfFloatPagePos(
  img: HfImage | HfTextBox,
  box: HfFloatBox,
): { x: number; y: number; translateX: 0 | -50 | -100; translateY: 0 | -50 | -100 } {
  let x: number
  let translateX: 0 | -50 | -100 = 0
  if (img.posXPx != null) {
    x = img.posHRel === 'page' ? img.posXPx : box.marginLeft + img.posXPx
  } else if (img.posH === 'center') {
    // legacy VML entries carry no rel and keep the page center
    x =
      img.posHRel === 'margin'
        ? box.marginLeft + (box.pageW - box.marginLeft - box.marginRight) / 2
        : box.pageW / 2
    translateX = -50
  } else if (img.posH === 'right') {
    x = img.posHRel === 'page' ? box.pageW : box.pageW - box.marginRight
    translateX = -100
  } else {
    x = img.posHRel === 'page' ? 0 : box.marginLeft
  }
  let y: number
  let translateY: 0 | -50 | -100 = 0
  if (img.posYPx != null) {
    // paragraph/margin-rel wrapped images use the same origins the body
    // push-down estimate measures from (header strip top / raw sectPr margin),
    // so the image never chases the pushed-down margin; watermarks keep the
    // effective margin
    const wrapped = img.wrap && img.wrap !== 'none' && !img.behind
    y =
      img.posVRel === 'page'
        ? img.posYPx
        : img.posVRel === 'paragraph'
          ? box.headerDist + img.posYPx
          : (wrapped ? box.sectMarginTop : box.marginTop) + img.posYPx
  } else if (img.posV === 'center') {
    y =
      img.posVRel === 'margin'
        ? box.marginTop + (box.pageH - box.marginTop - box.marginBottom) / 2
        : box.pageH / 2
    translateY = -50
  } else if (img.posV === 'bottom') {
    y = img.posVRel === 'page' ? box.pageH : box.pageH - box.marginBottom
    translateY = -100
  } else {
    y = img.posVRel === 'page' ? 0 : box.marginTop
  }
  return { x, y, translateX, translateY }
}

/**
 * Floating header image (picture watermark) for the canvas print view, drawn
 * once per page behind the body text (z-index -1; .view-print .doc-page
 * isolates). host 'gap': anchored in a page gap, whose bottom edge sits
 * marginTop above the next page's content. host 'lead': anchored in the
 * zero-height first-page widget at the content-box origin.
 */

/**
 * Word's washout preset (gain 19661f, blacklevel 22938f) is the linear map
 * out = 0.3*in + 0.7 per channel (MS-ODRAW brightness/contrast), so white
 * stays white and the image fades toward white. invert-brightness-invert
 * reproduces it exactly: 1 - 0.3*(1 - in); brightness/contrast chains clamp
 * at white and repaint white pixels as grey.
 */
export const HF_WASHOUT_FILTER = 'invert(1) brightness(0.3) invert(1)'
/** <img> for a header/footer picture; an a:srcRect crop becomes an
 *  overflow-hidden window over a scaled and offset image (body-path technique) */
function hfImgNode(img: {
  dataUrl: string
  widthPx?: number
  heightPx?: number
  crop?: HfImage['crop']
}): HTMLElement {
  const im = document.createElement('img')
  im.src = img.dataUrl
  im.alt = ''
  im.draggable = false
  const c = img.crop
  if (c && img.widthPx && img.heightPx) {
    const span = (a: number, b: number) => Math.max(0.01, 1 - a - b)
    const sw = img.widthPx / span(c.l, c.r)
    const sh = img.heightPx / span(c.t, c.b)
    const win = document.createElement('span')
    win.style.display = 'inline-block'
    win.style.overflow = 'hidden'
    win.style.width = `${img.widthPx}px`
    win.style.height = `${img.heightPx}px`
    im.style.cssText =
      `position:absolute;left:${(-c.l * sw).toFixed(1)}px;top:${(-c.t * sh).toFixed(1)}px;` +
      `width:${sw.toFixed(1)}px;height:${sh.toFixed(1)}px;max-width:none`
    // the window anchors the absolute img even when the caller's class does
    // not position the wrapper itself
    win.style.position = 'relative'
    win.append(im)
    return win
  }
  if (img.widthPx) im.style.width = `${img.widthPx}px`
  if (img.heightPx) im.style.height = `${img.heightPx}px`
  return im
}

export function makeHfFloatImgEl(img: HfImage, box: HfFloatBox, host: 'gap' | 'lead'): HTMLElement {
  const el = hfImgNode(img)
  el.className = 'page-hf-float-img'
  // a crop wrapper carries an inline position:relative for the strip hosts;
  // the float path must stay absolute (inline style beats the class)
  el.style.position = 'absolute'
  const p = hfFloatPagePos(img, box)
  if (host === 'gap') {
    el.style.left = `${p.x}px`
    el.style.top = `calc(100% + ${p.y - box.marginTop}px)`
  } else {
    el.style.left = `${p.x - box.marginLeft}px`
    el.style.top = `${p.y - box.marginTop}px`
  }
  if (p.translateX || p.translateY) {
    el.style.transform = `translate(${p.translateX}%, ${p.translateY}%)`
  }
  if (img.widthPx) el.style.width = `${img.widthPx}px`
  if (img.heightPx) el.style.height = `${img.heightPx}px`
  if (img.washout) el.style.filter = HF_WASHOUT_FILTER
  return el
}

/** page geometry (px) a strip's floating textboxes position against */
export interface HfStripGeom {
  pageW: number
  pageH: number
  marginLeft: number
  marginRight: number
  /** raw sectPr margins: origins of margin-relative offsets */
  marginTop: number
  marginBottom: number
  /** where the header strip's top edge sits on the page (w:headerDist unless the strip is pinned higher) */
  headerStripTop: number
  /** w:footerDist: the footer strip's bottom edge sits this far above the page bottom */
  footerDist: number
  /** page x of the strip's left edge; absent = the left margin (a centered strip on unequal side margins sits elsewhere) */
  stripLeft?: number
}

export function hfStripGeom(set: SectionSettings): HfStripGeom {
  const px = (twips: number) => (twips / 1440) * 96
  return {
    pageW: px(set.pageWidth),
    pageH: px(set.pageHeight),
    marginLeft: px(set.marginLeft),
    marginRight: px(set.marginRight),
    marginTop: px(set.marginTop),
    marginBottom: px(set.marginBottom),
    headerStripTop: px(set.headerDist ?? 720),
    footerDist: px(set.footerDist ?? 720),
  }
}

/** Word's textbox text insets (0.1in / 0.05in) when bodyPr declares none */
const TEXTBOX_INSETS_PX: [number, number, number, number] = [9.6, 4.8, 9.6, 4.8]

/**
 * Strip-relative placement of a floating textbox (the strip starts at the left
 * margin; a header strip's top edge is headerStripTop, a footer strip's bottom
 * edge is footerDist above the page bottom). null = no usable anchor position,
 * the box's paragraphs stack in the strip flow instead.
 */
export function hfTextBoxStyle(
  box: HfTextBox,
  kind: 'header' | 'footer',
  geom: HfStripGeom,
): Record<string, string> | null {
  if (box.posXPx == null && box.posYPx == null && !box.posH && !box.posV) return null
  const pos = hfFloatPagePos(box, {
    pageW: geom.pageW,
    pageH: geom.pageH,
    marginLeft: geom.marginLeft,
    marginRight: geom.marginRight,
    marginTop: geom.marginTop,
    marginBottom: geom.marginBottom,
    headerDist: geom.headerStripTop,
    sectMarginTop: geom.marginTop,
  })
  const w = box.widthPx
  const h = box.heightPx
  let x = pos.x + (w != null ? (w * pos.translateX) / 100 : 0)
  // Word keeps anchored objects on the paper: a box past the right edge is pulled back in
  if (w != null) x = Math.max(0, Math.min(x, geom.pageW - w))
  const y = pos.y + (h != null ? (h * pos.translateY) / 100 : 0)
  const [l, t, r, b] = box.insets ?? TEXTBOX_INSETS_PX
  const style: Record<string, string> = {
    left: `${x - (geom.stripLeft ?? geom.marginLeft)}px`,
    padding: `${t}px ${r}px ${b}px ${l}px`,
  }
  const tx: number = w == null ? pos.translateX : 0
  let ty: number = h == null ? pos.translateY : 0
  const paraRel = box.posVRel === 'paragraph'
  if (paraRel && box.posYPx == null && (box.posV === 'center' || box.posV === 'bottom')) {
    // aligned to the anchor paragraph's box (the host: that paragraph, else the strip)
    style.top = box.posV === 'center' ? '50%' : '100%'
    ty = box.posV === 'center' ? -50 : -100
  } else if (paraRel && box.posYPx != null && (kind === 'footer' || box.anchorPara != null)) {
    // paragraph-relative offsets measure from the host's top: the anchor
    // paragraph when the box shares one, else the strip itself
    style.top = `${box.posYPx}px`
  } else if (kind === 'footer') {
    // pinned by its bottom edge: an auto-height box grows upward, so the
    // anchor point (top / center / bottom of the box) is restored by
    // shifting it down by the rest of its height
    style.bottom = `${geom.pageH - geom.footerDist - (y + (h ?? 0))}px`
    if (h == null) ty = 100 + pos.translateY
  } else {
    style.top = `${y - geom.headerStripTop}px`
  }
  if (w != null) style.width = `${w}px`
  if (h != null) {
    style.height = `${h}px`
    // Word clips a fixed-size box's overflow (a wrapped tail past the declared height is not drawn)
    if (!box.autofit) style.overflow = 'hidden'
  }
  if (tx || ty) style.transform = `translate(${tx}%, ${ty}%)`
  if (box.vAlign === 'center') style.justifyContent = 'center'
  else if (box.vAlign === 'bottom') style.justifyContent = 'flex-end'
  // behindDoc: under the body text (a strip hosting boxes forms no stacking context)
  if (box.behind) style.zIndex = '-1'
  return style
}

/** paragraph element a box hangs off (paragraph-relative offsets), else null = the strip hosts it */
export function hfBoxAnchorEl<T>(box: HfTextBox, paraEls: T[]): T | null {
  return box.posVRel === 'paragraph' && box.anchorPara != null
    ? (paraEls[box.anchorPara] ?? null)
    : null
}

/** host classes: wrap="none" boxes keep their paragraphs on one line (the
 *  paragraphs' own pre-wrap would otherwise override a host-level nowrap) */
export function hfTextBoxClass(box: HfTextBox): string {
  return box.nowrap ? 'page-hf-textbox page-hf-nowrap' : 'page-hf-textbox'
}

export function makeGapHfEl(opts: {
  kind: 'header' | 'footer'
  value: HeaderFooter
  images?: HfImage[]
  /** page number shown for the PAGE marker (may be a section-formatted string) */
  pageNo: number | string
  /** total page count shown for the NUMPAGES marker */
  pageTotal: number
  /** page geometry: floating textboxes render at their anchor position (absent: stacked) */
  geom?: HfStripGeom
}): HTMLElement {
  const { kind, value, images, pageNo, pageTotal, geom } = opts
  const legacyHash = hfUsesLegacyHash(value)
  const display = (text: string) => {
    const substituted = text
      .replaceAll(TOTAL_PAGES_MARK, String(pageTotal))
      .replaceAll(PAGE_MARK, String(pageNo))
    return legacyHash ? substituted.replace('#', String(pageNo)) : substituted
  }
  const wrap = document.createElement('div')
  wrap.className = `page-hf page-hf-${kind} page-gap-hf`
  wrap.contentEditable = 'false'
  // Word sizes strip lines by their runs — the strip base is only the fallback
  // for runs without w:sz. When every text run declares a size, the strip font
  // (the em behind every strut/empty-line line-height) takes it, shrink-only
  // (growth would over-reserve push-down, prod_008/091). SAS prod_043: an
  // all-8pt header measured at the 10.5pt strut pushed the body top ~1px and
  // cost every two-column page its 42nd grid row (+1 page).
  const strutPt = hfDeclaredStrutPt(parasOf(value))
  if (strutPt != null) wrap.style.fontSize = `min(${strutPt}pt, var(--hf-default-fs, 10.5pt))`
  if (images && images.length > 0) {
    const imgWrap = document.createElement('div')
    imgWrap.className = 'page-hf-images'
    if (images[0].align === 'right') imgWrap.style.justifyContent = 'flex-end'
    else if (images[0].align === 'center') imgWrap.style.justifyContent = 'center'
    for (const img of images) {
      imgWrap.append(hfImgNode(img))
    }
    wrap.append(imgWrap)
  }
  // consecutive paragraphs of one floating textbox share a positioned host
  let boxHost: HTMLElement | null = null
  let boxId: number | undefined
  const paraEls: HTMLElement[] = []
  const paras = parasOf(value)
  const spacing = hfStackedSpacingPx(paras)
  for (const [index, para] of paras.entries()) {
    if (para.box?.id !== boxId) {
      boxId = para.box?.id
      boxHost = null
      const css = para.box && geom ? hfTextBoxStyle(para.box, kind, geom) : null
      if (css) {
        boxHost = document.createElement('div')
        boxHost.className = hfTextBoxClass(para.box!)
        Object.assign(boxHost.style, css)
        wrap.classList.add('page-hf-has-boxes')
        // a box sharing its paragraph with text hangs off that paragraph
        const anchor = hfBoxAnchorEl(para.box!, paraEls)
        if (anchor) anchor.classList.add('page-hf-anchor')
        ;(anchor ?? wrap).append(boxHost)
      }
    }
    const host = boxHost ?? wrap
    const p = document.createElement('div')
    paraEls[index] = p
    p.className = 'page-hf-para'
    const lh = hfParaLineHeightCss(para)
    if (lh) p.style.lineHeight = lh
    // floating-box content: displayed in the strip, but the push-down probe
    // excludes it (Word draws it at the anchor, off the strip flow)
    if (para.boxAnchored) p.classList.add('page-hf-box-anchored')
    const sp = spacing[index]
    if (sp.top) p.style.marginTop = `${sp.top}px`
    if (sp.bottom) p.style.marginBottom = `${sp.bottom}px`
    if (para.cells) {
      // layout-table row: flex columns sized by the cell widths
      p.classList.add('page-hf-row')
      Object.assign(p.style, hfRowStyle(para.row))
      for (const cell of para.cells) {
        const cellEl = document.createElement('div')
        cellEl.className = 'page-hf-cell'
        const geom = hfCellGeometry(cell)
        Object.assign(cellEl.style, geom.style)
        /* document content colors (w:shd / borders); the dark page reads the --dk-* twins */
        if (cell.fill) {
          cellEl.style.backgroundColor = `#${cell.fill}`
          setDkBackground(cellEl, `#${cell.fill}`)
        }
        for (const [side, css] of Object.entries(geom.borders)) {
          setDkBorder(cellEl, side as 't' | 'b' | 'l' | 'r', css)
        }
        const spansOf = (runs: Run[], host: HTMLElement) => {
          for (const run of runs) {
            if (run.image?.rule) {
              const rule = document.createElement('span')
              rule.className = INLINE_RULE_CLASS
              rule.style.cssText = inlineRuleDecls({
                ...run.image.rule,
                sizeHalfPoints: run.sizeHalfPoints,
              }).join(';')
              host.append(rule)
              if (!run.text) continue
            } else if (run.image) {
              const im = hfImgNode(run.image)
              im.classList.add('page-hf-cell-img')
              host.append(im)
              if (!run.text) continue
            }
            const span = document.createElement('span')
            span.textContent = display(run.text)
            applyRunStyle(span, run)
            host.append(span)
          }
        }
        // one block line per cell paragraph (Word stacks them; a lone empty
        // paragraph still reserves its line inside a shaded cell)
        const paras = cell.paras.length > 0 ? cell.paras : [[]]
        paras.forEach((runs, k) => {
          const props = cell.paraProps?.[k]
          const tabLines = hfCellTabLines(runs, props, geom, para.row, display)
          if (!tabLines) {
            const paraEl = document.createElement('div')
            paraEl.className = 'page-hf-cell-para'
            Object.assign(paraEl.style, hfCellParaStyle(props))
            if (runs.length === 0) paraEl.textContent = ' '
            spansOf(runs, paraEl)
            cellEl.append(paraEl)
            return
          }
          tabLines.forEach((line, m) => {
            const paraEl = document.createElement('div')
            paraEl.className = 'page-hf-cell-para page-hf-tabbed'
            const pos = { first: m === 0, last: m === tabLines.length - 1 }
            Object.assign(paraEl.style, hfCellParaStyle(props, pos), { textAlign: 'left' })
            if (line.minHeightPt) paraEl.style.minHeight = `${line.minHeightPt}pt`
            spansOf(line.lead, paraEl)
            for (const seg of line.segments) {
              const segEl = document.createElement('span')
              segEl.className = `page-hf-tabseg page-hf-tabseg-${seg.anchor}`
              Object.assign(segEl.style, hfCellSegStyle(seg))
              spansOf(seg.runs, segEl)
              paraEl.append(segEl)
            }
            cellEl.append(paraEl)
          })
        })
        p.append(cellEl)
      }
      host.append(p)
      continue
    }
    if (para.bidi) p.style.direction = 'rtl'
    if (para.align) {
      p.style.textAlign =
        para.align === 'left' || para.align === 'center' || para.align === 'right'
          ? para.align
          : 'justify'
    }
    // frame placement wins over the paragraph's own jc (the frame is narrower
    // than the column; its x position is what the reader sees)
    if (para.frameXAlign) {
      p.classList.add('page-hf-frame')
      p.style.textAlign = para.frameXAlign
    }
    /* document content colors (w:shd / w:pBdr); mirrors the body paragraph path,
       including the --dk-* twins the dark page reads */
    if (para.shadingFill) {
      p.style.backgroundColor = `#${para.shadingFill}`
      setDkBackground(p, `#${para.shadingFill}`)
    }
    if (para.borders) {
      const line = (side: 't' | 'b' | 'l' | 'r') => paraBorderCss(para.borderLines?.[side])
      if (para.borders.includes('t')) {
        p.style.borderTop = line('t')
        setDkBorder(p, 't', line('t'))
      }
      if (para.borders.includes('b')) {
        p.style.borderBottom = line('b')
        setDkBorder(p, 'b', line('b'))
      }
      if (para.borders.includes('l')) {
        p.style.borderLeft = line('l')
        setDkBorder(p, 'l', line('l'))
      }
      if (para.borders.includes('r')) {
        p.style.borderRight = line('r')
        setDkBorder(p, 'r', line('r'))
      }
      p.style.padding = '1px 4px'
    }
    const tabbed = hfTabSegments(para, display)
    if (tabbed) {
      p.classList.add('page-hf-tabbed')
      if (tabbed.minHeightPt) p.style.minHeight = `${tabbed.minHeightPt}pt`
      // tab layout happens in left-aligned space; w:jc becomes an explicit shift
      p.style.textAlign = 'left'
      const leadIndent = hfLeadIndentCss(tabbed)
      if (leadIndent) p.style.textIndent = leadIndent
      for (const run of tabbed.lead) {
        const span = document.createElement('span')
        span.textContent = display(run.text)
        applyRunStyle(span, run)
        p.append(span)
      }
      for (const seg of tabbed.segments) {
        const segEl = document.createElement('span')
        segEl.className = `page-hf-tabseg page-hf-tabseg-${seg.anchor}`
        segEl.style.left = hfSegLeftCss(seg, tabbed)
        for (const run of seg.runs) {
          const span = document.createElement('span')
          span.textContent = display(run.text)
          applyRunStyle(span, run)
          segEl.append(span)
        }
        p.append(segEl)
      }
      host.append(p)
      continue
    }
    if (para.runs.length === 0) p.textContent = ' '
    for (const run of para.runs) {
      const span = document.createElement('span')
      span.textContent = display(run.text)
      applyRunStyle(span, run)
      p.append(span)
    }
    host.append(p)
  }
  return wrap
}

/**
 * Word stacks w:spacing before/after of consecutive strip paragraphs (no
 * CSS-style collapsing): each paragraph's top margin carries the previous
 * paragraph's after, a layout-table row consumes the carry ahead of itself,
 * the last flow paragraph keeps its own after; anchored-box paragraphs are
 * drawn at the anchor and stay out of the flow. Shared by the canvas gap strip
 * and the preview/export strip so both reserve the same height.
 */
export function hfStackedSpacingPx(
  paras: readonly HfParagraph[],
): Array<{ top?: number; bottom?: number }> {
  const out: Array<{ top?: number; bottom?: number }> = paras.map(() => ({}))
  let carry = 0
  let lastFlow = -1
  paras.forEach((para, i) => {
    if (para.boxAnchored) return
    if (para.cells) {
      if (carry > 0) out[i].top = carry / 15
      carry = 0
      return
    }
    const before = carry + (para.spaceBefore ?? 0)
    if (before > 0) out[i].top = before / 15
    carry = para.spaceAfter ?? 0
    lastFlow = i
  })
  if (lastFlow >= 0 && carry > 0) out[lastFlow].bottom = carry / 15
  return out
}
