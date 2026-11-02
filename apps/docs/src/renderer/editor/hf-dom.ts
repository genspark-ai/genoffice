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
import { cssRunFontFamily, estimateHfHeight, runLetterSpacingCss } from '../line-metrics'
import {
  anchorXBand,
  anchorYBand,
  resolveAnchor,
  xAlignOf,
  xOriginOf,
  yOriginOf,
  type AnchorFrame,
} from './anchor-geometry'
import { borderCssStyle, borderDrawnPx, cellDiagonalCss, isDrawnBorder } from './border-metrics'
import { effectivePPrStyle, effectiveRPrStyle } from './effective-css'
import { setDkBackground, setDkBackgroundImage, setDkBorder, setDkColor } from './dark-page'
import { INLINE_RULE_CLASS, inlineRuleDecls } from './inline-rule'
import { hfParasOf } from './hf-text'
import { currentRevisionDisplay, runHiddenIn } from './revision-view'
import { borderMergeFlags, type ParaBorderAttrs } from './para-border-merge'
import {
  fromHfParagraph,
  hfLineFactorNumber,
  hfParaLineHeightCss,
  hfRunFace,
  toStyleRecord,
} from './para-layout-spec'
export { hfRunFace } from './para-layout-spec'
import {
  paraBorderCss,
  paraBorderLinesWithPad,
  paraBorderPadding,
  paraBorderShadowDecls,
} from './para-border-css'
export {
  PARA_BORDER_SHADOW,
  paraBorderCss,
  paraBorderLinesWithPad,
  paraBorderPadding,
  paraBorderPaddingDecls,
  paraBorderShadowDecls,
  type ParaBorderPadding,
  type ParaBorderSideLines,
} from './para-border-css'

/**
 * Plain-DOM header/footer rendering for the canvas page gaps (M4 always-on
 * pagination). Mirrors HeaderFooterArea's read-only display: same classes,
 * same PAGE_MARK / NUMPAGES substitution — but built imperatively because page gaps
 * are ProseMirror widget decorations, not React children.
 */

export type ParaBorderMergeFlags = { suppressTop: boolean; suppressBottom: boolean }

/** Word border groups over the strip paragraphs (identical pBdr + shading merge; rows split) */
export function hfBorderMergeFlags(paras: readonly HfParagraph[]): ParaBorderMergeFlags[] {
  return borderMergeFlags(
    paras.map((para): ParaBorderAttrs =>
      para.cells
        ? { break: true }
        : {
            borders: para.borders ?? null,
            borderLines: para.borderLines ? JSON.stringify(para.borderLines) : null,
            borderPad: para.borderPad ? JSON.stringify(para.borderPad) : null,
            shadingFill: para.shadingFill ?? null,
            shadingDisplay: para.shadingDisplay ?? null,
          },
    ),
  )
}

/** border lines, padding and shading-box shadow of one strip paragraph (camelCase style) */
export function hfParaBorderStyle(
  para: HfParagraph,
  flags?: ParaBorderMergeFlags,
): { style: Record<string, string>; borders: Partial<Record<'t' | 'b' | 'l' | 'r', string>> } {
  const style: Record<string, string> = {}
  const borders: Partial<Record<'t' | 'b' | 'l' | 'r', string>> = {}
  const lines = paraBorderLinesWithPad(para.borderLines, para.borderPad)
  const drawn = (para.borders ?? '')
    .split('')
    .filter((ch) => !(ch === 't' && flags?.suppressTop) && !(ch === 'b' && flags?.suppressBottom))
    .join('')
  const prop = { t: 'borderTop', b: 'borderBottom', l: 'borderLeft', r: 'borderRight' } as const
  for (const ch of ['t', 'b', 'l', 'r'] as const) {
    if (drawn.includes(ch)) style[prop[ch]] = borders[ch] = paraBorderCss(lines[ch] ?? undefined)
  }
  const padding = paraBorderPadding(drawn, lines)
  if (flags?.suppressTop) delete padding.paddingTop
  if (flags?.suppressBottom) delete padding.paddingBottom
  Object.assign(style, padding)
  const shdBg = para.shadingDisplay ?? para.shadingFill
  if (shdBg) style['--pbdr-fill'] = `#${shdBg}`
  for (const decl of paraBorderShadowDecls(lines)) {
    const i = decl.indexOf(':')
    const k = decl.slice(0, i)
    style[k.startsWith('--') ? k : 'boxShadow'] = decl.slice(i + 1)
  }
  return { style, borders }
}

/** One OOXML border → CSS border value; 'none' means explicitly borderless */
export function borderLineCss(
  b: { style: string; szEighths?: number; color?: string } | undefined | null,
): string | null {
  if (!b) return null
  if (!isDrawnBorder(b)) return 'none'
  const color = b.color && b.color !== 'auto' ? `#${b.color}` : '#000'
  return `${borderDrawnPx(b)}px ${borderCssStyle(b.style)} ${color}`
}

const px = (twips: number) => `${(twips / TWIPS_PER_PX).toFixed(1)}px`

/** Word's default cell margins (twips) for sides the table / cell leaves undeclared */
const HF_CELL_MAR = { top: 0, right: 108, bottom: 0, left: 108 } as const

/** cell box declarations (camelCase), its border CSS per side (callers add the
 *  dark-page twins) and the text-area width tabs lay out against */
export interface HfCellGeometry {
  style: Record<string, string>
  borders: Partial<Record<'t' | 'b' | 'l' | 'r', string>>
  /** w:tl2br / w:tr2bl as background-image layers (callers add the dark-page twin) */
  diagonals?: string
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
  const diagonals = (['tl2br', 'tr2bl'] as const)
    .map((d) => cellDiagonalCss(cell.borders?.[d], d))
    .filter((v): v is string => v !== null)
    .join(',')
  if (diagonals) style.backgroundImage = diagonals
  return {
    style,
    borders,
    ...(diagonals ? { diagonals } : {}),
    ...(textWidthPx != null ? { textWidthPx } : {}),
  }
}

/** strip geometry a positioned layout-table row (w:tblpPr / w:jc) is placed against, px */
export interface HfRowGeom {
  kind: 'header' | 'footer'
  /** strip column width */
  columnPx?: number
  /** page x of the strip's left edge */
  stripLeftPx?: number
  pageWPx?: number
  marginRightPx?: number
  marginTopPx?: number
  /** header strip top on the page: origin of page-anchored w:tblpY */
  headerStripTopPx?: number
}

export function hfRowGeomOf(
  kind: 'header' | 'footer',
  geom: HfStripGeom | undefined,
  columnPx = hfStripColumnPx(geom),
): HfRowGeom {
  if (!geom) return { kind, ...(columnPx != null ? { columnPx } : {}) }
  return {
    kind,
    ...(columnPx != null ? { columnPx } : {}),
    stripLeftPx: geom.stripLeft ?? geom.marginLeft,
    pageWPx: geom.pageW,
    marginRightPx: geom.marginRight,
    marginTopPx: geom.marginTop,
    headerStripTopPx: geom.headerStripTop,
  }
}

/** Left edge (px from the strip column's left edge) of a w:tblpPr-positioned
 *  row; null = the row stays in the strip flow (text-anchored, or the geometry
 *  it positions against is unknown). Word measures w:tblpX from the page edge
 *  or the margin per w:horzAnchor; w:tblpXSpec centring needs the table width. */
export function hfRowFloatLeftPx(row: HfTableRow, geom?: HfRowGeom): number | null {
  const f = row.float
  if (!f || !geom) return null
  const x = (f.xTwips ?? 0) / TWIPS_PER_PX
  const width = row.widthTwips ? row.widthTwips / TWIPS_PER_PX : undefined
  const align = xAlignOf(f.xSpec)
  const fromPage = f.horzAnchor === 'page'
  // numeric X (and the left / inside keyword): from the page edge or the column edge
  if (!align || align === 'left') {
    if (!fromPage) return x
    return geom.stripLeftPx == null ? null : x - geom.stripLeftPx
  }
  if (width == null) return null
  if (fromPage ? geom.stripLeftPx == null : geom.columnPx == null) return null
  const a = resolveAnchor(
    { relH: f.horzAnchor ?? null, xSpec: f.xSpec, widthPx: width },
    hfRowFrame(geom),
    HF_ROW_CTX,
  )
  return a.pageX == null ? null : a.pageX - (geom.stripLeftPx ?? 0)
}

const HF_ROW_CTX = { kind: 'hf', surface: 'static' } as const

/** page frame of a strip row: the margin band is the strip column (inherited widths included) */
function hfRowFrame(geom: HfRowGeom): AnchorFrame {
  const stripLeft = geom.stripLeftPx ?? 0
  return {
    ...(geom.pageWPx != null ? { page: { w: geom.pageWPx, h: Number.POSITIVE_INFINITY } } : {}),
    ...(geom.columnPx != null ? { column: { x: stripLeft, w: geom.columnPx } } : {}),
    // the margin band is the strip column (inherited widths included)
    ...(geom.marginTopPx != null
      ? {
          margins: {
            left: stripLeft,
            right:
              geom.pageWPx != null && geom.columnPx != null
                ? geom.pageWPx - stripLeft - geom.columnPx
                : 0,
            top: geom.marginTopPx,
            bottom: 0,
          },
        }
      : {}),
  }
}

/** page-anchored w:tblpY of the first row of a header table, as a gap below the
 *  strip top (px); the footer strip grows upward, so its tables stay in flow */
function hfRowFloatTopPx(row: HfTableRow, geom: HfRowGeom | undefined, first: boolean): number {
  const f = row.float
  if (!f || !geom || !first || geom.kind !== 'header' || f.yTwips == null) return 0
  if (geom.headerStripTopPx == null) return 0
  if (f.vertAnchor !== 'page' && f.vertAnchor !== 'margin') return 0
  const band = anchorYBand(f.vertAnchor, hfRowFrame(geom))
  if (!band) return 0
  return Math.max(0, band[0] + f.yTwips / TWIPS_PER_PX - geom.headerStripTopPx)
}

/** left edge and width (px, strip-column relative) a row's box occupies */
function hfRowExtentPx(row: HfTableRow, geom?: HfRowGeom): { left: number; width?: number } {
  const width = row.widthTwips ? row.widthTwips / TWIPS_PER_PX : undefined
  const floated = hfRowFloatLeftPx(row, geom)
  if (floated != null) return { left: floated, width }
  if (row.align && width != null && geom?.columnPx != null) {
    const free = geom.columnPx - width
    return { left: row.align === 'center' ? free / 2 : free, width }
  }
  return { left: (row.indentTwips ?? 0) / TWIPS_PER_PX, width }
}

/** how far layout-table rows hang into the start margin (negative tblInd, w:tblpPr past the margin), px */
export function hfIndentOverhangPx(paras: readonly HfParagraph[], geom?: HfRowGeom): number {
  return Math.max(0, ...paras.map((p) => (p.row ? -hfRowExtentPx(p.row, geom).left : 0)))
}

/** how far layout-table rows run past the strip column's right edge (an
 *  over-wide positioned table), px; capped at the paper edge like tab tails */
export function hfRowRightOverhangPx(paras: readonly HfParagraph[], geom?: HfRowGeom): number {
  if (!geom || geom.columnPx == null) return 0
  let over = 0
  for (const p of paras) {
    if (!p.row) continue
    const { left, width } = hfRowExtentPx(p.row, geom)
    if (width != null) over = Math.max(over, left + width - geom.columnPx)
  }
  return Math.min(over, geom.marginRightPx ?? over)
}

/** Layout-table row box: left offset of the table edge (w:tblInd, or the
 *  w:tblpPr / w:jc placement), the declared w:trHeight and a page-anchored
 *  table's gap below the strip top */
export function hfRowStyle(
  row: HfTableRow | undefined,
  geom?: HfRowGeom,
  first = false,
): Record<string, string> {
  const style: Record<string, string> = {}
  if (!row) return style
  const floated = hfRowFloatLeftPx(row, geom)
  if (floated != null) {
    style.marginLeft = `${floated.toFixed(1)}px`
  } else if (row.align) {
    // Word aligns the table in the column; an over-wide one overflows both sides
    style.justifyContent = row.align === 'center' ? 'center' : 'flex-end'
  } else if (row.indentTwips) {
    // a negative indent hangs the table into the margin (the strip widens its
    // clip box by --hf-ind-over) and the auto-width row grows by the overhang
    style[row.bidiVisual ? 'marginRight' : 'marginLeft'] = px(row.indentTwips)
  }
  const top = hfRowFloatTopPx(row, geom, first)
  if (top > 0.5) style.marginTop = `${top.toFixed(1)}px`
  if (row.bidiVisual) style.flexDirection = 'row-reverse'
  if (row.heightTwips) {
    if (row.heightRule === 'exact') {
      style.height = px(row.heightTwips)
      style.overflow = 'hidden'
    } else style.minHeight = px(row.heightTwips)
  }
  return style
}

/** w:ind of a stacked strip paragraph (twips): logical margins (RTL-aware), so a w:pBdr border
 *  keeps its own padding and draws at the indent edge like Word. Tabbed paragraphs
 *  stay full-width: their stops are laid out from the column edge (hfTabSegments). */
export function hfParaIndentStyle(para: {
  indentLeft?: number
  indentRight?: number
  indentFirstLine?: number
  runs?: Array<{ text: string }>
}): Record<string, string> {
  if (para.runs?.some((r) => r.text.includes('\t'))) return {}
  return effectivePPrStyle(
    {
      indentLeftTwips: para.indentLeft || null,
      indentRightTwips: para.indentRight || null,
      indentFirstLineTwips: para.indentFirstLine || null,
    },
    { length: px },
  )
}

/** cell paragraph box; a tab-wrapped paragraph spans several line boxes, so
 *  first-line-only (indent, space before) and last-line-only (space after)
 *  props land on the matching line */
export function hfCellParaBox(
  props: HfCellParaProps | undefined,
  line: { first: boolean; last: boolean } = { first: true, last: true },
  runs?: Run[],
  row?: Pick<HfTableRow, 'additiveSpacing'>,
): { style: Record<string, string>; classes: string[] } {
  const box = fromHfParagraph(props ?? {}, runs ?? [])
  const style = toStyleRecord(box)
  if (!props) return { style, classes: box.classes }
  if (props.align) {
    style.textAlign =
      props.align === 'left' || props.align === 'center' || props.align === 'right'
        ? props.align
        : 'justify'
  }
  Object.assign(
    style,
    effectivePPrStyle(
      {
        spaceBeforeTwips: props.spaceBefore || null,
        spaceAfterTwips: props.spaceAfter || null,
        indentLeftTwips: props.indentLeft || null,
        indentFirstLineTwips: props.indentFirstLine || null,
        additiveSpacing: row?.additiveSpacing === true,
      },
      { length: px, indentLeftAsPadding: true, line },
    ),
  )
  return { style, classes: box.classes }
}

export function hfCellParaStyle(...args: Parameters<typeof hfCellParaBox>): Record<string, string> {
  return hfCellParaBox(...args).style
}

/** Object.assign for style records that may carry custom properties */
function assignStyle(el: HTMLElement, style: Record<string, string>): void {
  for (const [k, v] of Object.entries(style)) {
    if (k.startsWith('--')) el.style.setProperty(k, v)
    else el.style[k as 'lineHeight'] = v
  }
}

function applyRunStyle(span: HTMLElement, run: Run): void {
  if (run.ins) span.classList.add('doc-ins')
  if (run.del) span.classList.add('doc-del')
  const cs = run.effective?.cs ? run.effective : undefined
  const face = hfRunFace(run)
  const bold = face.bold ?? cs?.bold
  const italic = face.italic ?? cs?.italic
  assignStyle(
    span,
    effectiveRPrStyle(
      {
        sizeHalfPoints: face.sizeHalfPoints ?? cs?.sizeHalfPoints ?? null,
        fontFamily: face.font || face.fontAscii ? (face.fontAscii ?? face.font ?? null) : null,
        bold: bold === true,
      },
      { boldWeight: '600', fontFamilyCss: () => cssRunFontFamily(face.fontAscii, face.font) },
    ),
  )
  if (bold === false) span.style.fontWeight = 'normal'
  if (italic) span.style.fontStyle = 'italic'
  else if (italic === false) span.style.fontStyle = 'normal'
  const deco = [run.underline && 'underline', run.strike && 'line-through'].filter(Boolean)
  if (deco.length > 0) span.style.textDecoration = deco.join(' ')
  if (run.color) {
    span.style.color = `#${run.color}`
    setDkColor(span, run.color) // dark-page twin; the authored color stays the declaration
  }
  const letterSpacing = runLetterSpacingCss(run)
  if (letterSpacing) span.style.letterSpacing = letterSpacing
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
  /** laid-out right edge of a px-placed segment (column px) */
  endPx?: number
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
  /** text after the last tab that does not fit before the column edge: Word
   *  wraps it like body text, so it stays in flow (gapPx = the tab's width)
   *  and its continuation lines start at the paragraph indent */
  flow?: { runs: Run[]; gapPx: number }
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
    const face = hfRunFace(run)
    const sizePt = face.sizeHalfPoints ? face.sizeHalfPoints / 2 : basePt
    if (hfMeasureCtx) {
      const family =
        face.font || face.fontAscii ? cssRunFontFamily(face.fontAscii, face.font) : 'sans-serif'
      // Detached-canvas font parsing cannot resolve var() (no element style context —
      // the whole assignment would be ignored): inline each var's fallback chain
      const canvasFamily = family.replace(/var\(--[\w-]+,([^)]*)\)/g, '$1')
      hfMeasureCtx.font = `${face.italic ? 'italic ' : ''}${face.bold ? '600 ' : ''}${(sizePt * 4) / 3}px ${canvasFamily}`
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
/** the paragraph's runs as the current Markup view shows them, ptabAligns
 *  re-indexed to the tabs that remain (it indexes every authored tab) */
function shownTabPara<T extends Pick<HfParagraph, 'runs' | 'ptabAligns'>>(para: T): T {
  const mode = currentRevisionDisplay()
  if (mode === 'all' || !para.runs.some((r) => runHiddenIn(mode, r))) return para
  const runs: Run[] = []
  const ptabAligns: HfParagraph['ptabAligns'] = []
  let tab = 0
  for (const run of para.runs) {
    const tabs = run.text.split('\t').length - 1
    if (!runHiddenIn(mode, run)) {
      runs.push(run)
      for (let k = 0; k < tabs; k++) ptabAligns.push(para.ptabAligns?.[tab + k])
    }
    tab += tabs
  }
  return { ...para, runs, ...(para.ptabAligns ? { ptabAligns } : {}) }
}

export function hfTabSegments(
  rawPara: Pick<HfParagraph, 'runs' | 'tabStops' | 'ptabAligns' | 'align'>,
  display: (text: string) => string = (t) => t,
  cell?: HfCellTabOpts,
  columnPx?: number,
): HfTabLayout | null {
  const para = shownTabPara(rawPara)
  const shown = para.runs
  if (!shown.some((r) => r.text.includes('\t'))) return null
  const chunks: Run[][] = [[]]
  for (const run of shown) {
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
  let flow: HfTabLayout['flow']
  let x = (cell?.startPx ?? 0) + hfRunsWidthPx(chunks[0], display)
  for (let k = 1; k < chunks.length; k++) {
    const runs = chunks[k].filter((r) => r.text !== '' || r.image)
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
    // the tail of a strip paragraph overflowing the column (an explicit stop may
    // sit in the margin and carry its text that far) wraps like body text
    if (
      !cell &&
      columnPx != null &&
      k === chunks.length - 1 &&
      segments.length === 0 &&
      !usedPct &&
      runs.length > 0 &&
      placed + segW > Math.max(columnPx, stop?.x ?? 0) + 0.5
    ) {
      flow = { runs, gapPx: Math.max(0, placed - x) }
      break
    }
    // only an explicit stop may carry text past the column (a default-grid tab wraps in Word)
    if (runs.length > 0) {
      segments.push({
        runs,
        left: { px: placed },
        anchor: 'left',
        ...(stop ? { endPx: placed + segW } : {}),
      })
    }
    x = placed + segW
  }
  // absolutely positioned segments add no flow height: an oversized run after a
  // tab (or an empty lead) would collapse to the strip's min-height and clip
  const maxHalfPoints = Math.max(0, ...shown.map((r) => r.sizeHalfPoints ?? 0))
  const faceFactor = hfLineFactorNumber(shown) ?? 1.3
  const align = para.align
  return {
    lead: chunks[0].filter((r) => r.text !== '' || r.image),
    segments,
    ...(maxHalfPoints > 0 ? { minHeightPt: (maxHalfPoints / 2) * faceFactor } : {}),
    ...(!usedPct && !cell && !flow && (align === 'center' || align === 'right')
      ? { shift: { align, lineEndPx: x } }
      : {}),
    ...(rest?.length ? { rest } : {}),
    ...(flow ? { flow } : {}),
  }
}

/** Tab layouts of a strip paragraph, one per w:br line: every line lays its
 *  tabs out from the column edge again, w:ptab alignments keep their overall
 *  order across lines. Null when the paragraph has no tab. */
export function hfTabLines(
  rawPara: Pick<HfParagraph, 'runs' | 'tabStops' | 'ptabAligns' | 'align'>,
  display: (text: string) => string = (t) => t,
  columnPx?: number,
): HfTabLayout[] | null {
  const para = shownTabPara(rawPara)
  const shown = para.runs
  if (!shown.some((r) => r.text.includes('\t'))) return null
  if (!shown.some((r) => r.text.includes('\n'))) {
    const single = hfTabSegments(para, display, undefined, columnPx)
    return single ? [single] : null
  }
  const lines: Run[][] = [[]]
  for (const run of shown) {
    const pieces = run.text.split('\n')
    lines[lines.length - 1].push({ ...run, text: pieces[0] })
    for (const piece of pieces.slice(1)) lines.push([{ ...run, text: piece }])
  }
  let tabsBefore = 0
  return lines.map((runs) => {
    const line = { ...para, runs, ptabAligns: para.ptabAligns?.slice(tabsBefore) }
    tabsBefore += runs.reduce((n, r) => n + r.text.split('\t').length - 1, 0)
    return (
      hfTabSegments(line, display, undefined, columnPx) ?? {
        lead: runs.filter((r) => r.text !== '' || r.image),
        segments: [],
      }
    )
  })
}

/** tallest picture inside the line's positioned segments (px): out of flow, so
 *  the line needs an in-flow strut of that height to keep the shared baseline */
export function hfTabSegImageHeightPx(layout: HfTabLayout): number {
  return Math.max(
    0,
    ...layout.segments.flatMap((s) =>
      s.runs.map((r) => (r.image?.rule ? 0 : (r.image?.heightPx ?? 0))),
    ),
  )
}

export function hfSegHasImage(seg: HfTabSegment): boolean {
  return seg.runs.some((r) => !!r.image && !r.image.rule)
}

/** text-column width (px) strip tabs lay out against */
export function hfStripColumnPx(geom: HfStripGeom | undefined): number | undefined {
  return geom ? (geom.columnW ?? geom.pageW - geom.marginLeft - geom.marginRight) : undefined
}

/** How far (px) tab segments run past the text column: Word keeps a stop set
 *  beyond the right margin and lets its text reach into the margin (a footer's
 *  right stop at 9360 twips on a 9070-twip column), so the strip's clip box
 *  widens by that much, never past the paper edge */
export function hfTabOverflowPx(lines: HfTabLayout[], geom: HfStripGeom | undefined): number {
  if (!geom) return 0
  const column = hfStripColumnPx(geom)!
  let end = 0
  for (const line of lines) for (const seg of line.segments) end = Math.max(end, seg.endPx ?? 0)
  return Math.min(Math.max(0, end - column), Math.max(0, geom.marginRight))
}

/** a tab line whose lead shows nothing (footer starting with a tab) has no
 *  in-flow line box: without a strut the line collapses to zero height and
 *  the strip's overflow clips the positioned segments */
export function hfTabLeadNeedsStrut(layout: HfTabLayout): boolean {
  return !layout.lead.some((r) => r.text.trim() || r.image)
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
        : {
            ...p,
            runs: p.runs
              .map((r) => ({ ...r, text: strip(r.text) }))
              .filter((r) => r.text || r.image),
          },
    )
    // dedicated page-number paragraphs go away entirely; table rows and user-typed blank lines stay
    .filter((p, i) => p.cells != null || p.runs.length > 0 || value.paras![i].runs.length === 0)
  const text = paras
    .map((p) =>
      [...p.runs, ...(p.cells?.flatMap((c) => c.paras.flat()) ?? [])].map((r) => r.text).join(''),
    )
    .join('')
  if (!text && !paras.some((p) => p.runs.some((r) => r.image))) {
    return { ...value, text: '', paras: undefined, pageNumber: false }
  }
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

/**
 * Rendered strip heights keyed by content signature (pagination recomputes per
 * page). Least-recently-used eviction: a document with hundreds of sections
 * carries one distinct strip per section, and clearing the whole map once it
 * filled up made every pass re-probe every strip (a forced layout each).
 */
const HF_HEIGHT_CACHE_MAX = 2000
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
    // layout containment: the probe swaps its children per measure, and the
    // forced layout that reads them back must not re-lay out the document
    host.style.cssText =
      'position:absolute;left:-99999px;top:0;visibility:hidden;pointer-events:none;contain:layout style'
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

/** the mounted doc-scoped stylesheets (`<style data-doc-css>`, rendered by App) */
function mountedDocCssStyles(): HTMLStyleElement[] {
  // a tag-name collection is served from the DOM's cache; an attribute
  // selector walks every node (a third of a pass on a 300k-node document)
  const out: HTMLStyleElement[] = []
  for (const style of document.getElementsByTagName('style')) {
    if (style.hasAttribute('data-doc-css')) out.push(style)
  }
  return out
}

/**
 * djb2 over the mounted doc-scoped stylesheets, memoized per style element.
 * A pagination pass asks for the signature once per section strip (hundreds
 * of times on a long document) while the stylesheets never change in between:
 * re-joining megabytes of CSS on every call was most of a pass on large files.
 * An unchanged text node hands back the same string, so the per-element
 * comparison is a pointer check on the hot path.
 */
let mountedCssMemo: { parts: string[]; hash: string } | null = null
function mountedDocCssSig(): string {
  const styles = mountedDocCssStyles()
  const memo = mountedCssMemo
  if (memo && memo.parts.length === styles.length) {
    let same = true
    for (let i = 0; i < styles.length; i++) {
      if ((styles[i].textContent ?? '') !== memo.parts[i]) {
        same = false
        break
      }
    }
    if (same) return memo.hash
  }
  const parts = Array.from(styles, (s) => s.textContent ?? '')
  const text = parts.join('\0')
  let h = 5381
  for (let i = 0; i < text.length; i++) h = ((h << 5) + h + text.charCodeAt(i)) | 0
  const hash = `${text.length}:${h}`
  mountedCssMemo = { parts, hash }
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
/** probe geometry a floating header table's w:tblpY resolves against; only
 *  strips that carry one are keyed on it so plain strips still share entries */
export function hfProbeGeomKey(
  value: HeaderFooter | null,
  geom: { marginTopPx: number; headerDistPx: number } | undefined,
): string {
  if (!geom || !value?.paras?.some((p) => p.row?.float?.yTwips != null)) return ''
  return `${Math.round(geom.headerDistPx)}:${Math.round(geom.marginTopPx)}`
}

/** header geometry the estimate's anchored-drawing floor measures in */
export interface HfProbeGeom {
  marginTopPx: number
  headerDistPx: number
  /** page x of the strip column: page-relative drawings' side room */
  marginLeftPx?: number
}

const isStripGeom = (g: HfStripGeom | HfProbeGeom): g is HfStripGeom => 'pageW' in g

/** the estimate floor is a header rule (the body sits under a header's wrapTopAndBottom drawing) */
function hfEstimateGeom(
  kind: 'header' | 'footer',
  geom: HfStripGeom | HfProbeGeom | undefined,
): HfProbeGeom | undefined {
  if (!geom) return undefined
  if (!isStripGeom(geom)) return geom
  if (kind !== 'header') return undefined
  return {
    marginTopPx: geom.marginTop,
    headerDistPx: geom.headerStripTop,
    marginLeftPx: geom.stripLeft ?? geom.marginLeft,
  }
}

/** flow height of a strip element: the sentinel after the last flow block sits
 *  under the collapsed bottom margin and ignores floats hanging below the text
 *  (a side-wrapped drawing below the header lines reserves nothing, L260) */
function hfFlowHeightPx(el: HTMLElement): number {
  const end = document.createElement('div')
  end.className = 'page-hf-flow-end'
  el.append(end)
  const h = end.getBoundingClientRect().top - el.getBoundingClientRect().top
  end.remove()
  return h > 0 ? h : el.getBoundingClientRect().height
}

/** the part as the current Markup view shows it: view-hidden tracked runs leave
 *  the strip model so they neither reserve height nor place tab segments */
export function hfValueInView(value: HeaderFooter | null): HeaderFooter | null {
  const mode = currentRevisionDisplay()
  if (mode === 'all' || !value?.paras) return value
  const shown = (runs: Run[]) => runs.filter((r) => !runHiddenIn(mode, r))
  return {
    ...value,
    paras: value.paras.map((p) => ({
      ...shownTabPara(p),
      ...(p.cells ? { cells: p.cells.map((c) => ({ ...c, paras: c.paras.map(shown) })) } : {}),
    })),
  }
}

export function hfReservedHeightPx(
  kind: 'header' | 'footer',
  rawValue: HeaderFooter | null,
  contentWidthPx: number,
  images?: HfImage[] | null,
  /** full strip geometry lays the probe out like the canvas gap strip (floating
   *  drawings as flow shapers at their page position); the header-only pair
   *  keeps the positioned-row and floor maths of callers without a page */
  geom?: HfStripGeom | HfProbeGeom,
): number {
  const value = hfValueInView(rawValue)
  const estGeom = hfEstimateGeom(kind, geom)
  const est = estimateHfHeight(value, contentWidthPx, images, estGeom)
  const stripGeom = geom && isStripGeom(geom) ? { ...geom, columnW: contentWidthPx } : undefined
  const probeImages = stripGeom ? (images ?? []) : (images ?? []).filter((img) => !img.floating)
  const inline = probeImages.filter((img) => !img.floating)
  // blank paragraphs hold their line and spacing like any other strip paragraph
  // (Word probes 2026-09-06/17): the rendered strip decides for them too
  const blankParas = !!value?.paras?.some((p) => !p.boxAnchored && !p.cells)
  if ((!hfHasVisibleContent(value, inline) && !blankParas) || contentWidthPx <= 0) return est
  if (typeof document === 'undefined' || !document.body) return est
  // cell-run images can carry megabytes of base64: key on the dataUrl length only
  const noDataUrl = (k: string, v: unknown) => (k === 'dataUrl' ? String(v).length : v)
  const geomKey = stripGeom
    ? `s${[stripGeom.marginLeft, stripGeom.marginRight, stripGeom.marginTop, stripGeom.marginBottom, stripGeom.headerStripTop, stripGeom.footerDist, stripGeom.pageW, stripGeom.pageH, stripGeom.stripLeft ?? ''].map((v) => Math.round(Number(v))).join(':')}`
    : hfProbeGeomKey(value, estGeom)
  const key =
    `${kind}|${Math.round(contentWidthPx)}|${geomKey}|${mountedDocCssSig()}|${hfProbeFontEpoch}|` +
    `${JSON.stringify(value, noDataUrl)}|` +
    probeImages
      .map((im) => `${JSON.stringify({ ...im, dataUrl: im.dataUrl.length, wordArt: undefined })}`)
      .join(',')
  let dom = hfHeightCache.get(key)
  if (dom !== undefined) {
    // re-insert: Map iteration order is insertion order, so the oldest entry is first
    hfHeightCache.delete(key)
    hfHeightCache.set(key, dom)
  } else {
    const host = hfProbeHost()
    if (!host) return est
    const el = makeGapHfEl({
      kind,
      value: value ?? { text: '' },
      images: probeImages,
      pageNo: 88,
      pageTotal: 88,
      columnPx: contentWidthPx,
      ...(stripGeom
        ? { geom: stripGeom }
        : estGeom
          ? {
              rowGeom: {
                kind,
                columnPx: contentWidthPx,
                marginTopPx: estGeom.marginTopPx,
                headerStripTopPx: estGeom.headerDistPx,
              },
            }
          : {}),
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
    dom = hfFlowHeightPx(el)
    host.replaceChildren()
    if (hfHeightCache.size >= HF_HEIGHT_CACHE_MAX) {
      const oldest = hfHeightCache.keys().next().value
      if (oldest !== undefined) hfHeightCache.delete(oldest)
    }
    hfHeightCache.set(key, dom)
  }
  if (dom <= 0) return est
  return Math.max(dom, estimateHfHeight(null, contentWidthPx, images, estGeom))
}

/** a side-wrapped drawing's box in strip coordinates, as the float that shapes the strip's lines */
export interface HfObstacle {
  side: 'left' | 'right'
  top: number
  height: number
  /** float width from the strip edge on `side` to the drawing's far edge */
  width: number
}

const SIDE_WRAPS = new Set(['square', 'tight', 'through'])
/** narrowest gap that still takes a word (the body wrap engine's sliver, L462 approximation) */
const HF_SIDE_WRAP_SLIVER_PX = 36

/**
 * Side-wrapped (wrapSquare / tight / through) floating drawings of a strip as
 * flow shapers: Word lays header paragraphs around them with the body rule
 * (L462, "the next word fits"; a blocked line lands at the drawing's bottom),
 * and never pushes the body below them (L260). A footer strip grows upward,
 * so only its paragraph-relative anchors have a known strip position.
 */
export function hfSideWrapObstacles(
  kind: 'header' | 'footer',
  images: readonly HfImage[] | undefined,
  geom: HfStripGeom,
): HfObstacle[] {
  const column = hfStripColumnPx(geom)
  if (!images?.length || column === undefined || column <= 0) return []
  const stripLeft = geom.stripLeft ?? geom.marginLeft
  const out: HfObstacle[] = []
  for (const img of images) {
    if (!img.floating || img.behind || !img.wrap || !SIDE_WRAPS.has(img.wrap)) continue
    if (hfImageHangsOnPara(img)) continue
    const w = img.widthPx ?? 0
    const h = Math.max(1, img.heightPx ?? 0)
    if (w <= 0) continue
    if (kind === 'footer' && (img.posVRel !== 'paragraph' || img.posYPx == null)) continue
    const pos = hfFloatPagePos(img, {
      pageW: geom.pageW,
      pageH: geom.pageH,
      marginLeft: geom.marginLeft,
      marginRight: geom.marginRight,
      marginTop: geom.marginTop,
      marginBottom: geom.marginBottom,
      headerDist: geom.headerStripTop,
      sectMarginTop: geom.marginTop,
      sectMarginBottom: geom.marginBottom,
      ...(kind === 'footer' ? { paraOriginY: 0 } : {}),
    })
    const left = pos.x + (pos.translateX / 100) * w - stripLeft
    const right = left + w
    let top = pos.y + (pos.translateY / 100) * h - (kind === 'header' ? geom.headerStripTop : 0)
    let height = h
    if (top < 0) {
      height += top
      top = 0
    }
    if (height <= 0 || right <= 0 || left >= column) continue
    // text order: the gap on the left takes the line when it holds a word,
    // otherwise the line moves past the drawing (probe 2026-10-29 CASE02/04)
    const side: HfObstacle['side'] = left >= HF_SIDE_WRAP_SLIVER_PX ? 'right' : 'left'
    const width = Math.min(column, side === 'left' ? right : column - left)
    out.push({ side, top, height, width })
  }
  return out.sort((a, b) => a.top - b.top)
}

/** inline styles of the obstacle floats, in document order: a zero-width
 *  spacer float lowers each box to its top (a margin would be part of the
 *  float's margin box and shorten the lines above it); each side stacks under
 *  its own `clear`, so a spacer bridges from the previous box's bottom */
export function hfObstacleStyles(obstacles: readonly HfObstacle[]): Array<Record<string, string>> {
  const bottom = { left: 0, right: 0 }
  const out: Array<Record<string, string>> = []
  for (const o of obstacles) {
    const gap = o.top - bottom[o.side]
    if (gap > 0)
      out.push({ float: o.side, clear: o.side, width: '0px', height: `${gap.toFixed(1)}px` })
    bottom[o.side] = Math.max(bottom[o.side], o.top) + o.height
    out.push({
      float: o.side,
      clear: o.side,
      width: `${o.width.toFixed(1)}px`,
      height: `${o.height.toFixed(1)}px`,
    })
  }
  return out
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
  /** footer images: the footer strip top (page bottom - footerDist - strip height) replaces headerDist as that origin */
  paraOriginY?: number
  /** raw sectPr top margin (px, before push-down): origin wrapped margin-relative images reserve from */
  sectMarginTop: number
  /** raw sectPr bottom margin (px, before footer push-up): the bottomMargin band; absent = marginBottom */
  sectMarginBottom?: number
  /** left edge of this page on the shared canvas paper (differing-width documents center narrower pages) */
  paperX?: number
  /** inset of the lead host's content box from the paper edge (the canvas padding: the first section's odd-page left margin); absent = marginLeft */
  hostLeft?: number
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
  const vRel = img.posVRel
  // paragraph/margin-rel wrapped images use the same origins the body
  // push-down estimate measures from (header strip top / raw sectPr margin),
  // so the image never chases the pushed-down margin; watermarks keep the
  // effective margin
  const wrapped = img.wrap && img.wrap !== 'none' && !img.behind
  const frame: AnchorFrame = {
    page: { w: box.pageW, h: box.pageH },
    margins: {
      left: box.marginLeft,
      right: box.marginRight,
      top: img.posYPx != null && vRel === 'margin' && wrapped ? box.sectMarginTop : box.marginTop,
      bottom: box.marginBottom,
    },
    sectMargins: { top: box.sectMarginTop, bottom: box.sectMarginBottom ?? box.marginBottom },
    paragraph: { top: box.paraOriginY ?? box.headerDist },
  }
  // no relativeFrom (legacy VML) measures in the margin band
  const [bandL, bandR] = anchorXBand(img.posHRel ? xOriginOf(img.posHRel) : 'margin', frame)!
  let x: number
  let translateX: 0 | -50 | -100 = 0
  if (img.posXPx != null) {
    x = bandL + img.posXPx
  } else if (img.posH === 'center') {
    // legacy VML entries carry no rel and keep the page center
    x = img.posHRel ? (bandL + bandR) / 2 : box.pageW / 2
    translateX = -50
  } else if (img.posH === 'right') {
    x = bandR
    translateX = -100
  } else {
    x = bandL
  }
  // a paragraph-relative alignment keyword measures in the margin band
  const yOrigin = !vRel || (vRel === 'paragraph' && img.posYPx == null) ? 'margin' : yOriginOf(vRel)
  const [bandT, bandB] = anchorYBand(yOrigin, frame)!
  let y: number
  let translateY: 0 | -50 | -100 = 0
  if (img.posYPx != null) {
    y = bandT + img.posYPx
  } else if (img.posV === 'center') {
    y = vRel && vRel !== 'paragraph' ? (bandT + bandB) / 2 : box.pageH / 2
    translateY = -50
  } else if (img.posV === 'bottom') {
    y = bandB
    translateY = -100
  } else {
    y = bandT
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
 * v:imagedata gain/blacklevel as a CSS filter chain. Measured against Word's
 * rendering of its washout preset (gain 0.3, blacklevel 0.35): black maps to
 * 0.805 and everything above 1 - blacklevel is white, i.e. brightness adds
 * blacklevel first (clamped at white), then gain scales the distance from
 * white: out = 1 - gain * (1 - min(1, in + blacklevel)).
 * brightness(1/(1-b)) is that clamped add; invert-brightness-invert the scale.
 */
export function hfWashoutFilter(w: NonNullable<HfImage['washout']>): string {
  const b = Math.max(0, Math.min(0.99, w.blackLevel))
  const g = Math.max(0, Math.min(1, w.gain))
  const r = (v: number) => Math.round(v * 1000) / 1000
  return `brightness(${r(1 / (1 - b))}) invert(1) brightness(${r(g * (1 - b))}) invert(1)`
}

/** CSS transform of a floating header/footer shape: anchor translate, then the VML rotation about the box center */
export function hfFloatTransform(
  img: Pick<HfImage, 'rotationDeg'>,
  p: { translateX: number; translateY: number },
): string | undefined {
  const parts: string[] = []
  if (p.translateX || p.translateY) parts.push(`translate(${p.translateX}%, ${p.translateY}%)`)
  if (img.rotationDeg) parts.push(`rotate(${img.rotationDeg}deg)`)
  return parts.length > 0 ? parts.join(' ') : undefined
}

let inkCtx: CanvasRenderingContext2D | null | undefined
/**
 * WordArt watermark as inline SVG: the string's glyph ink box (canvas
 * measureText) is stretched onto the shape box, which is how Word's
 * textpath fitshape draws it (tall narrow "DRAFT" across a wide box).
 * Empty when the environment cannot measure text.
 */
export function wordArtSvgMarkup(img: HfImage): string {
  const wa = img.wordArt
  const w = img.widthPx ?? 0
  const h = img.heightPx ?? 0
  if (!wa || w <= 0 || h <= 0) return ''
  if (inkCtx === undefined) inkCtx = document.createElement('canvas').getContext('2d')
  if (!inkCtx || typeof inkCtx.measureText !== 'function') return ''
  const family = wa.fontFamily ? `"${wa.fontFamily.replace(/"/g, '')}", sans-serif` : 'sans-serif'
  const weight = wa.bold ? 'bold' : 'normal'
  const styleKw = wa.italic ? 'italic' : 'normal'
  inkCtx.font = `${styleKw} ${weight} 100px ${family}`
  const m = inkCtx.measureText(wa.text)
  const left = m.actualBoundingBoxLeft ?? 0
  const inkW = left + (m.actualBoundingBoxRight ?? m.width)
  const ascent = m.actualBoundingBoxAscent ?? 0
  const inkH = ascent + (m.actualBoundingBoxDescent ?? 0)
  if (!(inkW > 0) || !(inkH > 0)) return ''
  const r = (v: number) => Math.round(v * 1000) / 1000
  const esc = (t: string) =>
    t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">` +
    `<text transform="scale(${r(w / inkW)} ${r(h / inkH)}) translate(${r(left)} ${r(ascent)})"` +
    ` font-family=${JSON.stringify(esc(family))} font-size="100" font-weight="${weight}" font-style="${styleKw}"` +
    ` fill="#${wa.colorHex}" fill-opacity="${wa.opacity}" xml:space="preserve">${esc(wa.text)}</text></svg>`
  )
}

/** <img> for a header/footer picture; an a:srcRect crop becomes an
 *  overflow-hidden window over a scaled and offset image (body-path technique) */
/** run spans plus run pictures / rules, appended to host in run order */
function hfRunNodes(
  runs: Run[],
  host: HTMLElement,
  display: (text: string) => string,
  imgClass: string,
): void {
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
      im.classList.add(imgClass)
      host.append(im)
      if (!run.text) continue
    }
    const span = document.createElement('span')
    span.textContent = display(run.text)
    applyRunStyle(span, run)
    host.append(span)
  }
}

function hfTabGapNode(widthPx: number): HTMLElement {
  const gap = document.createElement('span')
  gap.className = 'page-hf-tab-gap'
  gap.style.width = `${widthPx.toFixed(1)}px`
  return gap
}

function hfTabStrutNode(heightPx: number): HTMLElement {
  const strut = document.createElement('span')
  strut.className = 'page-hf-tab-strut'
  strut.style.height = `${heightPx}px`
  return strut
}

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
  let el: HTMLElement
  if (img.wordArt) {
    el = document.createElement('span')
    el.innerHTML = wordArtSvgMarkup(img)
  } else {
    el = hfImgNode(img)
  }
  el.className = 'page-hf-float-img'
  if (!img.behind) el.style.zIndex = '1'
  // a crop wrapper carries an inline position:relative for the strip hosts;
  // the float path must stay absolute (inline style beats the class)
  el.style.position = 'absolute'
  const p = hfFloatPagePos(img, box)
  const paperX = (box.paperX ?? 0) + p.x
  if (host === 'gap') {
    el.style.left = `${paperX}px`
    // alignGapHfStrips re-anchors it to the paper: gap boxes don't start at the paper edge
    el.dataset.paperX = paperX.toFixed(1)
    el.style.top = `calc(100% + ${p.y - box.marginTop}px)`
  } else {
    el.style.left = `${paperX - (box.hostLeft ?? box.marginLeft)}px`
    el.style.top = `${p.y - box.marginTop}px`
  }
  const transform = hfFloatTransform(img, p)
  if (transform) el.style.transform = transform
  if (img.widthPx) el.style.width = `${img.widthPx}px`
  // hairline shapes (a 0.25pt rule) still get one device pixel, as in Word
  if (img.heightPx) el.style.height = `${Math.max(1, img.heightPx)}px`
  if (img.washout) el.style.filter = hfWashoutFilter(img.washout)
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
  /** where the header strip's top edge sits on the page: w:headerDist, also under a fixed top margin above it (probe 2026-10-30) */
  headerStripTop: number
  /** w:footerDist: the footer strip's bottom edge sits this far above the page bottom */
  footerDist: number
  /** page x of the strip's left edge; absent = the left margin (a centered strip on unequal side margins sits elsewhere) */
  stripLeft?: number
  /** column width the part is laid out at when it differs from the page's own
   *  (a part inherited from a section of another width keeps that width) */
  columnW?: number
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
    sectMarginBottom: geom.marginBottom,
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
  // the anchor-point translate composes ahead of the rotation: centering
  // shifts stay axis-aligned while the box itself turns about its centre
  const transforms = [
    tx || ty ? `translate(${tx}%, ${ty}%)` : '',
    box.rotDeg ? `rotate(${box.rotDeg}deg)` : '',
  ].filter(Boolean)
  if (transforms.length > 0) style.transform = transforms.join(' ')
  // the shape's own outline (a:ln / VML stroke): Word draws header shapes
  // with their authored stroke, not as bare text rectangles
  if (box.borderColor) {
    style.border = `${box.borderWidthPx ?? 1}px solid #${box.borderColor}`
    style.boxSizing = 'border-box'
  }
  if (box.vAlign === 'center') style.justifyContent = 'center'
  else if (box.vAlign === 'bottom') style.justifyContent = 'flex-end'
  // behindDoc: under the body text (a strip hosting boxes forms no stacking context)
  if (box.behind) style.zIndex = '-1'
  return style
}

/** floating image whose paragraph-relative offset hangs off a strip row (cell-anchored logos) */
export function hfImageHangsOnPara(img: HfImage): boolean {
  return (
    !!img.floating && img.anchorPara != null && img.posVRel === 'paragraph' && img.posYPx != null
  )
}

/** placement of such an image inside its (position: relative) anchor row */
export function hfAnchoredImgStyle(img: HfImage, geom?: HfStripGeom): Record<string, string> {
  const style: Record<string, string> = { position: 'absolute', top: `${img.posYPx ?? 0}px` }
  if (geom) {
    const pos = hfFloatPagePos(img, {
      pageW: geom.pageW,
      pageH: geom.pageH,
      marginLeft: geom.marginLeft,
      marginRight: geom.marginRight,
      marginTop: geom.marginTop,
      marginBottom: geom.marginBottom,
      headerDist: geom.headerStripTop,
      sectMarginTop: geom.marginTop,
      sectMarginBottom: geom.marginBottom,
    })
    style.left = `${pos.x - (geom.stripLeft ?? geom.marginLeft)}px`
    if (pos.translateX) style.transform = `translateX(${pos.translateX}%)`
  } else style.left = `${img.posXPx ?? 0}px`
  if (img.widthPx) style.width = `${img.widthPx}px`
  if (img.heightPx) style.height = `${Math.max(1, img.heightPx)}px`
  if (img.behind) style.zIndex = '-1'
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
  /** text-column width tab tails wrap at when no geom is given (the height probe) */
  columnPx?: number
  /** positioned-row geometry when only part of the page geometry is known (the height probe) */
  rowGeom?: HfRowGeom
}): HTMLElement {
  const { kind, images, pageNo, pageTotal, geom } = opts
  const value = hfValueInView(opts.value)!
  const columnPx = opts.columnPx ?? hfStripColumnPx(geom)
  const rowGeom = opts.rowGeom ?? hfRowGeomOf(kind, geom, columnPx)
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
  if (geom) {
    for (const style of hfObstacleStyles(hfSideWrapObstacles(kind, images, geom))) {
      const ob = document.createElement('div')
      ob.className = 'page-hf-obstacle'
      assignStyle(ob, style)
      wrap.append(ob)
    }
  }
  const stacked = (images ?? []).filter((img) => !img.floating)
  if (stacked.length > 0) {
    const imgWrap = document.createElement('div')
    imgWrap.className = 'page-hf-images'
    if (stacked[0].align === 'right') imgWrap.style.justifyContent = 'flex-end'
    else if (stacked[0].align === 'center') imgWrap.style.justifyContent = 'center'
    for (const img of stacked) {
      imgWrap.append(hfImgNode(img))
    }
    wrap.append(imgWrap)
  }
  // consecutive paragraphs of one floating textbox share a positioned host
  let boxHost: HTMLElement | null = null
  let boxId: number | undefined
  const paraEls: HTMLElement[] = []
  const paras = hfParasOf(value, images)
  const spacing = hfStackedSpacingPx(paras)
  const mergeFlags = hfBorderMergeFlags(paras)
  let tabOver = 0
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
    if (para.cells) {
      const lh = hfParaLineHeightCss(para)
      if (lh) p.style.lineHeight = lh
    } else {
      const box = fromHfParagraph(para)
      assignStyle(p, toStyleRecord(box))
      p.classList.add(...box.classes)
    }
    // floating-box content: displayed in the strip, but the push-down probe
    // excludes it (Word draws it at the anchor, off the strip flow)
    if (para.boxAnchored) p.classList.add('page-hf-box-anchored')
    const sp = spacing[index]
    if (sp.top) p.style.marginTop = `${sp.top}px`
    if (sp.bottom) p.style.marginBottom = `${sp.bottom}px`
    if (para.cells) {
      // layout-table row: flex columns sized by the cell widths
      p.classList.add('page-hf-row')
      Object.assign(p.style, hfRowStyle(para.row, rowGeom, index === 0))
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
        if (geom.diagonals) setDkBackgroundImage(cellEl, geom.diagonals)
        const spansOf = (runs: Run[], host: HTMLElement) =>
          hfRunNodes(runs, host, display, 'page-hf-cell-img')
        // one block line per cell paragraph (Word stacks them; a lone empty
        // paragraph still reserves its line inside a shaded cell)
        const paras = cell.paras.length > 0 ? cell.paras : [[]]
        paras.forEach((runs, k) => {
          const props = cell.paraProps?.[k]
          const tabLines = hfCellTabLines(runs, props, geom, para.row, display)
          if (!tabLines) {
            const paraEl = document.createElement('div')
            paraEl.className = 'page-hf-cell-para'
            const box = hfCellParaBox(props, undefined, runs, para.row)
            assignStyle(paraEl, box.style)
            paraEl.classList.add(...box.classes)
            if (runs.length === 0) paraEl.textContent = ' '
            spansOf(runs, paraEl)
            cellEl.append(paraEl)
            return
          }
          tabLines.forEach((line, m) => {
            const paraEl = document.createElement('div')
            paraEl.className = 'page-hf-cell-para page-hf-tabbed'
            const pos = { first: m === 0, last: m === tabLines.length - 1 }
            const box = hfCellParaBox(props, pos, runs, para.row)
            assignStyle(paraEl, box.style)
            paraEl.classList.add(...box.classes)
            paraEl.style.textAlign = 'left'
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
    Object.assign(p.style, hfParaIndentStyle(para))
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
    {
      const bs = hfParaBorderStyle(para, mergeFlags[index])
      assignStyle(p, bs.style)
      for (const [side, css] of Object.entries(bs.borders)) {
        setDkBorder(p, side as 't' | 'b' | 'l' | 'r', css)
      }
    }
    const tabLines = hfTabLines(para, display, columnPx)
    if (tabLines) {
      tabOver = Math.max(tabOver, hfTabOverflowPx(tabLines, geom))
      // a w:br paragraph stacks one positioned line per break inside the
      // paragraph block (which keeps the spacing and borders)
      const single = tabLines.length === 1
      for (const tabbed of tabLines) {
        const line = single ? p : document.createElement('div')
        line.classList.add('page-hf-tabbed')
        if (tabbed.minHeightPt) line.style.minHeight = `${tabbed.minHeightPt}pt`
        // tab layout happens in left-aligned space; w:jc becomes an explicit
        // shift (a wrapping tail keeps the paragraph's own jc per line)
        if (tabbed.flow) line.classList.add('page-hf-tabbed-flow')
        else line.style.textAlign = 'left'
        const leadIndent = hfLeadIndentCss(tabbed)
        if (leadIndent) line.style.textIndent = leadIndent
        if (hfTabLeadNeedsStrut(tabbed)) line.append('\u200b')
        const segImgH = hfTabSegImageHeightPx(tabbed)
        if (segImgH > 0) line.append(hfTabStrutNode(segImgH))
        hfRunNodes(tabbed.lead, line, display, 'page-hf-run-img')
        if (tabbed.flow) {
          if (tabbed.flow.gapPx > 0) line.append(hfTabGapNode(tabbed.flow.gapPx))
          hfRunNodes(tabbed.flow.runs, line, display, 'page-hf-run-img')
        }
        for (const seg of tabbed.segments) {
          const segEl = document.createElement('span')
          segEl.className = `page-hf-tabseg page-hf-tabseg-${seg.anchor}`
          if (hfSegHasImage(seg)) segEl.classList.add('page-hf-tabseg-img')
          segEl.style.left = hfSegLeftCss(seg, tabbed)
          hfRunNodes(seg.runs, segEl, display, 'page-hf-run-img')
          line.append(segEl)
        }
        if (!single) p.append(line)
      }
      host.append(p)
      continue
    }
    if (para.runs.length === 0) p.textContent = ' '
    hfRunNodes(para.runs, p, display, 'page-hf-run-img')
    host.append(p)
  }
  const rightOver = Math.max(tabOver, hfRowRightOverhangPx(paras, rowGeom))
  if (rightOver > 0) wrap.style.setProperty('--hf-tab-over', `${rightOver.toFixed(1)}px`)
  const indOver = hfIndentOverhangPx(paras, rowGeom)
  if (indOver > 0) wrap.style.setProperty('--hf-ind-over', `${indOver.toFixed(1)}px`)
  for (const img of (images ?? []).filter(hfImageHangsOnPara)) {
    const anchor = paraEls[img.anchorPara!]
    if (!anchor) continue
    anchor.classList.add('page-hf-anchor')
    const el = hfImgNode(img)
    assignStyle(el, hfAnchoredImgStyle(img, geom))
    anchor.append(el)
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
