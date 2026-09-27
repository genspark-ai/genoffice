import {
  mergePPrFormat,
  parseLazyMediaUrl,
  setPPrChange,
  stripPPrChange,
  ommlToLatex,
  patchFieldParagraphXml,
  applyImageWrap,
  applyImageZOrder,
  patchImageParagraphXml,
  patchMathTokens,
  patchTableCellTexts,
  type CellParaPatch,
  type CellTextsPatch,
  patchDrawingExtent,
  patchTextboxSizes,
  patchShapeStyles,
  type ShapeStylePatch,
  type TextboxSizePatch,
  patchTextboxParas,
  generateTableModelXml,
  symbolGlyph,
  symbolPuaChar,
  type Block,
  type ChartDisplay,
  type ChartPatch,
  type ChartSeriesPatch,
  type FieldDisplay,
  type FieldTextPatch,
  type FormulaDisplay,
  type GeneratedBlock,
  type ImageWrap,
  type NewChart,
  type NewImage,
  type ParaFormat,
  type Run,
  type SaveBlock,
  type SdtShell,
  type SectionInfo,
  type TableCell,
  type TableModel,
  type TableParagraph,
  type TextboxDisplay,
  type TextboxParaPatch,
  type TextboxParasPatchSet,
} from '@genoffice/docx-engine'
import { t } from '../i18n/locale'
import { symbolFontCovers } from '../font-check'
import {
  canvasMetrics,
  charScaleEm,
  maxWordWidthPx,
  textHasComplexScript,
  type FontMetricsProvider,
  wordKerns,
} from '../line-metrics'
import { firstStrongDir } from './direction'
import { inlineMathML } from './equation'
import { isStraightLineKind } from './shape-svg'
import { borderDrawnPx, borderTotalPt, collapsedEdgePx } from './border-metrics'
import { parseRunBorderAttr } from './run-border'
import { parseTextOutlineAttr } from './text-outline'
import { charScaleXAttr, parseGlowAttr, parseTextEffectAttr } from './text-effects'

/** minimal ProseMirror JSON shapes */
export interface PmMark {
  type: string
  attrs?: Record<string, unknown>
}
export interface PmNode {
  type: string
  attrs?: Record<string, unknown>
  content?: PmNode[]
  text?: string
  marks?: PmMark[]
}

// ---- Block[] -> ProseMirror doc ----

/** document-level layout switches for the display models built here */
export interface PmDocOptions {
  /** settings.xml compatibilityMode < 15: w:tblInd measures to the cell text, so
   *  the table border sits one left cell margin further out (Word 2013+ measures
   *  to the border) */
  legacyTableIndent?: boolean
  /** rPrDefault size and Latin face: runs inheriting them carry neither */
  defaultRun?: RunDefaults
}

export interface RunDefaults {
  sizeHalfPoints?: number
  font?: string
}

export function pmDocOptions(parsed: {
  compatibilityMode?: number
  docDefaults?: { sizeHalfPoints?: number; asciiFont?: string }
}): PmDocOptions {
  const { sizeHalfPoints, asciiFont } = parsed.docDefaults ?? {}
  return {
    legacyTableIndent: (parsed.compatibilityMode ?? 0) < 15,
    defaultRun: {
      ...(sizeHalfPoints ? { sizeHalfPoints } : {}),
      ...(asciiFont ? { font: asciiFont } : {}),
    },
  }
}

export function blocksToPmDoc(
  blocks: Block[],
  sections?: SectionInfo[],
  options: PmDocOptions = {},
): PmNode {
  const content: PmNode[] = []
  for (const block of blocks) {
    if (block.hidden) continue
    content.push(
      blockToPmNode(
        block,
        sectionRowCapTwips(sections, block.docxIndex),
        sectionWidthBudget(sections, block.docxIndex),
        options,
      ),
    )
  }
  if (content.length === 0) {
    content.push({
      type: 'docParagraph',
      attrs: { docxIndex: null, styleId: null, aiChanged: false },
    })
  }
  return { type: 'doc', content }
}

/** Word lays out RTL-script / run-rtl paragraphs with an RTL base direction even
 * without w:bidi (HTML-converted docs omit it). The first strong character wins
 * (dir=auto semantics); run w:rtl only decides weak-only text. Render-only: never
 * written back, and never subject to the explicit-bidi jc left/right swap. */
export function inferredBidi(format: ParaFormat | undefined, runs: Run[] | undefined): boolean {
  if (format?.bidi || !runs?.length) return false
  const strong = firstStrongDir(runs.map((r) => r.text).join(''))
  if (strong) return strong === 'rtl'
  return runs.some((r) => r.rtl)
}

function formatAttrs(format: ParaFormat | undefined, runs?: Run[]): Record<string, unknown> {
  return {
    align: format?.align ?? null,
    lineSpacing: format?.lineSpacing ?? null,
    lineRule: format?.lineRule ?? null,
    lineRawTwips: format?.lineRawTwips ?? null,
    indentLeft: format?.indentLeft ?? null,
    indentRight: format?.indentRight ?? null,
    indentFirstLine: format?.indentFirstLine ?? null,
    spaceBefore: format?.spaceBefore ?? null,
    spaceAfter: format?.spaceAfter ?? null,
    spaceBeforeAuto: format?.spaceBeforeAuto ?? null,
    spaceAfterAuto: format?.spaceAfterAuto ?? null,
    contextualSpacing: format?.contextualSpacing ?? null,
    keepNext: format?.keepNext ?? null,
    keepLines: format?.keepLines ?? null,
    widowControl: format?.widowControl ?? null,
    suppressLineNumbers: format?.suppressLineNumbers ?? null,
    pageBreakBefore: format?.pageBreakBefore ?? false,
    shadingFill: format?.shadingFill ?? null,
    shadingDisplay: format?.shadingDisplay ?? null,
    shadingClear: format?.shadingClear ?? null,
    borderReset: format?.borderReset ?? null,
    borders: format?.borders ?? null,
    borderLines: format?.borderLines ? JSON.stringify(format.borderLines) : null,
    tabStops: format?.tabStops ? JSON.stringify(format.tabStops) : null,
    dropCap: format?.dropCap ? JSON.stringify(format.dropCap) : null,
    frame: format?.frame ? JSON.stringify(format.frame) : null,
    frameBox: format?.frameBox ? JSON.stringify(format.frameBox) : null,
    textDirection: format?.textDirection ?? null,
    bidi: format?.bidi ?? false,
    bidiInferred: inferredBidi(format, runs),
    autoSpace: format?.autoSpace ?? null,
    wordWrap: format?.wordWrap ?? null,
    overflowPunct: format?.overflowPunct ?? null,
    eaLang: format?.eastAsiaLang ?? null,
    snapToGrid: format?.snapToGrid ?? null,
    emptyRunSize: format?.emptyRunSizeHalfPoints ?? null,
    emptyRunFont: format?.emptyRunFontFamily ?? null,
    markSize: format?.markSizeHalfPoints ?? null,
  }
}

/** Effective row-height ceiling: one page of section content (Word truncates taller rows within the page) */
function sectionRowCapTwips(
  sections: SectionInfo[] | undefined,
  docxIndex: number | null,
): number | null {
  if (!sections?.length || docxIndex == null) return null
  const sec =
    sections.find((s) => docxIndex >= s.firstBlockIndex && docxIndex <= s.lastBlockIndex) ??
    sections[sections.length - 1]
  const cap = sec.settings.pageHeight - sec.settings.marginTop - sec.settings.marginBottom
  return cap > 0 ? cap : null
}

/** table width budget: Word fixed layout lets over-wide tables overflow the
 * text column — left-aligned ones spill right up to the paper edge (avail =
 * left page margin to the paper's right edge), centered ones spill both
 * margins symmetrically (paper = full page width, the hard compression cap);
 * fit = text column width, the target autofit growth is reclaimed back to */
export interface TableWidthBudget {
  avail: number
  fit: number
  paper: number
}
function sectionWidthBudget(
  sections: SectionInfo[] | undefined,
  docxIndex: number | null,
): TableWidthBudget | null {
  if (!sections?.length || docxIndex == null) return null
  const sec =
    sections.find((s) => docxIndex >= s.firstBlockIndex && docxIndex <= s.lastBlockIndex) ??
    sections[sections.length - 1]
  const paper = sec.settings.pageWidth
  const avail = paper - sec.settings.marginLeft
  if (avail <= 0) return null
  const fit = avail - sec.settings.marginRight
  return { avail, fit: fit > 0 ? fit : avail, paper }
}

/** centered tables spill both margins, so their growth bound is the paper width */
function budgetAvail(model: TableModel, budget: TableWidthBudget): number {
  return model.align === 'center' ? budget.paper : budget.avail
}

/** Cap declared row heights (incl. nested tables); returns the input model when nothing exceeds the cap */
export function capTableRowHeights(model: TableModel, capTwips: number): TableModel {
  let changed = false
  const rowHeightsTwips = model.rowHeightsTwips?.map((h) => {
    if (h != null && h > capTwips) {
      changed = true
      return capTwips
    }
    return h
  })
  const rows = model.rows.map((row) =>
    row.map((cell) => {
      if (!cell.nestedTables?.length) return cell
      const nested = cell.nestedTables.map((nt) => capTableRowHeights(nt, capTwips))
      if (nested.every((nt, i) => nt === cell.nestedTables![i])) return cell
      changed = true
      return { ...cell, nestedTables: nested }
    }),
  )
  return changed ? { ...model, rows, rowHeightsTwips } : model
}

const DEFAULT_CELL_MAR = 108
/** compression floor for real columns (24px); sliver grid slots below it keep their width */
const MIN_COL_TWIPS = 360

function tableIndentTwips(model: TableModel): number {
  return model.align !== 'center' && model.align !== 'right' && (model.indentTwips ?? 0) > 0
    ? model.indentTwips!
    : 0
}

/** nested tables are bounded by their cell's content width (grid slice minus
 * side padding); without a usable grid the outer budget still bounds them */
function mapNestedTables(
  model: TableModel,
  widths: number[] | undefined,
  budget: number,
  transform: (nt: TableModel, availTwips: number) => TableModel,
): { rows: TableModel['rows']; changed: boolean } {
  let changed = false
  const rows = model.rows.map((row) => {
    let column = 0
    let cellChanged = false
    const cells = row.map((cell) => {
      const start = column
      const span = cell.colSpan ?? 1
      column += span
      if (!cell.nestedTables?.length) return cell
      const cols = widths?.slice(start, start + span) ?? []
      const mar = cell.cellMarTwips ?? model.cellMarTwips
      const pad = (mar?.left ?? DEFAULT_CELL_MAR) + (mar?.right ?? DEFAULT_CELL_MAR)
      const cellWidth = cols.length === span ? cols.reduce((a, b) => a + b, 0) - pad : budget
      if (cellWidth <= 0) return cell
      const nested = cell.nestedTables.map((nt) => transform(nt, cellWidth))
      if (nested.every((nt, i) => nt === cell.nestedTables![i])) return cell
      cellChanged = true
      return { ...cell, nestedTables: nested }
    })
    if (!cellChanged) return row
    changed = true
    return cells
  })
  return { rows, changed }
}

function withColWidths(
  model: TableModel,
  rows: TableModel['rows'],
  widths: number[] | undefined,
): TableModel {
  if (!widths) return { ...model, rows }
  const total = widths.reduce((a, b) => a + b, 0)
  return {
    ...model,
    rows,
    colWidthsTwips: widths,
    ...(total > 0 ? { colWidthsPct: widths.map((w) => (w / total) * 100) } : {}),
  }
}

/**
 * Last-resort display clamp for grids wider than the hard cap (the paper width
 * for top-level tables, the cell content width for nested ones — Word lets
 * over-wide tables spill into the page margins, clipped at the paper, so
 * merely exceeding the text column must not narrow columns and wrap their
 * text onto extra lines). The cut comes proportionally out of the surplus
 * over each column's min-content, so a garbage gridCol absorbs it while real
 * columns keep their words unbroken. Fixed table layout grows past
 * width/max-width to the sum of absolute column widths, so an unclamped
 * garbage gridCol pushes content megapixels off page. Display-only like
 * capTableRowHeights: untouched tables still save their original bytes, and
 * structural rebuilds already regenerate the grid from these display widths
 * (same as user column resizes, which are clamped too).
 */
export function clampTableColWidths(
  model: TableModel,
  capTwips: number,
  metrics: FontMetricsProvider | null = canvasMetrics(),
): TableModel {
  const budget = capTwips - tableIndentTwips(model)
  let widths = model.colWidthsTwips
  let widthsChanged = false
  if (widths && budget > 0 && widths.reduce((a, b) => a + b, 0) > budget) {
    const mins = minContentColTwips(model, widths.length, metrics).map((m, i) =>
      Math.min(Math.max(m, MIN_COL_TWIPS), budget, widths![i]),
    )
    const overflow = widths.reduce((a, b) => a + b, 0) - budget
    const surplus = widths.map((w, i) => w - mins[i])
    const totalSurplus = surplus.reduce((a, b) => a + b, 0)
    if (totalSurplus > 0) {
      const take = Math.min(overflow, totalSurplus)
      widths = widths.map((w, i) => w - Math.round((surplus[i] / totalSurplus) * take))
    }
    // min-contents alone can exceed the cap: compress proportionally
    const sum = widths.reduce((a, b) => a + b, 0)
    if (sum > budget) {
      const scale = budget / sum
      widths = widths.map((w) => Math.max(1, Math.round(w * scale)))
    }
    widthsChanged = true
  }
  const { rows, changed: rowsChanged } = mapNestedTables(model, widths, budget, (nt, cap) =>
    clampTableColWidths(nt, cap, metrics),
  )
  if (!widthsChanged && !rowsChanged) return model
  return withColWidths(model, rows, widthsChanged ? widths : undefined)
}

/**
 * Browser fallback faces run wider than the 0.52em heuristic average (Arial /
 * Helvetica digits and most lowercase advance 0.556em, +7%); an under-floored
 * column re-shatters the very word the floor exists for (RFP sample: two-digit
 * row numbers broke one digit per line; Word grants that column ~11% over the
 * estimate). Measured widths only need rounding headroom.
 */
const MIN_CONTENT_SLACK = 1.08
const MEASURED_SLACK = 1.02
/** collapsed half-borders plus px rounding eat ~1px per side of the measured word (CI cap tables shattered "Board") */
const MEASURED_EDGE_PX = 2

/** per grid column (twips): widest unbreakable word in single-span cells, plus side
 * padding; `bare` drops the slack (a grid Word laid out already holds its words) */
function minContentColTwips(
  model: TableModel,
  colCount: number,
  metrics: FontMetricsProvider | null,
  bare = false,
  defaults?: RunDefaults,
): number[] {
  const slack = bare ? 1 : metrics ? MEASURED_SLACK : MIN_CONTENT_SLACK
  const mins = new Array<number>(colCount).fill(0)
  for (const row of model.rows) {
    let column = 0
    for (const cell of row) {
      const start = column
      column += cell.colSpan ?? 1
      if ((cell.colSpan ?? 1) !== 1 || start >= colCount) continue
      // merged-away and vertical-text cells don't demand word width
      if (cell.vMerge === 'continue' || cell.textDirection) continue
      const mar = cell.cellMarTwips ?? model.cellMarTwips
      const pad = (mar?.left ?? DEFAULT_CELL_MAR) + (mar?.right ?? DEFAULT_CELL_MAR)
      for (const nt of cell.nestedTables ?? []) {
        const ntMin = nestedTableMinTwips(nt, metrics, defaults)
        if (ntMin > 0) mins[start] = Math.max(mins[start], ntMin + pad)
      }
      let wordPx = cellMaxWordPx(cell, metrics, defaults)
      if (wordPx <= 0) continue
      if (metrics && !bare) wordPx += MEASURED_EDGE_PX
      mins[start] = Math.max(mins[start], Math.ceil(wordPx * slack * 15) + pad)
    }
  }
  return mins
}

/** the autofit reclaim pass never compresses a column below its min-content
 * or MIN_COL_TWIPS; sliver slots already under that keep their width */
function reclaimFloorTwips(width: number, minContent: number): number {
  return Math.max(minContent, Math.min(width, MIN_COL_TWIPS))
}

/** w:tblCellSpacing: the table spans its w:tblGrid and the (n + 1) gaps of 2s come
 * out of the columns as 2s each plus s more on the two outer ones (Word-saved grids:
 * 45/30/30/45 for s = 15, 60 for a single column) */
export function cellSpacingGridSharesTwips(colCount: number, spacingTwips: number): number[] {
  const n = Math.max(1, colCount)
  return Array.from(
    { length: n },
    (_, i) => 2 * spacingTwips + (i === 0 ? spacingTwips : 0) + (i === n - 1 ? spacingTwips : 0),
  )
}

/** side cell margins of a table's own cells (twips) */
function sideMarginTwips(model: TableModel): number {
  const mar = model.cellMarTwips
  return (mar?.left ?? DEFAULT_CELL_MAR) + (mar?.right ?? DEFAULT_CELL_MAR)
}

/** narrowest width a nested table can take inside its cell: fixed grids keep
 * their declared width, pct tables follow the cell, autofit ones shrink to
 * every column's min-content (an empty cell keeps its side margins). Word
 * straddles the outer vertical borders, so the box is half a border wider on
 * each side (a regression sample: eight nesting levels, each parent column
 * = nested grid + 216 + 10 twips for sz 4 borders). */
function nestedTableMinTwips(
  nt: TableModel,
  metrics: FontMetricsProvider | null,
  defaults?: RunDefaults,
): number {
  const widths = nt.colWidthsTwips
  if (!widths?.length || nt.widthPct) return 0
  const indent = tableIndentTwips(nt)
  const border = ((borderTotalPt(nt.borders?.left) + borderTotalPt(nt.borders?.right)) / 2) * 20
  if (nt.fixedLayout) return widths.reduce((a, b) => a + b, 0) + indent + border
  const mins = minContentColTwips(nt, widths.length, metrics, false, defaults)
  return (
    mins.reduce((sum, m, i) => sum + nestedFloorTwips(widths[i], m, sideMarginTwips(nt)), 0) +
    indent +
    border
  )
}

/** a nested table over its cell width shrinks each column to its min-content;
 * an empty column keeps its side margins (Word: tcW 360/360 holding a nested
 * table and "8" in a 544-twip cell -> 222/328), sliver grid slots keep their width */
function nestedFloorTwips(width: number, minContent: number, sideMar: number): number {
  return Math.max(minContent, Math.min(width, sideMar))
}

/** paragraph indent inside a cell eats line width like the word itself does;
 * list fallbacks mirror the .doc-li per-level CSS defaults (0.55in + 0.3in/level) */
function paraIndentPx(para: TableParagraph): number {
  const leftTw = para.indentLeft ?? (para.list ? 792 + 432 * (para.list.ilvl || 0) : 0)
  const firstTw = Math.max(para.indentFirstLine ?? 0, 0)
  return Math.max(leftTw + firstTw, 0) / 15
}

function cellMaxWordPx(
  cell: TableCell,
  metrics: FontMetricsProvider | null,
  defaults?: RunDefaults,
): number {
  const measure = (runs: Parameters<typeof maxWordWidthPx>[0]) =>
    metrics ? maxWordWidthPx(runs, metrics) : maxWordWidthPx(runs)
  let max = 0
  if (cell.richParas?.length) {
    for (const para of cell.richParas) {
      const runs: Parameters<typeof maxWordWidthPx>[0] = []
      for (const r of para.runs) {
        if (!r.text) continue
        // Latin words render in the ascii slot; `font` may be the East Asian face
        const family = r.fontAscii ?? r.font ?? defaults?.font
        const size = r.sizeHalfPoints ?? defaults?.sizeHalfPoints
        runs.push({
          text: r.text,
          ...(family ? { fontFamily: family } : {}),
          ...(size ? { sizeHalfPoints: size } : {}),
          ...(r.bold ? { bold: true } : {}),
          ...(r.italic ? { italic: true } : {}),
        })
      }
      if (!runs.length) continue
      max = Math.max(max, measure(runs) + paraIndentPx(para))
    }
  } else {
    for (const text of cell.paras) {
      if (!text) continue
      max = Math.max(
        max,
        measure([
          {
            text,
            ...(cell.bold ? { bold: true } : {}),
            ...(defaults?.font ? { fontFamily: defaults.font } : {}),
            ...(defaults?.sizeHalfPoints ? { sizeHalfPoints: defaults.sizeHalfPoints } : {}),
          },
        ]),
      )
    }
  }
  return max
}

/**
 * Word autofit (tblW auto/absent, no fixed tblLayout): a column grows to its
 * min-content — the widest unbreakable word — so narrow declared widths don't
 * shatter words char by char (heuristic estimate, see maxWordWidthPx). Growth
 * past the fit width, like a declared total already past it, is reclaimed
 * proportionally from columns with surplus over their own min-content; any
 * remainder spills and is bounded by clampTableColWidths. Display-only, like
 * the clamp.
 */
export function expandAutofitColWidths(
  model: TableModel,
  availTwips: number,
  fitTwips: number = availTwips,
  metrics: FontMetricsProvider | null = canvasMetrics(),
  nested = false,
  legacy = false,
  defaults?: RunDefaults,
): TableModel {
  const indent = tableIndentTwips(model)
  const budget = availTwips - indent
  let widths = model.colWidthsTwips
  let widthsChanged = false
  // a dxa tblW is only a preferred width: without w:tblLayout fixed Word still
  // grows a column to its widest unbreakable word and reclaims the growth from
  // the other columns' surplus, keeping the declared total (measured: a long
  // merge-field token grew its tcW 2835 -> 3000, neighbours shrank)
  const autofit = model.autoLayout || (!model.fixedLayout && !model.widthPct)
  // pct-width autofit tables: the min-content floor applies to the resolved
  // widths (Word never wraps below the widest word even when w:tblW is a tiny
  // percentage); growth converts the table to absolute widths for display.
  // pct resolves against the full text column — the render path draws
  // width:N% of the content box and applies the indent as a margin
  let resolvedPct = false
  if (autofit && model.widthPct && model.colWidthsPct?.length && budget > 0) {
    const tableW = (Math.min(fitTwips, availTwips) * model.widthPct) / 100
    if (tableW > 0) {
      widths = model.colWidthsPct.map((p) => Math.round((p / 100) * tableW))
      resolvedPct = true
    }
  }
  if (autofit && (resolvedPct || !model.widthPct) && widths?.length && budget > 0) {
    const mins = minContentColTwips(
      model,
      widths.length,
      metrics,
      !!model.layoutGrid,
      defaults,
    ).map((m) => Math.min(m, budget))
    const declared = widths.reduce((a, b) => a + b, 0)
    const fit = Math.min(fitTwips, availTwips)
    const target = fit - indent
    // Over-wide tblW-auto preferred widths compress to the column less the
    // signed table indent (Word probe 2026-09-17, compat 14 and 15: tblInd -108
    // widens the box by 108, +558 narrows it by 558, the right edge stays on
    // the margin). Compat < 15 measures the indent to the cell text, so the
    // border box hangs by the side cell margins - the left one is already in
    // the legacy indent shift, the right one is added here (centered tables
    // shift nothing and hang on both sides). A nested table sits inside the
    // cell padding, so it never hangs. An explicit dxa tblW is honoured as
    // long as the table stays on the paper (measured: landscape and portrait
    // forms 7-22% wider than the column, negative tblInd, drawn at full width).
    const sideMar = sideMarginTwips(model)
    const leftAligned = model.align !== 'center' && model.align !== 'right'
    const legacyHang = !legacy
      ? 0
      : leftAligned && !model.floatSide
        ? (model.cellMarTwips?.right ?? DEFAULT_CELL_MAR)
        : sideMar
    const box = fit - (leftAligned ? (model.indentTwips ?? 0) : 0) + legacyHang
    const hang = leftAligned ? Math.min(model.indentTwips ?? 0, 0) : 0
    const explicitDxa = !model.autoLayout && !nested && declared + hang <= budget
    const kept =
      resolvedPct || nested
        ? target
        : Math.max(target, explicitDxa ? declared : Math.min(declared, box))
    if (mins.some((m, i) => m > widths![i]) || (!resolvedPct && declared > kept)) {
      const grown = widths.map((w, i) => Math.max(w, mins[i]))
      const overflow = grown.reduce((a, b) => a + b, 0) - kept
      if (overflow > 0) {
        const floors = grown.map((w, i) =>
          nested ? nestedFloorTwips(w, mins[i], sideMar) : reclaimFloorTwips(w, mins[i]),
        )
        const surplus = grown.map((w, i) => Math.max(0, w - floors[i]))
        const totalSurplus = surplus.reduce((a, b) => a + b, 0)
        if (totalSurplus > 0) {
          const take = Math.min(overflow, totalSurplus)
          for (let i = 0; i < grown.length; i++)
            grown[i] -= Math.round((surplus[i] / totalSurplus) * take)
        }
        // min-contents alone can exceed the budget (one huge unbreakable
        // token): compress proportionally like Word instead of handing the
        // excess to clampTableColWidths, which zeroes trailing columns
        const sum = grown.reduce((a, b) => a + b, 0)
        if (sum > kept) {
          const scale = kept / sum
          for (let i = 0; i < grown.length; i++)
            grown[i] = Math.max(1, Math.round(grown[i] * scale))
        }
      }
      widths = grown
      widthsChanged = true
    }
  }
  const { rows, changed: rowsChanged } = mapNestedTables(model, widths, budget, (nt, avail) =>
    expandAutofitColWidths(nt, avail, avail, metrics, true, false, defaults),
  )
  if (!widthsChanged && !rowsChanged) return model
  const result = withColWidths(model, rows, widthsChanged ? widths : undefined)
  if (widthsChanged && resolvedPct) delete result.widthPct
  return result
}

/** a floating table hung on the left margin without an offset: Word places
 *  its cell text at the margin like an inline table (compat < 15 probe) */
function marginHungFloat(model: TableModel): boolean {
  const pos = model.floatPos
  return model.floatSide === 'left' && pos?.horzAnchor === 'margin' && !pos.xSpec && !pos.xTwips
}

/** legacy tblInd (measured to the cell text) as a border-edge indent; display-only */
export function legacyIndentTable(model: TableModel): TableModel {
  if (model.align === 'center' || model.align === 'right') return model
  if (model.floatSide && !marginHungFloat(model)) return model
  // measured to the first cell's text: its own w:tcMar, then the 2s gap and
  // the whole (uncollapsed) border a w:tblCellSpacing table draws before it
  const first = model.rows[0]?.[0]
  const spacing = model.cellSpacingTwips ?? 0
  const shift =
    (first?.cellMarTwips?.left ?? model.cellMarTwips?.left ?? DEFAULT_CELL_MAR) +
    (spacing > 0
      ? 2 * spacing + borderDrawnPx(first?.borders?.left ?? model.borders?.left) * 15
      : 0)
  if (shift <= 0) return model
  // w:tblpPr supersedes w:tblInd, so a float hangs by the cell margin alone
  return { ...model, indentTwips: (model.floatSide ? 0 : (model.indentTwips ?? 0)) - shift }
}

function displayTable(
  table: TableModel | null,
  rowCapTwips: number | null,
  budget: TableWidthBudget | null,
  options: PmDocOptions = {},
): TableModel | null {
  if (!table) return null
  let t = table
  if (options.legacyTableIndent) t = legacyIndentTable(t)
  if (rowCapTwips != null) t = capTableRowHeights(t, rowCapTwips)
  if (budget != null) {
    t = expandAutofitColWidths(
      t,
      budgetAvail(t, budget),
      budget.fit,
      undefined,
      false,
      options.legacyTableIndent ?? false,
      options.defaultRun,
    )
    t = clampTableColWidths(t, budget.paper)
  }
  return t
}

function blockToPmNode(
  block: Block,
  rowCapTwips: number | null = null,
  budget: TableWidthBudget | null = null,
  options: PmDocOptions = {},
): PmNode {
  if (block.altChunk && options.legacyTableIndent)
    options = { ...options, legacyTableIndent: false }
  switch (block.type) {
    case 'heading':
      return {
        type: 'docHeading',
        attrs: {
          docxIndex: block.docxIndex,
          styleId: block.styleId ?? null,
          aiChanged: false,
          level: block.level ?? 1,
          outlineOnly: block.outlineOnly ?? null,
          bookmarks: block.bookmarks ?? null,
          hiddenBookmarks: block.hiddenBookmarks ?? null,
          commentStarts: block.commentStarts ?? null,
          commentEnds: block.commentEnds ?? null,
          pPrChange: block.pPrChangeInfo ? JSON.stringify(block.pPrChangeInfo) : null,
          paraMarkDel: block.paraMarkDel ? JSON.stringify(block.paraMarkDel) : null,
          blockRevision: block.blockRevision ?? null,
          ...formatAttrs(block.format, block.runs),
        },
        content: runsToInline(block.runs ?? []),
      }
    case 'listItem':
      return {
        type: 'docListItem',
        attrs: {
          docxIndex: block.docxIndex,
          styleId: block.styleId ?? null,
          aiChanged: false,
          kind: block.list?.kind ?? 'bullet',
          numId: block.list?.numId ?? null,
          ilvl: block.list?.ilvl ?? 0,
          bookmarks: block.bookmarks ?? null,
          hiddenBookmarks: block.hiddenBookmarks ?? null,
          commentStarts: block.commentStarts ?? null,
          commentEnds: block.commentEnds ?? null,
          pPrChange: block.pPrChangeInfo ? JSON.stringify(block.pPrChangeInfo) : null,
          paraMarkDel: block.paraMarkDel ? JSON.stringify(block.paraMarkDel) : null,
          blockRevision: block.blockRevision ?? null,
          ...formatAttrs(block.format, block.runs),
        },
        content: runsToInline(block.runs ?? []),
      }
    case 'paragraph':
      return {
        type: 'docParagraph',
        attrs: {
          docxIndex: block.docxIndex,
          styleId: block.styleId ?? null,
          aiChanged: false,
          bookmarks: block.bookmarks ?? null,
          hiddenBookmarks: block.hiddenBookmarks ?? null,
          commentStarts: block.commentStarts ?? null,
          commentEnds: block.commentEnds ?? null,
          sdtShell: block.sdtShell ? JSON.stringify(block.sdtShell) : null,
          moveRevision: block.moveRevision ?? null,
          pPrChange: block.pPrChangeInfo ? JSON.stringify(block.pPrChangeInfo) : null,
          paraMarkDel: block.paraMarkDel ? JSON.stringify(block.paraMarkDel) : null,
          blockRevision: block.blockRevision ?? null,
          ...formatAttrs(block.format, block.runs),
        },
        content: runsToInline(block.runs ?? []),
      }
    case 'table': {
      const table = block.table ?? { rows: [[{ paras: [''] }]] }
      const node = tableModelToPmNode(
        table,
        block.docxIndex,
        block.blockRevision ?? null,
        rowCapTwips,
        budget ? budgetAvail(table, budget) : null,
        budget?.fit ?? null,
        budget?.paper ?? null,
        options.legacyTableIndent ?? false,
        options.defaultRun,
      )
      // content-control member tables need the shell for chrome hit-testing
      if (block.sdtShell) node.attrs = { ...node.attrs, sdtShell: JSON.stringify(block.sdtShell) }
      return node
    }
    default:
      // image / passthrough: protected whole-unit blocks
      return {
        type: 'docProtected',
        attrs: {
          docxIndex: block.docxIndex,
          blockRevision: block.blockRevision ?? null,
          blockType: block.type,
          styleId: block.styleId ?? null,
          label: block.label ?? block.type,
          previewText: block.previewText ?? '',
          imageDataUrl: block.imageDataUrl ?? null,
          oleProgId: block.oleProgId ?? null,
          imageWidthPx: block.imageWidthPx ?? null,
          imageHeightPx: block.imageHeightPx ?? null,
          imageCrop: block.imageCrop ?? null,
          imageEffects: block.imageEffects ?? null,
          imageFillRect: block.imageFillRect ?? null,
          imageLeadingText: block.imageLeadingText ?? null,
          imageLeadingFont: block.imageLeadingFont ?? null,
          imageLeadingExplicitSpaceWidthPx: block.imageLeadingExplicitSpaceWidthPx ?? null,
          imageLeadingImplicitSpaceCount: block.imageLeadingImplicitSpaceCount ?? null,
          imageParagraphIndentLeft: block.imageParagraphIndentLeft ?? null,
          imageParagraphIndentRight: block.imageParagraphIndentRight ?? null,
          imageParagraphIndentFirstLine: block.imageParagraphIndentFirstLine ?? null,
          imageParagraphSpaceBefore: block.imageParagraphSpaceBefore ?? null,
          imageParagraphSpaceAfter: block.imageParagraphSpaceAfter ?? null,
          imageEffectExtentTopPx: block.imageEffectExtentTopPx ?? null,
          imageEffectExtentBottomPx: block.imageEffectExtentBottomPx ?? null,
          imageAlign: block.imageAlign ?? null,
          imageWrap: block.imageWrap ?? null,
          imageBand: block.imageBand ?? false,
          imageWrapDistTopEmu: block.imageWrapDistTopEmu ?? null,
          imageWrapDistBottomEmu: block.imageWrapDistBottomEmu ?? null,
          imageWrapDistLeftEmu: block.imageWrapDistLeftEmu ?? null,
          imageWrapDistRightEmu: block.imageWrapDistRightEmu ?? null,
          imageZOrder: block.imageZOrder ?? null,
          imageOffsetXEmu: block.imageOffsetXEmu ?? null,
          imageOffsetYEmu: block.imageOffsetYEmu ?? null,
          imageRelV: block.imageRelV ?? null,
          imageAnchorLocked: block.imageAnchorLocked ?? false,
          imagePosH: block.imagePosH ?? null,
          imagePosV: block.imagePosV ?? null,
          imageRotDeg: block.imageRotDeg ?? null,
          imageFlipH: block.imageFlipH ?? false,
          imageFlipV: block.imageFlipV ?? false,
          imageBorder: block.imageBorder ?? null,
          table: displayTable(block.table ?? null, rowCapTwips, budget, options),
          fieldDisplay: block.fieldDisplay ?? null,
          diagramDisplay: block.diagramDisplay ?? null,
          decorative: block.decorative ?? false,
          ruleColorHex: block.ruleColorHex ?? null,
          ruleThicknessPx: block.ruleThicknessPx ?? null,
          ruleWidthPx: block.ruleWidthPx ?? null,
          brokenImage: block.brokenImage ?? false,
          invisibleMarker: block.invisibleMarker ?? false,
          anchorLine: block.anchorLine
            ? {
                styleId: block.anchorLine.styleId ?? null,
                ...formatAttrs(block.anchorLine.format),
                ...(block.anchorLine.clear ? { clear: block.anchorLine.clear } : {}),
              }
            : null,
          textboxes: block.textboxes ?? null,
          anchorSnapToGrid: block.anchorSnapToGrid ?? null,
          strayRuns: block.strayRuns ?? null,
          strayStyleId: block.strayStyleId ?? null,
          strayIndent: block.strayIndent ?? null,
          strayAlign: block.strayAlign ?? null,
          strayList: block.strayList ?? null,
          formulaDisplay: block.formulaDisplay ?? null,
          chartDisplay: block.chartDisplay ?? null,
        },
      }
  }
}

function cellRowSpan(model: TableModel, row: number, cell: number, positions: number[][]): number {
  const current = model.rows[row][cell]
  if (current.vMerge !== 'restart') return 1
  const gridColumn = positions[row][cell]
  let span = 1
  for (let r = row + 1; r < model.rows.length; r++) {
    const index = positions[r].indexOf(gridColumn)
    if (index === -1 || model.rows[r][index].vMerge !== 'continue') break
    span++
  }
  return span
}

const cellBorderPx = borderDrawnPx

/** Inner clip-box height (twips) for cells of hRule="exact" rows — Word clips overflow
 *  instead of growing the row. null = not exact, or the cell spans rows (its clip height
 *  would be the sum of the spanned rows; not handled yet). */
/** collapsed left/right line widths (drawn px) a cell's text inset must absorb, the
 *  neighbours' lines included; null when both sides fall back to the table-level default */
export function collapsedCellBw(
  row: TableModel['rows'][number],
  index: number,
): { l?: number; r?: number } | null {
  const cell = row[index]
  const l = collapsedEdgePx(cell.borders?.left, row[index - 1]?.borders?.right)
  const r = collapsedEdgePx(cell.borders?.right, row[index + 1]?.borders?.left)
  if (l === undefined && r === undefined) return null
  return { ...(l !== undefined ? { l } : {}), ...(r !== undefined ? { r } : {}) }
}

export function cellClipTwips(
  model: TableModel,
  rowIndex: number,
  cell: TableCell,
  rowSpan: number,
): number | null {
  if (rowSpan > 1 || model.rowHeightRules?.[rowIndex] !== 'exact') return null
  const h = model.rowHeightsTwips?.[rowIndex]
  if (!h) return null
  const padTop = cell.cellMarTwips?.top ?? model.cellMarTwips?.top ?? 0
  const padBottom = cell.cellMarTwips?.bottom ?? model.cellMarTwips?.bottom ?? 0
  // collapsed borders straddle the row edges: half of each eats into the row height
  const bTop = cell.borders?.top ?? (rowIndex === 0 ? model.borders?.top : model.borders?.insideH)
  const bBottom =
    cell.borders?.bottom ??
    (rowIndex === model.rows.length - 1 ? model.borders?.bottom : model.borders?.insideH)
  const borderTwips = Math.round(((cellBorderPx(bTop) + cellBorderPx(bBottom)) / 2) * 15)
  return Math.max(0, h - padTop - padBottom - borderTwips)
}

/** free side beside a text-anchored float below which Word's side text is only word fragments */
const FLOAT_SIDE_MIN_TWIPS = 1440

export function tableModelToPmNode(
  model: TableModel,
  docxIndex: number | null = null,
  blockRevision: Block['blockRevision'] | null = null,
  rowCapTwips: number | null = null,
  availTwips: number | null = null,
  fitTwips: number | null = null,
  paperTwips: number | null = null,
  legacyIndent = false,
  defaults?: RunDefaults,
): PmNode {
  if (legacyIndent) model = legacyIndentTable(model)
  if (rowCapTwips != null) model = capTableRowHeights(model, rowCapTwips)
  if (availTwips != null) {
    model = expandAutofitColWidths(
      model,
      availTwips,
      fitTwips ?? availTwips,
      undefined,
      false,
      legacyIndent,
      defaults,
    )
    model = clampTableColWidths(model, paperTwips ?? availTwips)
  }
  const positions = model.rows.map((row) => {
    let column = 0
    return row.map((cell) => {
      const at = column
      column += cell.colSpan ?? 1
      return at
    })
  })
  // 1 px = 15 twips (96dpi); use real values when absolute grid widths exist, fall back to percentage approximation otherwise.
  // No px floor on real grids: form grids hold dozens of sliver columns that only sit inside spanned cells
  const widthPx = model.colWidthsTwips
    ? model.colWidthsTwips.map((width) => Math.max(1, Math.round(width / 15)))
    : model.colWidthsPct?.map((width) => Math.max(24, Math.round(width * 6.24)))
  const hasHeaderRow =
    model.rows.length > 1 && model.rows[0].every((cell) => cell.bold || cell.fill !== undefined)
  // w:tblpPr tables taller than any single page cannot float: Word flows/splits
  // them across pages, while the zero-flow-height float model collapses the
  // whole document to one page (137-row glossary, LO batch2 sample 0219).
  // Minimum height = one 12pt line per row; 12960 twips ≈ the shortest common
  // usable page (Letter, 1in margins).
  const minHeightTwips = model.rows.reduce(
    (sum, _row, i) => sum + Math.max(model.rowHeightsTwips?.[i] ?? 0, 240),
    0,
  )
  const tblFloatSource = model.floatSide ?? null
  // Word splits text-anchored floating tables across page boundaries instead of
  // pushing them whole; a float that leaves less than ~1in beside it hosts a
  // few word fragments at most (Word probe 2026-09-04), so flowing it inline
  // reproduces the split and stops the push from opening a near-blank page
  // (prod100r4/68). It also keeps the next block's line boxes, which CSS drops
  // below a float they cannot sit beside, from turning that block into one
  // float-spanning over-page line that clips away a page of content (SAS
  // batch 2 sample 028). Only the subset that STARTS at the column's left edge
  // (within ~1in of the paper edge for page anchors, at the margin otherwise):
  // inline flow reproduces its geometry, while a mid-page X keeps the
  // clamped-float rendering (#1111 cap-table corpus).
  const floatX = model.floatPos?.xTwips ?? 0
  const floatSpan =
    (model.floatPos?.horzAnchor === 'page' ? 0 : floatX) +
    (model.colWidthsTwips?.reduce((a, b) => a + b, 0) ?? 0) +
    (model.floatPos?.distanceTwips?.right ?? 0)
  const floatNoSideRoom =
    (model.floatPos?.vertAnchor ?? 'text') === 'text' &&
    (model.floatPos?.horzAnchor === 'page' ? floatX <= 1440 : floatX <= 720) &&
    fitTwips != null &&
    floatSpan > fitTwips - FLOAT_SIDE_MIN_TWIPS
  const tblFloatSuppressed = tblFloatSource !== null && (minHeightTwips > 12960 || floatNoSideRoom)
  const tblFloat = tblFloatSuppressed ? null : tblFloatSource
  const table: PmNode = {
    type: 'docTable',
    attrs: {
      docxIndex,
      blockRevision,
      colWidthsPct: model.colWidthsPct ?? null,
      // never narrower than the grid: Word's autofit columns hold their widest
      // line with zero slack, so rounding the table down wraps that line
      widthPx:
        !model.widthPct && model.colWidthsTwips && widthPx
          ? Math.max(
              widthPx.length,
              Math.ceil(model.colWidthsTwips.reduce((sum, width) => sum + width, 0) / 15 - 1e-6),
            )
          : null,
      widthPct: model.widthPct ?? null,
      cellMar: model.cellMarTwips ?? null,
      cellSpacingTwips: model.cellSpacingTwips ?? null,
      tblFill: model.fill ?? null,
      cellMarEdited: false,
      borders: model.borders ?? null,
      tblAlign: model.align ?? null,
      tblFloat,
      tblFloatSource,
      tblFloatSuppressed,
      tblFloatXTwips: model.floatPos?.xTwips ?? null,
      tblFloatYTwips: model.floatPos?.yTwips ?? null,
      tblFloatHorzAnchor: model.floatPos?.horzAnchor ?? null,
      tblFloatVertAnchor: model.floatPos?.vertAnchor ?? null,
      tblFloatXSpec: model.floatPos?.xSpec ?? null,
      tblFloatYSpec: model.floatPos?.ySpec ?? null,
      tblFloatDistance: model.floatPos?.distanceTwips ?? null,
      tblFloatWidthPx:
        tblFloatSource && widthPx ? widthPx.reduce((sum, width) => sum + width, 0) : null,
      tblFloatEdited: false,
      tblAutoFit: model.autoFit ?? (model.autoLayout ? 'contents' : 'fixed'),
      tblAutoFitEdited: false,
      tblFixedLayout: model.fixedLayout ?? false,
      // w:tblpPr supersedes w:tblInd: a float keeps only the legacy cell-margin hang
      indentTwips:
        tblFloat && !(legacyIndent && marginHungFloat(model)) ? null : (model.indentTwips ?? null),
      tblStyleId: model.tblStyleId ?? null,
      tblLook: model.tableLook ?? null,
      tblLookEdited: false,
      bidiVisual: model.bidiVisual ?? false,
      tblCaption: model.caption ?? null,
      tblDescription: model.description ?? null,
      tblAltEdited: false,
      originalStructure: null,
      originalFormatting: null,
    },
    content: model.rows.map((row, rowIndex) => ({
      type: 'docTableRow',
      attrs: {
        heightTwips: model.rowHeightsTwips?.[rowIndex] ?? null,
        heightRule: model.rowHeightRules?.[rowIndex] ?? null,
        repeatHeader: model.repeatHeaderRows?.[rowIndex] ?? false,
        repeatHeaderEdited: false,
        cantSplit: model.rowCantSplit?.[rowIndex] ?? false,
        cantSplitEdited: false,
        rawTrPr: model.rawTrPrs?.[rowIndex] ?? null,
        rowRevision: model.rowRevisions?.[rowIndex] ?? null,
      },
      content: row.flatMap((cell, cellIndex) => {
        if (cell.vMerge === 'continue') return []
        const start = positions[rowIndex][cellIndex]
        const colspan = cell.colSpan ?? 1
        const rowspan = cellRowSpan(model, rowIndex, cellIndex, positions)
        return [
          {
            type: hasHeaderRow && rowIndex === 0 ? 'docTableHeader' : 'docTableCell',
            attrs: {
              colspan,
              rowspan,
              clipHeightTwips: cellClipTwips(model, rowIndex, cell, rowspan),
              colwidth: widthPx ? widthPx.slice(start, start + colspan) : null,
              gridGap: cell.gridGap ?? false,
              cellMar: cell.cellMarTwips ?? null,
              textDirection: cell.textDirection ?? null,
              noWrap: cell.noWrap ?? null,
              tcPrEdited: false,
              fill: cell.fill ?? null,
              color: cell.styleColor ?? null,
              bold: cell.styleBold ?? false,
              align: cell.align ?? null,
              vAlign: cell.vAlign ?? null,
              borders: cell.borders ?? null,
              collapsedBw: collapsedCellBw(row, cellIndex),
              rawTcPr: cell.rawTcPr ?? null,
              cellRevision: cell.cellRevision ?? null,
            },
            content: cellContentNodes(cell),
          },
        ]
      }),
    })),
  }
  table.attrs!.originalStructure = tableStructureSignature(table)
  table.attrs!.originalFormatting = tableFormattingSignature(table)
  return table
}

/** cell paragraphs with nested tables (read-only atoms) spliced in at their anchors */
function cellContentNodes(cell: TableCell): PmNode[] {
  const paraNodes: PmNode[] = (
    cell.richParas?.length
      ? cell.richParas
      : (cell.paras.length > 0 ? cell.paras : ['']).map((text) => ({
          align: cell.align,
          runs: text === '' ? [] : [{ text, bold: cell.bold, color: cell.color }],
        }))
  ).map((paragraph) => {
    const list = 'list' in paragraph ? paragraph.list : undefined
    const styleId = ('styleId' in paragraph ? paragraph.styleId : undefined) ?? null
    return list
      ? {
          type: 'docListItem',
          attrs: {
            ...formatAttrs(paragraph, paragraph.runs),
            styleId,
            kind: list.kind,
            numId: list.numId,
            ilvl: list.ilvl,
          },
          content: runsToInline(paragraph.runs),
        }
      : {
          type: 'docParagraph',
          attrs: { ...formatAttrs(paragraph, paragraph.runs), styleId },
          content: runsToInline(paragraph.runs),
        }
  })
  // fidelity on save comes from the outer table's bytes; reverse insertion keeps anchors valid
  const inserts: Array<{ at: number; node: PmNode }> = []
  const nested = cell.nestedTables ?? []
  for (let i = 0; i < nested.length; i++) {
    inserts.push({
      at: Math.min(cell.nestedTableAnchors?.[i] ?? paraNodes.length, paraNodes.length),
      node: { type: 'docNestedTable', attrs: { model: nested[i] } },
    })
  }
  // anchored shapes/textboxes (display-only): a zero-width float strut before
  // each group's anchor paragraph, so the boxes' positionV offsets resolve from
  // that paragraph like Word and the row grows to max(text, boxes)
  const boxes = cell.anchoredBoxes ?? []
  const boxGroups = new Map<number, TextboxDisplay[]>()
  for (let i = 0; i < boxes.length; i++) {
    const at = Math.min(cell.anchoredBoxAnchors?.[i] ?? 0, paraNodes.length)
    const group = boxGroups.get(at)
    if (group) group.push(boxes[i])
    else boxGroups.set(at, [boxes[i]])
  }
  for (const [at, group] of boxGroups) {
    inserts.push({ at, node: { type: 'docCellBoxes', attrs: { boxes: group } } })
  }
  const content = [...paraNodes]
  inserts.sort((a, b) => a.at - b.at)
  for (let i = inserts.length - 1; i >= 0; i--) {
    content.splice(inserts[i].at, 0, inserts[i].node)
  }
  return content
}

export function tableStructureSignature(table: PmNode): string {
  return JSON.stringify({
    widths: table.attrs?.colWidthsPct ?? null,
    widthPx: table.attrs?.widthPx ?? null,
    widthPct: table.attrs?.widthPct ?? null,
    autoFit: table.attrs?.tblAutoFit ?? null,
    autoFitEdited: table.attrs?.tblAutoFitEdited ?? false,
    cellMar: table.attrs?.cellMar ?? null,
    cellMarEdited: table.attrs?.cellMarEdited ?? false,
    tblStyle: table.attrs?.tblStyleId ?? null,
    tblAlign: table.attrs?.tblAlign ?? null,
    tblFloat: table.attrs?.tblFloat ?? null,
    tblFloatXTwips: table.attrs?.tblFloatXTwips ?? null,
    tblFloatYTwips: table.attrs?.tblFloatYTwips ?? null,
    tblFloatHorzAnchor: table.attrs?.tblFloatHorzAnchor ?? null,
    tblFloatVertAnchor: table.attrs?.tblFloatVertAnchor ?? null,
    tblFloatDistance: table.attrs?.tblFloatDistance ?? null,
    tblFloatEdited: table.attrs?.tblFloatEdited ?? false,
    tblLook: table.attrs?.tblLook ?? null,
    tblLookEdited: table.attrs?.tblLookEdited ?? false,
    tblCaption: table.attrs?.tblCaption ?? null,
    tblDescription: table.attrs?.tblDescription ?? null,
    tblAltEdited: table.attrs?.tblAltEdited ?? false,
    rows: (table.content ?? []).map((row) => [
      row.attrs?.heightTwips ?? null,
      row.attrs?.heightRule ?? null,
      row.attrs?.repeatHeader ?? false,
      row.attrs?.repeatHeaderEdited ?? false,
      row.attrs?.cantSplit ?? false,
      row.attrs?.cantSplitEdited ?? false,
      // accepting/rejecting revisions strips records from trPr/tcPr: include them in the signature to trigger regeneration (works for empty rows too)
      row.attrs?.rawTrPr ?? null,
      row.attrs?.rowRevision ?? null,
      ...(row.content ?? []).map((cell) => [
        cell.type,
        cell.attrs?.colspan ?? 1,
        cell.attrs?.rowspan ?? 1,
        cell.attrs?.colwidth ?? null,
        cell.attrs?.fill ?? null,
        cell.attrs?.color ?? null,
        cell.attrs?.bold ?? false,
        cell.attrs?.align ?? null,
        cell.attrs?.vAlign ?? null,
        cell.attrs?.borders ?? null,
        cell.attrs?.rawTcPr ?? null,
        cell.attrs?.cellRevision ?? null,
        cell.attrs?.textDirection ?? null,
        cell.attrs?.cellMar ?? null,
        cell.attrs?.noWrap ?? null,
        cell.attrs?.tcPrEdited ?? false,
      ]),
    ]),
  })
}

function inlineFormattingSignature(content: PmNode[] | undefined): string[] {
  const styles: string[] = []
  for (const node of content ?? []) {
    // the schema re-sorts marks by rank on load, so the parse-side order must not count
    const style =
      node.type === 'text'
        ? JSON.stringify(
            [...(node.marks ?? [])]
              .sort((a, b) => a.type.localeCompare(b.type))
              .map((mark) => [mark.type, mark.attrs ?? null]),
          )
        : node.type
    if (styles[styles.length - 1] !== style) styles.push(style)
  }
  return styles
}

export function tableFormattingSignature(table: PmNode): string {
  return JSON.stringify(
    (table.content ?? []).map((row) =>
      (row.content ?? []).map((cell) =>
        (cell.content ?? []).map((paragraph) => [
          normalizedFormat(nodeFormat(paragraph)),
          // list toggles must force regeneration (the text patch keeps old pPr)
          paragraph.type === 'docListItem'
            ? [paragraph.attrs?.kind, paragraph.attrs?.numId, paragraph.attrs?.ilvl]
            : null,
          inlineFormattingSignature(paragraph.content),
        ]),
      ),
    ),
  )
}

function paragraphText(node: PmNode): string {
  return inlineToRuns(node.content ?? [])
    .map((run) => run.text)
    .join('')
}

function unwrapBlockRevisionXml(xml: string): string {
  const match = /^<w:(ins|del)\b[^>]*>([\s\S]*)<\/w:\1>$/.exec(xml)
  return match ? match[2] : xml
}

/** Convert native table cells back to the physical OOXML-style TableModel grid. */
export function pmTableToModel(table: PmNode): TableModel {
  type ActiveSpan = {
    remaining: number
    span: number
    cell: Omit<TableModel['rows'][number][number], 'paras'>
  }
  let active = new Map<number, ActiveSpan>()
  const rows: TableModel['rows'] = []
  let maximumColumns = 0

  const rowHeightsTwips: Array<number | null> = []
  const rowHeightRules: NonNullable<TableModel['rowHeightRules']> = []
  const repeatHeaderRows: Array<boolean | null> = []
  const rowCantSplit: Array<boolean | null> = []
  const rawTrPrs: Array<string | null> = []
  const rowRevisions: TableModel['rowRevisions'] = []
  for (const rowNode of table.content ?? []) {
    rowHeightsTwips.push((rowNode.attrs?.heightTwips as number | null) ?? null)
    rowHeightRules.push(
      (rowNode.attrs?.heightRule as NonNullable<TableModel['rowHeightRules']>[number]) ?? null,
    )
    repeatHeaderRows.push(
      rowNode.attrs?.repeatHeaderEdited || table.attrs?.docxIndex == null
        ? !!rowNode.attrs?.repeatHeader
        : null,
    )
    rowCantSplit.push(
      rowNode.attrs?.cantSplitEdited || table.attrs?.docxIndex == null
        ? !!rowNode.attrs?.cantSplit
        : null,
    )
    rawTrPrs.push((rowNode.attrs?.rawTrPr as string | null) ?? null)
    rowRevisions.push(
      (rowNode.attrs?.rowRevision as NonNullable<TableModel['rowRevisions']>[number]) ?? null,
    )
    const entries: Array<{ column: number; cell: TableModel['rows'][number][number] }> = []
    const occupied = new Set<number>()
    for (const [column, span] of active) {
      entries.push({
        column,
        cell: {
          ...span.cell,
          paras: [''],
          colSpan: span.span > 1 ? span.span : undefined,
          vMerge: 'continue',
        },
      })
      for (let i = 0; i < span.span; i++) occupied.add(column + i)
    }

    const added = new Map<number, ActiveSpan>()
    let cursor = 0
    for (const cellNode of rowNode.content ?? []) {
      while (occupied.has(cursor)) cursor++
      const colspan = Math.max(1, Number(cellNode.attrs?.colspan) || 1)
      const rowspan = Math.max(1, Number(cellNode.attrs?.rowspan) || 1)
      const style = {
        fill: (cellNode.attrs?.fill as string | null) ?? undefined,
        color: (cellNode.attrs?.color as string | null) ?? undefined,
        bold: cellNode.attrs?.bold ? true : undefined,
        align: (cellNode.attrs?.align as TableModel['rows'][number][number]['align']) ?? undefined,
        vAlign: (cellNode.attrs?.vAlign as TableCell['vAlign'] | null) ?? undefined,
        borders: (cellNode.attrs?.borders as TableCell['borders'] | null) ?? undefined,
        rawTcPr: (cellNode.attrs?.rawTcPr as string | null) ?? undefined,
        gridGap: cellNode.attrs?.gridGap ? true : undefined,
        // edited tcPr groups are written from the attrs; a null direction / margin
        // set means "explicitly horizontal" / "same as the table" (element removed)
        ...(cellNode.attrs?.tcPrEdited || table.attrs?.docxIndex == null
          ? {
              tcPrEdited: true,
              textDirection:
                (cellNode.attrs?.textDirection as TableCell['textDirection']) ?? 'lrTb',
              cellMarTwips: (cellNode.attrs?.cellMar as TableCell['cellMarTwips']) ?? {},
              noWrap: !!cellNode.attrs?.noWrap,
            }
          : {}),
      }
      const cellParas = (cellNode.content ?? []).filter(
        (n) => n.type === 'docParagraph' || n.type === 'docListItem',
      )
      // nested tables round-trip with the model: text edits use the surgical patch, structural regeneration emits the whole table (nested tables no longer dropped)
      const nestedModels: TableModel[] = []
      const nestedAnchors: number[] = []
      let paraCount = 0
      const cellBoxes: TextboxDisplay[] = []
      const cellBoxAnchors: number[] = []
      for (const n of cellNode.content ?? []) {
        if (n.type === 'docParagraph' || n.type === 'docListItem') paraCount++
        else if (n.type === 'docNestedTable' && n.attrs?.model) {
          nestedModels.push(n.attrs.model as TableModel)
          nestedAnchors.push(paraCount)
        } else if (n.type === 'docCellBoxes' && Array.isArray(n.attrs?.boxes)) {
          for (const box of n.attrs.boxes as TextboxDisplay[]) {
            cellBoxes.push(box)
            cellBoxAnchors.push(paraCount)
          }
        }
      }
      entries.push({
        column: cursor,
        cell: {
          ...style,
          paras: cellParas.map(paragraphText),
          richParas: cellParas.map((paragraph) => ({
            ...nodeFormat(paragraph),
            ...(paragraph.type === 'docListItem' && paragraph.attrs?.numId
              ? {
                  list: {
                    kind: (paragraph.attrs.kind as 'bullet' | 'ordered') ?? 'bullet',
                    numId: String(paragraph.attrs.numId),
                    ilvl: Number(paragraph.attrs.ilvl) || 0,
                  },
                }
              : {}),
            runs: inlineToRuns(paragraph.content ?? []),
          })),
          ...(nestedModels.length > 0
            ? { nestedTables: nestedModels, nestedTableAnchors: nestedAnchors }
            : {}),
          ...(cellBoxes.length > 0
            ? { anchoredBoxes: cellBoxes, anchoredBoxAnchors: cellBoxAnchors }
            : {}),
          colSpan: colspan > 1 ? colspan : undefined,
          vMerge: rowspan > 1 ? 'restart' : undefined,
        },
      })
      if (rowspan > 1) added.set(cursor, { remaining: rowspan - 1, span: colspan, cell: style })
      for (let i = 0; i < colspan; i++) occupied.add(cursor + i)
      cursor += colspan
    }
    maximumColumns = Math.max(maximumColumns, occupied.size)
    rows.push(entries.sort((a, b) => a.column - b.column).map((entry) => entry.cell))

    const next = new Map<number, ActiveSpan>()
    for (const [column, span] of active) {
      if (span.remaining > 1) next.set(column, { ...span, remaining: span.remaining - 1 })
    }
    for (const [column, span] of added) next.set(column, span)
    active = next
  }

  let colWidthsPct = (table.attrs?.colWidthsPct as number[] | null) ?? undefined
  let colWidthsTwips: number[] | undefined
  const firstRow = table.content?.[0]?.content ?? []
  const explicit = firstRow.flatMap((cell) => (cell.attrs?.colwidth as number[] | null) ?? [])
  if (explicit.length === maximumColumns && explicit.every((width) => width > 0)) {
    const total = explicit.reduce((sum, width) => sum + width, 0)
    colWidthsPct = explicit.map((width) => (width / total) * 100)
    colWidthsTwips = explicit.map((width) => Math.round(width * 15))
  } else if (colWidthsPct?.length !== maximumColumns) {
    colWidthsPct = Array.from({ length: maximumColumns }, () => 100 / maximumColumns)
  }
  const tblStyleAttr = table.attrs?.tblStyleId as string | null | undefined
  const isNew = table.attrs?.docxIndex == null
  const tablePropsEdited = (name: string): boolean => isNew || table.attrs?.[name] === true
  const visibleFloat =
    table.attrs?.tblFloat === 'left' || table.attrs?.tblFloat === 'right'
      ? (table.attrs.tblFloat as 'left' | 'right')
      : null
  const suppressedFloat =
    table.attrs?.tblFloatSuppressed &&
    (table.attrs?.tblFloatSource === 'left' || table.attrs?.tblFloatSource === 'right')
      ? (table.attrs.tblFloatSource as 'left' | 'right')
      : null
  const floatSide = visibleFloat ?? suppressedFloat
  const floatPosition =
    floatSide && (table.attrs?.tblFloatXTwips != null || table.attrs?.tblFloatYTwips != null)
      ? {
          xTwips: Number(table.attrs?.tblFloatXTwips) || 0,
          yTwips: Number(table.attrs?.tblFloatYTwips) || 0,
          ...(table.attrs?.tblFloatHorzAnchor
            ? {
                horzAnchor: table.attrs.tblFloatHorzAnchor as NonNullable<
                  TableModel['floatPos']
                >['horzAnchor'],
              }
            : {}),
          ...(table.attrs?.tblFloatVertAnchor
            ? {
                vertAnchor: table.attrs.tblFloatVertAnchor as NonNullable<
                  TableModel['floatPos']
                >['vertAnchor'],
              }
            : {}),
          ...(table.attrs?.tblFloatDistance
            ? {
                distanceTwips: table.attrs.tblFloatDistance as NonNullable<
                  TableModel['floatPos']
                >['distanceTwips'],
              }
            : {}),
        }
      : undefined
  return {
    rows,
    ...(colWidthsPct ? { colWidthsPct } : {}),
    ...(colWidthsTwips ? { colWidthsTwips } : {}),
    ...(table.attrs?.cellSpacingTwips
      ? { cellSpacingTwips: Number(table.attrs.cellSpacingTwips) }
      : {}),
    ...(table.attrs?.tblFill ? { fill: String(table.attrs.tblFill) } : {}),
    ...(rowHeightsTwips.some((h) => h !== null) ? { rowHeightsTwips, rowHeightRules } : {}),
    ...(repeatHeaderRows.some((value) => value !== null) ? { repeatHeaderRows } : {}),
    ...(rowCantSplit.some((value) => value !== null) ? { rowCantSplit } : {}),
    ...(rawTrPrs.some((r) => r !== null) ? { rawTrPrs } : {}),
    ...(rowRevisions.some((revision) => revision !== null) ? { rowRevisions } : {}),
    // null (cleared) → '' removes explicitly; undefined leaves it alone
    ...(tblStyleAttr !== undefined ? { tblStyleId: tblStyleAttr ?? '' } : {}),
    // explicit only when set ('left' = remove w:jc); null leaves the original
    // tblPr untouched so unmapped w:jc values (e.g. 'start') survive rebuilds
    ...(table.attrs?.tblAlign ? { align: table.attrs.tblAlign as TableModel['align'] } : {}),
    ...(tablePropsEdited('tblAutoFitEdited')
      ? { autoFit: table.attrs?.tblAutoFit as NonNullable<TableModel['autoFit']> }
      : {}),
    ...(tablePropsEdited('cellMarEdited') && table.attrs?.cellMar
      ? { cellMarTwips: table.attrs.cellMar as NonNullable<TableModel['cellMarTwips']> }
      : {}),
    ...(tablePropsEdited('tblFloatEdited')
      ? { floatSide, ...(floatPosition ? { floatPos: floatPosition } : {}) }
      : {}),
    ...(tablePropsEdited('tblLookEdited') && table.attrs?.tblLook
      ? { tableLook: table.attrs.tblLook as NonNullable<TableModel['tableLook']> }
      : {}),
    // new tables only carry alt text when set; edited ones write '' to drop the element
    ...(table.attrs?.tblAltEdited ||
    (isNew && (table.attrs?.tblCaption || table.attrs?.tblDescription))
      ? {
          caption: String(table.attrs?.tblCaption ?? ''),
          description: String(table.attrs?.tblDescription ?? ''),
        }
      : {}),
  }
}

export function runsToInline(runs: Run[]): PmNode[] {
  const nodes: PmNode[] = []
  for (const run of runs) {
    if (run.math) {
      nodes.push({
        type: 'docInlineMath',
        attrs: {
          omml: run.math.omml,
          mathml: inlineMathML(run.math.omml),
          // recovered LaTeX makes parsed formulas re-editable; null = atom only
          latex: ommlToLatex(run.math.omml),
          text: run.text,
        },
      })
      continue
    }
    if (run.ruby) {
      nodes.push({
        type: 'docRuby',
        attrs: { base: run.text, rt: run.ruby.rt, xml: run.ruby.xml },
      })
      continue
    }
    if (run.noteRef) {
      nodes.push({
        type: 'docNoteRef',
        attrs: {
          kind: run.noteRef.kind,
          id: run.noteRef.id,
          num: parseInt(run.text, 10) || 1,
        },
      })
      continue
    }
    const marks = runMarks(run)
    // \n = soft line break, \f = in-paragraph page break, \v = column break, \u001e = text wrapping break
    const breakMarks = marks.length > 0 ? { marks } : {}
    // eslint-disable-next-line no-control-regex
    for (const segment of symDisplayText(run).split(/([\n\f\v\u001e])/)) {
      if (segment === '\n') nodes.push({ type: 'hardBreak', ...breakMarks })
      else if (segment === '\f')
        nodes.push({ type: 'hardBreak', attrs: { pageBreak: true }, ...breakMarks })
      else if (segment === '\v')
        nodes.push({ type: 'hardBreak', attrs: { colBreak: true }, ...breakMarks })
      else if (segment === '\u001e')
        nodes.push({ type: 'hardBreak', attrs: { wrapBreak: true }, ...breakMarks })
      else if (segment !== '') {
        nodes.push({ type: 'text', text: segment, ...(marks.length > 0 ? { marks } : {}) })
      }
    }
    // A run can carry both w:t text and a w:drawing; generate.ts writes the
    // text before the drawing, so emit the image after the text segments
    if (run.image) {
      nodes.push({
        type: 'docInlineImage',
        attrs: {
          dataUrl: run.image.dataUrl,
          widthPx: run.image.widthPx ?? null,
          heightPx: run.image.heightPx ?? null,
          xml: run.image.xml,
          wrap: run.image.wrap ?? null,
          offsetXEmu: run.image.offsetXEmu ?? null,
          offsetYEmu: run.image.offsetYEmu ?? null,
          relV: run.image.relV ?? null,
          wrapDistTopEmu: run.image.wrapDistTopEmu ?? null,
          wrapDistBottomEmu: run.image.wrapDistBottomEmu ?? null,
          wrapDistLeftEmu: run.image.wrapDistLeftEmu ?? null,
          wrapDistRightEmu: run.image.wrapDistRightEmu ?? null,
          border: run.image.border ?? null,
          lineCenterV: run.image.lineCenterV ?? false,
          rotDeg: run.image.rotDeg ?? null,
          flipH: run.image.flipH ?? false,
          flipV: run.image.flipV ?? false,
          rule: run.image.rule
            ? {
                ...run.image.rule,
                ...(run.sizeHalfPoints ? { sizeHalfPoints: run.sizeHalfPoints } : {}),
              }
            : null,
          rawRPr: run.rawRPr ?? null,
        },
      })
    }
    if (run.xeTerm !== undefined) {
      nodes.push({ type: 'docXeMark', attrs: { term: run.xeTerm } })
    }
  }
  markLeadRule(nodes)
  return nodes
}

const isRuleNode = (n: PmNode): boolean => n.type === 'docInlineImage' && !!n.attrs?.rule

/** a rule opening a paragraph that goes on with content sits on a compact line in
 *  Word (rule height + ~5.5pt), unlike a rule alone or after text (a full line) */
function markLeadRule(nodes: PmNode[]): void {
  const first = nodes[0]
  if (!first || !isRuleNode(first)) return
  if (nodes.slice(1).some((n) => !isRuleNode(n))) first.attrs = { ...first.attrs, leadRule: true }
}

let nextZoteroFieldId = 1

function runMarks(run: Run): PmMark[] {
  const marks: PmMark[] = []
  if (run.bold) marks.push({ type: 'bold' })
  if (run.italic) marks.push({ type: 'italic' })
  if (run.underline) marks.push({ type: 'underline' })
  if (run.strike) marks.push({ type: 'strike' })
  if (run.link)
    marks.push({
      type: 'link',
      attrs: {
        href: run.link.href,
        rId: run.link.rId ?? null,
        tooltip: run.link.tooltip ?? null,
        plain: run.link.plain === true,
      },
    })
  if (run.refField !== undefined)
    marks.push({
      type: 'refField',
      attrs: { name: run.refField, instr: run.refInstr ?? null, dirty: run.fldDirty === true },
    })
  if (run.sym) marks.push({ type: 'docSym', attrs: { font: run.sym.font, char: run.sym.char } })
  if (run.instrField !== undefined) {
    const isZotero = /^\s*(?:ADDIN\s+)?(?:ZOTERO_|CSL_)/i.test(run.instrField)
    let fieldId: number | null = null
    if (isZotero) {
      if (Number.isSafeInteger(run.zoteroFieldId) && Number(run.zoteroFieldId) > 0) {
        fieldId = Number(run.zoteroFieldId)
        nextZoteroFieldId = Math.max(nextZoteroFieldId, fieldId + 1)
      } else {
        fieldId = nextZoteroFieldId++
      }
    }
    marks.push({
      type: 'instrField',
      attrs: {
        instr: run.instrField,
        beginXml: run.fldBeginXml ?? null,
        dirty: run.fldDirty === true,
        fieldId,
        fieldPart: run.zoteroFieldPart ?? null,
      },
    })
  }
  if (run.sdtCheckboxXml !== undefined)
    marks.push({ type: 'ctrlCheckbox', attrs: { sdtPr: run.sdtCheckboxXml } })
  if (run.commentIds?.length)
    marks.push({ type: 'comment', attrs: { ids: run.commentIds.join(' ') } })
  if (run.ins) {
    marks.push({
      type: 'ins',
      attrs: { author: run.ins.author, date: run.ins.date ?? null, id: run.ins.id ?? null },
    })
  }
  if (run.del) {
    marks.push({
      type: 'del',
      attrs: { author: run.del.author, date: run.del.date ?? null, id: run.del.id ?? null },
    })
  }
  if (run.rPrChange) {
    marks.push({
      type: 'rprChange',
      attrs: {
        author: run.rPrChange.author,
        date: run.rPrChange.date ?? null,
        id: run.rPrChange.id ?? null,
        old: run.rPrChange.old ?? null,
      },
    })
  }
  if (
    run.color ||
    run.sizeHalfPoints ||
    run.font ||
    run.fontAscii ||
    run.charSpacingTwips !== undefined ||
    run.charScalePct ||
    run.kernHalfPoints !== undefined ||
    run.highlight ||
    run.shading ||
    run.shadingDisplay ||
    run.textOutline ||
    run.textEffect ||
    run.dstrike ||
    run.glow ||
    run.positionHalfPoints ||
    run.bdr ||
    run.vertAlign ||
    run.em ||
    run.bold === false ||
    run.italic === false ||
    run.caps ||
    run.vanish ||
    run.eastAsiaLang ||
    run.cs ||
    run.rtl !== undefined ||
    run.styleId ||
    run.rawRPr
  ) {
    // an inline-block cannot wrap and drops inherited underlines: only
    // whitespace-free, undecorated runs get the real glyph compression
    const scaleX =
      run.charScalePct && !/\s/.test(run.text) && !run.underline && !run.strike && !run.link
        ? charScaleXAttr(run.text, run.charScalePct, charScaleEm(run.text, run.charScalePct))
        : null
    marks.push({
      type: 'docTextStyle',
      attrs: {
        color: run.color ?? null,
        sizeHalfPoints: run.sizeHalfPoints ?? null,
        font: run.font ?? null,
        eaSlotEmpty: run.eaSlotEmpty ?? null,
        fontAscii: run.fontAscii ?? null,
        eastAsiaFont: run.eastAsiaFont ?? null,
        // cs chain only kicks in for complex-script text (Word's w:cs semantics)
        csFont: run.csFont && textHasComplexScript(run.text) ? run.csFont : null,
        charSpacingTwips: run.charSpacingTwips ?? null,
        charScaleEm: run.charScalePct && !scaleX ? charScaleEm(run.text, run.charScalePct) : null,
        charScaleX: scaleX,
        charScalePct: run.charScalePct ?? null,
        kern: wordKerns(run.kernHalfPoints, run.sizeHalfPoints) ?? null,
        kernHalfPoints: run.kernHalfPoints ?? null,
        highlight: run.highlight ?? null,
        shading: run.shading ?? null,
        shadingDisplay: run.shadingDisplay ?? null,
        textOutline: run.textOutline ? JSON.stringify(run.textOutline) : null,
        textEffect: run.textEffect ?? null,
        dstrike: run.dstrike ?? null,
        glow: run.glow ? JSON.stringify(run.glow) : null,
        positionHalfPoints: run.positionHalfPoints ?? null,
        bdr: run.bdr ? JSON.stringify(run.bdr) : null,
        vertAlign: run.vertAlign ?? null,
        em: run.em ?? null,
        boldOff: run.bold === false || null,
        italicOff: run.italic === false || null,
        caps: run.caps ?? null,
        vanish: run.vanish ?? null,
        vanishOwn: run.vanishOwn ?? null,
        eaLang: run.eastAsiaLang ?? null,
        cs: run.cs ?? null,
        rtl: run.rtl ?? null,
        styleId: run.styleId ?? null,
        rawRPr: run.rawRPr ?? null,
        themeRFonts: run.themeRFonts ? JSON.stringify(run.themeRFonts) : null,
        themeColor: run.themeColor ?? null,
      },
    })
  }
  return marks
}

// ---- ProseMirror doc -> SaveBlock[] (dirty detection via content signatures) ----

export interface SavePlan {
  saveBlocks: SaveBlock[]
  /**
   * Chart data edits. These patch the chart's own zip part, not the body
   * paragraph (which stays byte-identical); App applies patchChartPartXml
   * and hands the result to saveDocx via options.partXml.
   */
  chartPatches: Array<{ partPath: string; patch: ChartPatch }>
  /**
   * docxIndex of an original block -> its index in saveBlocks. Ink
   * annotations anchor by docxIndex and the engine wants finalBlocks indexes.
   */
  saveBlockIndexByDocx: Map<number, number>
  /** count of blocks whose content was regenerated */
  changedCount: number
  /** count of original blocks removed */
  deletedCount: number
  totalOriginal: number
}

const ZOTERO_FIELD_INSTR_RE = /^\s*(?:ADDIN\s+)?(?:ZOTERO_|CSL_)/i
type ZoteroFieldPart = 'single' | 'begin' | 'inside' | 'end'
interface ZoteroTextRef {
  blockIndex: number
  /** ordinal of the leaf among the block's leaves, for adjacency checks */
  leafIndex: number
  node: PmNode
  markIndex: number
  instr: string
  fieldId: number | null
  part: string | null
}

/**
 * Zotero fields whose result spans paragraphs share a fieldId, and each run's
 * part decides where the field begin/end runs are written. Editing can leave a
 * group without those runs (paragraph deleted, begin run split by Enter, last
 * run's text removed, pasted copy without ids), so the parts are recomputed
 * from document order before saving: the first run begins, the last one ends,
 * and a group interrupted by another field or a non-paragraph block becomes
 * separate fields instead of swallowing what lies between.
 */
export function normalizeZoteroFieldParts(doc: PmNode): PmNode {
  const blocks = doc.content ?? []
  const refs: ZoteroTextRef[] = []
  const leafCounts: number[] = []
  const collect = (node: PmNode, blockIndex: number) => {
    if (node.content) {
      for (const child of node.content) collect(child, blockIndex)
      return
    }
    const leafIndex = leafCounts[blockIndex]++
    if (node.text === undefined) return
    const markIndex = (node.marks ?? []).findIndex(
      (m) => m.type === 'instrField' && ZOTERO_FIELD_INSTR_RE.test(String(m.attrs?.instr ?? '')),
    )
    if (markIndex < 0) return
    const attrs = node.marks![markIndex].attrs ?? {}
    const id = Number(attrs.fieldId)
    refs.push({
      blockIndex,
      leafIndex,
      node,
      markIndex,
      instr: String(attrs.instr ?? ''),
      fieldId: Number.isSafeInteger(id) && id > 0 ? id : null,
      part: typeof attrs.fieldPart === 'string' ? attrs.fieldPart : null,
    })
  }
  blocks.forEach((block, index) => {
    leafCounts[index] = 0
    collect(block, index)
  })
  if (refs.length === 0) return doc

  let nextId = Math.max(0, ...refs.map((ref) => ref.fieldId ?? 0)) + 1
  // runs that lost their id (clipboard round trip) rejoin as one field while they
  // touch each other with the same instruction (or end one paragraph and open the
  // next); two copies of a citation with text between them stay two fields, and
  // they never join a field that kept its id
  const touches = (prev: ZoteroTextRef, next: ZoteroTextRef) =>
    prev.blockIndex === next.blockIndex
      ? next.leafIndex === prev.leafIndex + 1
      : next.blockIndex === prev.blockIndex + 1 &&
        prev.leafIndex === leafCounts[prev.blockIndex] - 1 &&
        next.leafIndex === 0
  const lostId = new Set<ZoteroTextRef>()
  for (let i = 0; i < refs.length; i++) {
    const ref = refs[i]
    if (ref.fieldId !== null) continue
    const prev = refs[i - 1]
    const joins =
      prev !== undefined && lostId.has(prev) && prev.instr === ref.instr && touches(prev, ref)
    ref.fieldId = joins ? prev.fieldId : nextId++
    ref.part = null
    lostId.add(ref)
  }
  const groups = new Map<number, ZoteroTextRef[]>()
  const idsInBlock = new Map<number, Set<number>>()
  for (const ref of refs) {
    const id = ref.fieldId!
    groups.set(id, [...(groups.get(id) ?? []), ref])
    idsInBlock.set(ref.blockIndex, new Set([...(idsInBlock.get(ref.blockIndex) ?? []), id]))
  }
  const isParagraphBlock = (index: number) => {
    const type = blocks[index]?.type
    return type === 'docParagraph' || type === 'docHeading' || type === 'docListItem'
  }
  const desired = new Map<PmNode, { markIndex: number; fieldId: number; part: ZoteroFieldPart }>()
  for (const [id, members] of groups) {
    const segments: ZoteroTextRef[][] = [[members[0]]]
    for (let i = 1; i < members.length; i++) {
      const prev = members[i - 1]
      const cur = members[i]
      let split = false
      for (let b = prev.blockIndex + 1; b < cur.blockIndex && !split; b++) {
        const others = idsInBlock.get(b)
        split = !isParagraphBlock(b) || (others !== undefined && [...others].some((o) => o !== id))
      }
      if (split) segments.push([cur])
      else segments[segments.length - 1].push(cur)
    }
    segments.forEach((segment, si) => {
      const fieldId = si === 0 ? id : nextId++
      segment.forEach((ref, ri) => {
        const part: ZoteroFieldPart =
          segment.length === 1
            ? 'single'
            : ri === 0
              ? 'begin'
              : ri === segment.length - 1
                ? 'end'
                : 'inside'
        const idAttr = Number(ref.node.marks![ref.markIndex].attrs?.fieldId)
        if (idAttr !== fieldId || ref.part !== part)
          desired.set(ref.node, { markIndex: ref.markIndex, fieldId, part })
      })
    })
  }
  if (desired.size === 0) return doc

  const rebuild = (node: PmNode): PmNode => {
    const want = desired.get(node)
    if (want) {
      const marks = node.marks!.map((mark, i) =>
        i === want.markIndex
          ? { ...mark, attrs: { ...mark.attrs, fieldId: want.fieldId, fieldPart: want.part } }
          : mark,
      )
      return { ...node, marks }
    }
    if (!node.content) return node
    let changed = false
    const content = node.content.map((child) => {
      const next = rebuild(child)
      if (next !== child) changed = true
      return next
    })
    return changed ? { ...node, content } : node
  }
  return rebuild(doc)
}

export function pmDocToSavePlan(inputDoc: PmNode, originalBlocks: Block[]): SavePlan {
  const doc = normalizeZoteroFieldParts(inputDoc)
  const originalByIndex = new Map<number, Block>()
  for (const block of originalBlocks) {
    if (!block.hidden && block.docxIndex !== null) originalByIndex.set(block.docxIndex, block)
  }

  // Mixed-encoding guard: parse compresses wild producer relativeHeight values
  // (LibreOffice writes 1, 2, …) to compact ranks, but the raw XML keeps the
  // wild bytes. Word paints by the raw value, so rewriting ONE anchor to the
  // base+rank encoding while wild siblings survive would invert the saved
  // paint order. Once any wrap/z-order edit exists, every normalized anchor is
  // rewritten to base+rank; untouched documents still keep their bytes.
  const hasZOrderEdit = (nodes: PmNode[] | undefined): boolean => {
    for (const n of nodes ?? []) {
      if (n.type === 'docProtected' && n.attrs?.blockType === 'image') {
        const idx = n.attrs.docxIndex as number | null | undefined
        const orig = idx !== null && idx !== undefined ? originalByIndex.get(idx) : undefined
        if (orig) {
          const wrap = (n.attrs.imageWrap as ImageWrap | null) ?? null
          const z = n.attrs.imageZOrder != null ? Number(n.attrs.imageZOrder) : undefined
          if (wrap !== (orig.imageWrap ?? null) || (z !== undefined && z !== orig.imageZOrder))
            return true
        }
      }
      if (hasZOrderEdit(n.content)) return true
    }
    return false
  }
  const harmonizeZOrders =
    originalBlocks.some((b) => b.imageZOrderNormalized) && hasZOrderEdit(doc.content)

  // Multi-block sdt groups (one w:sdt split into N blocks): the shell open/close
  // bytes live on the first/last member. Track which surviving member emits
  // them so deleting or regenerating a boundary member never unbalances the sdt.
  const sdtGroupOf = new Map<number, number>()
  const sdtGroupShells = new Map<number, SdtShell>()
  // first/last original member per group: nested content controls leave wrapper
  // fragments on MIDDLE members' closeXml, so "emits closeXml" alone no longer
  // means "emits the group close" — only the boundary members do
  const sdtGroupFirst = new Map<number, number>()
  const sdtGroupLast = new Map<number, number>()
  for (const block of originalBlocks) {
    const shell = block.sdtShell
    if (block.hidden || block.docxIndex === null || shell?.group === undefined) continue
    sdtGroupOf.set(block.docxIndex, shell.group)
    if (!sdtGroupFirst.has(shell.group)) sdtGroupFirst.set(shell.group, block.docxIndex)
    sdtGroupLast.set(shell.group, block.docxIndex)
    const agg = sdtGroupShells.get(shell.group)
    if (!agg) sdtGroupShells.set(shell.group, { ...shell })
    else {
      if (shell.openXml) agg.openXml = shell.openXml
      if (shell.closeXml) agg.closeXml = shell.closeXml
    }
  }
  const sdtGroupEmits = new Map<
    number,
    Array<{ at: number; idx: number; open: boolean; close: boolean }>
  >()

  const saveBlocks: SaveBlock[] = []
  const chartPatches: Array<{ partPath: string; patch: ChartPatch }> = []
  const saveBlockIndexByDocx = new Map<number, number>()
  let changedCount = 0
  const usedIndexes = new Set<number>()
  /** record where an original anchor ended up in saveBlocks (set before push) */
  const mapAnchor = (idx: number) => saveBlockIndexByDocx.set(idx, saveBlocks.length)

  const recordSdtEmit = (node: PmNode, sb: SaveBlock, at: number) => {
    const idx = node.attrs?.docxIndex as number | null | undefined
    if (idx === null || idx === undefined) return
    if (saveBlockIndexByDocx.get(idx) !== at) return
    const group = sdtGroupOf.get(idx)
    if (group === undefined) return
    const shell = sdtGroupShells.get(group)!
    let open = false
    let close = false
    if (sb.kind === 'original') {
      const original = originalByIndex.get(idx)
      open = !!original?.sdtShell?.openXml
      close = !!original?.sdtShell?.closeXml
    } else if (sb.kind === 'generated') {
      open = !!sb.block.sdtShell?.openXml
      close = !!sb.block.sdtShell?.closeXml
    } else if (sb.kind === 'xml') {
      open = !!shell.openXml && sb.xml.startsWith(shell.openXml)
      close = !!shell.closeXml && sb.xml.endsWith(shell.closeXml)
    }
    // middle members may emit nested wrapper fragments; only the boundary
    // members' bytes are the group open/close the rebalance below cares about
    open = open && sdtGroupFirst.get(group) === idx
    close = close && sdtGroupLast.get(group) === idx
    let emits = sdtGroupEmits.get(group)
    if (!emits) sdtGroupEmits.set(group, (emits = []))
    emits.push({ at, idx, open, close })
  }

  for (const node of doc.content ?? []) {
    const revision = node.attrs?.blockRevision as {
      kind: 'ins' | 'del'
      author: string
      date?: string
      id?: string
    } | null
    const pushBlock = (block: SaveBlock) => {
      recordSdtEmit(node, block, saveBlocks.length)
      saveBlocks.push(revision ? { ...block, revision } : block)
    }
    if (node.type === 'docTable') {
      const idx = node.attrs?.docxIndex as number | null
      const original =
        idx !== null && idx !== undefined && !usedIndexes.has(idx)
          ? originalByIndex.get(idx)
          : undefined
      const model = pmTableToModel(node)
      if (original?.type === 'table') {
        usedIndexes.add(idx!)
        mapAnchor(idx!)
        const structurallyUnchanged =
          node.attrs?.originalStructure === tableStructureSignature(node)
        const formattingUnchanged =
          node.attrs?.originalFormatting === tableFormattingSignature(node)
        if (
          original.blockRevision &&
          !revision &&
          structurallyUnchanged &&
          formattingUnchanged &&
          original.originalXml
        ) {
          changedCount++
          pushBlock({ kind: 'xml', xml: unwrapBlockRevisionXml(original.originalXml) })
          continue
        }
        const tableTexts =
          structurallyUnchanged && formattingUnchanged
            ? tableTextsPatchFromModel(model, original)
            : null
        if (structurallyUnchanged && formattingUnchanged && tableTexts && original.originalXml) {
          changedCount++
          pushBlock({ kind: 'xml', xml: patchTableCellTexts(original.originalXml, tableTexts) })
        } else if (structurallyUnchanged && formattingUnchanged) {
          pushBlock({ kind: 'original', docxIndex: idx! })
        } else {
          changedCount++
          pushBlock({
            kind: 'xml',
            xml: generateTableModelXml(model, original.originalXml ?? undefined),
          })
        }
      } else {
        changedCount++
        pushBlock({ kind: 'xml', xml: generateTableModelXml(model) })
      }
      continue
    }
    if (node.type === 'docProtected') {
      const idx = node.attrs?.docxIndex as number | null
      if (idx !== null && idx !== undefined && originalByIndex.has(idx) && !usedIndexes.has(idx)) {
        usedIndexes.add(idx)
        mapAnchor(idx)
        const original = originalByIndex.get(idx)!
        if (original.blockRevision && !revision && original.originalXml) {
          changedCount++
          pushBlock({ kind: 'xml', xml: unwrapBlockRevisionXml(original.originalXml) })
          continue
        }
        const imagePatch = imagePatchOf(node, original)
        const tableTexts = tableTextsPatch(node, original)
        const textboxTexts = textboxParasPatch(node, original)
        const textboxSizes = textboxSizesPatch(node, original)
        const textboxStyles = textboxStylesPatch(node, original)
        const textboxOffsetX =
          node.attrs?.imageOffsetXEmu != null ? Number(node.attrs.imageOffsetXEmu) : undefined
        const textboxOffsetY =
          node.attrs?.imageOffsetYEmu != null ? Number(node.attrs.imageOffsetYEmu) : undefined
        const textboxPositionChanged =
          !!original.textboxes?.length &&
          textboxOffsetX !== undefined &&
          textboxOffsetY !== undefined &&
          (textboxOffsetX !== original.imageOffsetXEmu ||
            textboxOffsetY !== original.imageOffsetYEmu)
        const fieldText = fieldTextPatch(node, original)
        const formulaTokens = formulaTokensPatch(node, original)
        // chart data edits live in the chart's own part; the body block stays original
        const chartPatch = chartPatchOf(node, original)
        if (chartPatch) {
          changedCount++
          chartPatches.push(chartPatch)
        }
        // chart resize rewrites the body drawing's extent
        const chartDisplay = node.attrs?.chartDisplay as ChartDisplay | null
        const chartSize =
          chartDisplay?.widthPx &&
          chartDisplay.heightPx &&
          (chartDisplay.widthPx !== original.chartDisplay?.widthPx ||
            chartDisplay.heightPx !== original.chartDisplay?.heightPx)
            ? { w: chartDisplay.widthPx, h: chartDisplay.heightPx }
            : null
        const imageReplace =
          original.type === 'image'
            ? (node.attrs?.imageReplace as { base64: string; mime: string } | null)
            : null
        if ((imagePatch || imageReplace) && original.originalXml) {
          changedCount++
          let xml = imagePatch
            ? patchImageParagraphXml(original.originalXml, imagePatch)
            : original.originalXml
          if (imagePatch?.wrap !== undefined) {
            const posOffset =
              imagePatch.posOffsetX !== undefined && imagePatch.posOffsetY !== undefined
                ? { x: imagePatch.posOffsetX, y: imagePatch.posOffsetY }
                : undefined
            const marginAlign =
              posOffset === undefined && imagePatch.posH && imagePatch.posV
                ? { h: imagePatch.posH, v: imagePatch.posV }
                : undefined
            // a wrap-only change must not reset an existing stacking rank: the
            // patch carries zOrder only when the rank itself changed
            xml = applyImageWrap(
              xml,
              imagePatch.wrap,
              posOffset,
              marginAlign,
              imagePatch.zOrder ?? original.imageZOrder,
            )
          } else if (
            imagePatch?.zOrder !== undefined &&
            (node.attrs?.imageWrap as ImageWrap | null) != null
          ) {
            // z-order changed without a wrap-mode change on a still-floating
            // image: re-encode ONLY relativeHeight — a full anchor rebuild
            // would reset the position basis (relativeFrom) and distances
            xml = applyImageZOrder(xml, imagePatch.zOrder)
          } else if (
            imagePatch &&
            (imagePatch.posOffsetX !== undefined || imagePatch.posOffsetY !== undefined)
          ) {
            // posOffset changed without wrap change: patchImageParagraphXml already rewrote it
          } else if (
            harmonizeZOrders &&
            original.type === 'image' &&
            original.imageZOrderNormalized &&
            (node.attrs?.imageWrap as ImageWrap | null) != null
          ) {
            // geometry-only edit (resize/align/rotate) on a sibling that still
            // carries a wild producer relativeHeight: re-encode base+rank on
            // top of the geometry patch — relativeHeight only, the anchor's
            // position basis and wrap bytes must survive (see applyImageZOrder)
            xml = applyImageZOrder(xml, original.imageZOrder)
          }
          pushBlock({
            kind: 'xml',
            xml,
            ...(imageReplace
              ? {
                  replaceImage: {
                    base64: imageReplace.base64,
                    mime: imageReplace.mime as NewImage['mime'],
                  },
                }
              : {}),
          })
        } else if (tableTexts && original.originalXml) {
          changedCount++
          pushBlock({ kind: 'xml', xml: patchTableCellTexts(original.originalXml, tableTexts) })
        } else if (
          (textboxTexts || textboxSizes || textboxStyles || textboxPositionChanged) &&
          original.originalXml
        ) {
          changedCount++
          let xml = original.originalXml
          if (textboxTexts) xml = patchTextboxParas(xml, textboxTexts)
          if (textboxSizes) xml = patchTextboxSizes(xml, textboxSizes)
          if (textboxStyles) xml = patchShapeStyles(xml, textboxStyles)
          if (textboxPositionChanged) {
            const wrap =
              (node.attrs?.imageWrap as ImageWrap | null) ?? original.imageWrap ?? 'square-left'
            xml = applyImageWrap(xml, wrap, { x: textboxOffsetX!, y: textboxOffsetY! })
          }
          pushBlock({ kind: 'xml', xml })
        } else if (fieldText && original.originalXml) {
          changedCount++
          pushBlock({ kind: 'xml', xml: patchFieldParagraphXml(original.originalXml, fieldText) })
        } else if (formulaTokens && original.originalXml) {
          changedCount++
          pushBlock({ kind: 'xml', xml: patchMathTokens(original.originalXml, formulaTokens) })
        } else if (chartSize && original.originalXml) {
          changedCount++
          pushBlock({
            kind: 'xml',
            xml: patchDrawingExtent(original.originalXml, chartSize.w, chartSize.h),
          })
        } else if (
          harmonizeZOrders &&
          original.type === 'image' &&
          original.imageZOrderNormalized &&
          original.imageWrap &&
          original.originalXml
        ) {
          // untouched anchor still carrying a wild producer relativeHeight:
          // re-encode ONLY that attribute to base+rank so it stays ordered
          // against the anchors this session re-encoded (harmonizeZOrders
          // pre-scan). Position/wrap bytes stay untouched — the picture must
          // not shift from a reorder it wasn't even part of.
          changedCount++
          pushBlock({
            kind: 'xml',
            xml: applyImageZOrder(original.originalXml, original.imageZOrder),
          })
        } else {
          pushBlock({ kind: 'original', docxIndex: idx })
        }
      } else if (idx !== null && idx !== undefined && originalByIndex.has(idx)) {
        // pasted copy: the anchor is consumed by the first occurrence, so clone from the block's own data
        const original = originalByIndex.get(idx)!
        const image = imageFromProtectedAttrs(node)
        if (image) {
          changedCount++
          pushBlock({ kind: 'image', image })
        } else if (original.originalXml) {
          changedCount++
          pushBlock({ kind: 'xml', xml: stripAnchorMarkers(original.originalXml) })
        } else {
          console.warn('dropping unclonable copy of protected block', node.attrs?.label)
        }
      } else if (node.attrs?.genXml) {
        // editor-created table or other self-contained fragment
        changedCount++
        let xml = String(node.attrs.genXml)
        const table = node.attrs?.table as TableModel | null
        if (table && node.attrs?.blockType === 'table') {
          xml = patchTableCellTexts(
            xml,
            table.rows.map((row) => row.map((cell) => cell.paras)),
          )
        }
        // editor-generated TOC lines: write the auto-refreshed page number back
        // (right only — the title is already in the genXml, and a left+right
        // patch bails out entirely when the title text nodes don't line up)
        const genField = node.attrs?.fieldDisplay as FieldDisplay | null
        if (genField?.kind === 'tocLine' && genField.right) {
          xml = patchFieldParagraphXml(xml, { right: genField.right })
        }
        // patch textbox paragraphs for newly-inserted textboxes/shapes with text
        const genTextboxes = node.attrs?.textboxes as TextboxDisplay[] | null
        const hasNonEmptyTextbox =
          genTextboxes != null &&
          genTextboxes.length > 0 &&
          genTextboxes.some((box) => box.paras.some((p) => p.runs.some((r) => r.text !== '')))
        if (hasNonEmptyTextbox) {
          // editor-built XML has exactly one txbxContent per box, in box order
          xml = patchTextboxParas(xml, {
            byIndex: genTextboxes!.map((box) =>
              box.paras.map((p) => ({ runs: mergeRuns(p.runs), align: p.align ?? null })),
            ),
          })
        }
        // persist resize/autogrow of newly-inserted single-box shapes/lines;
        // horizontal lines keep their zero-height extent (display box is a grab band)
        const genBox = genTextboxes?.length === 1 ? genTextboxes[0] : null
        if (genBox && (genBox.widthPx || genBox.heightPx)) {
          xml = patchTextboxSizes(xml, [
            {
              wPx: genBox.widthPx ?? null,
              hPx: isStraightLineKind(genBox.prst) ? null : (genBox.heightPx ?? null),
            },
          ])
        }
        if (genBox) {
          xml = patchShapeStyles(xml, [
            { fillHex: genBox.fill ?? null, borderHex: genBox.borderColor ?? null },
          ])
        }
        // apply wrap changes for floating textboxes/shapes
        const genWrap = node.attrs?.imageWrap as ImageWrap | null
        const genOffsetX =
          node.attrs?.imageOffsetXEmu != null ? Number(node.attrs.imageOffsetXEmu) : undefined
        const genOffsetY =
          node.attrs?.imageOffsetYEmu != null ? Number(node.attrs.imageOffsetYEmu) : undefined
        if (genWrap !== undefined && genWrap !== null) {
          const posOffset =
            genOffsetX !== undefined && genOffsetY !== undefined
              ? { x: genOffsetX, y: genOffsetY }
              : undefined
          xml = applyImageWrap(xml, genWrap, posOffset)
        }
        pushBlock({ kind: 'xml', xml })
      } else if (node.attrs?.genImage) {
        changedCount++
        const image = { ...(node.attrs.genImage as NewImage) }
        if (node.attrs.imageWidthPx) image.widthPx = Number(node.attrs.imageWidthPx)
        if (node.attrs.imageHeightPx) image.heightPx = Number(node.attrs.imageHeightPx)
        const align = node.attrs.imageAlign as NewImage['align'] | null
        if (align) image.align = align
        const wrap = node.attrs.imageWrap as ImageWrap | null
        if (wrap) image.wrap = wrap
        if (node.attrs.imageZOrder != null) image.zOrder = Number(node.attrs.imageZOrder)
        if (node.attrs.imageRotDeg) image.rotDeg = Number(node.attrs.imageRotDeg)
        if (node.attrs.imageFlipH) image.flipH = true
        if (node.attrs.imageFlipV) image.flipV = true
        pushBlock({ kind: 'image', image })
      } else if (node.attrs?.genChart) {
        // in-place edits live in chartDisplay; the saved part reflects them
        changedCount++
        const spec = { ...(node.attrs.genChart as NewChart) }
        const display = node.attrs.chartDisplay as ChartDisplay | null
        if (display) {
          if (display.title !== undefined) spec.title = display.title
          spec.categories = display.categories
          spec.series = display.series.map((s, i) => ({
            name: s.name ?? spec.series[i]?.name ?? t('editorChartSeries', { num: i + 1 }),
            values: s.values,
          }))
        }
        pushBlock({
          kind: 'chart',
          chart: spec,
          ...(display?.widthPx && display.heightPx
            ? { extentPx: { w: display.widthPx, h: display.heightPx } }
            : {}),
        })
      } else {
        // picture that traveled through clipboard HTML (r136 cross-document
        // paste): no anchor and no genImage — rebuild from the preview bytes
        // or it silently vanishes on save
        const image = imageFromProtectedAttrs(node)
        if (image) {
          changedCount++
          pushBlock({ kind: 'image', image })
        }
      }
      // any other protected node without an anchor or generated payload cannot be regenerated; drop it
      continue
    }

    const generated = pmNodeToGeneratedBlock(node)
    const idx = node.attrs?.docxIndex as number | null
    const original = idx !== null && idx !== undefined ? originalByIndex.get(idx) : undefined
    if (
      original &&
      !usedIndexes.has(idx!) &&
      signatureOfBlock(original) === signatureOfGenerated(generated)
    ) {
      usedIndexes.add(idx!)
      mapAnchor(idx!)
      pushBlock({ kind: 'original', docxIndex: idx! })
    } else {
      changedCount++
      // an in-place edit consumes its original anchor: it is replaced, not
      // deleted — and may reuse the anchor's original pPr bytes. A split twin
      // (anchor already consumed) gets no rawPPr, so revision records and
      // section props are never duplicated.
      const replacesOriginal = original !== undefined && !usedIndexes.has(idx!)
      if (replacesOriginal) {
        usedIndexes.add(idx!)
        mapAnchor(idx!)
        applyRawPPr(generated, original!)
      } else if (original) {
        // A split twin inherits the source node's attrs, but pPrChange and
        // bookmark/comment anchors belong only to the anchored original.
        delete generated.pPrChange
        delete generated.bookmarks
        delete generated.hiddenBookmarks
        delete generated.commentStarts
        delete generated.commentEnds
        // A grouped shell (partial open/close of a multi-block sdt) must be
        // emitted exactly once, by the anchor.
        if (generated.sdtShell?.group !== undefined) delete generated.sdtShell
      }
      pushBlock({
        kind: 'generated',
        block: generated,
        ...(replacesOriginal ? { docxIndex: idx! } : {}),
      })
    }
  }

  const attachShell = (emit: { at: number; idx: number }, side: 'open' | 'close') => {
    const group = sdtGroupOf.get(emit.idx)!
    const shell = sdtGroupShells.get(group)!
    const add = side === 'open' ? shell.openXml : shell.closeXml
    const sb = saveBlocks[emit.at]
    if (sb.kind === 'original') {
      const xml = originalByIndex.get(emit.idx)?.originalXml
      if (xml === undefined) return
      saveBlocks[emit.at] = {
        kind: 'xml',
        xml: side === 'open' ? add + xml : xml + add,
        docxIndex: emit.idx,
        ...(sb.revision ? { revision: sb.revision } : {}),
      }
    } else if (sb.kind === 'generated') {
      const gShell = sb.block.sdtShell ?? { ...shell, openXml: '', closeXml: '' }
      // concatenate: a nested wrapper fragment may already sit here, and the
      // group open must precede it (the group close must follow it)
      if (side === 'open') gShell.openXml = add + (gShell.openXml ?? '')
      else gShell.closeXml = (gShell.closeXml ?? '') + add
      sb.block.sdtShell = gShell
    } else if (sb.kind === 'xml') {
      sb.xml = side === 'open' ? add + sb.xml : sb.xml + add
    }
  }
  // Nested content controls put inner wrapper fragments on MIDDLE members'
  // closeXml (between-children bytes: "</inner sdt><next inner sdt>"). A
  // deleted middle member must not take its fragment with it, or the group
  // saves unbalanced sdt XML. Reattach the fragment to the nearest surviving
  // member before it (else the first after), before the group-level open/close
  // rebalance below so the outer close still lands last.
  const appendXml = (at: number, frag: string, before: boolean, group: number) => {
    const sb = saveBlocks[at]
    if (sb.kind === 'original') {
      const xml = originalByIndex.get(sb.docxIndex)?.originalXml
      if (xml === undefined) return
      saveBlocks[at] = {
        kind: 'xml',
        xml: before ? frag + xml : xml + frag,
        docxIndex: sb.docxIndex,
        ...(sb.revision ? { revision: sb.revision } : {}),
      }
    } else if (sb.kind === 'generated') {
      const shell = sb.block.sdtShell ?? {
        ...sdtGroupShells.get(group)!,
        openXml: '',
        closeXml: '',
      }
      if (before) shell.openXml = frag + (shell.openXml ?? '')
      else shell.closeXml = (shell.closeXml ?? '') + frag
      sb.block.sdtShell = shell
    } else if (sb.kind === 'xml') {
      sb.xml = before ? frag + sb.xml : sb.xml + frag
    }
  }
  for (const [group, emits] of sdtGroupEmits) {
    const members = [...sdtGroupOf.entries()]
      .filter(([, g]) => g === group)
      .map(([idx]) => idx)
      .sort((a, b) => a - b)
    const last = members[members.length - 1]
    const sorted = [...emits].sort((a, b) => a.at - b.at)
    // deleted leading members all prepend to the first survivor: collect their
    // fragments in document order and prepend once, or they stack reversed
    const prefixFrags: string[] = []
    for (const idx of members) {
      if (idx === last || usedIndexes.has(idx)) continue
      const frag = originalByIndex.get(idx)?.sdtShell?.closeXml
      if (!frag) continue
      const prev = [...sorted].reverse().find((e) => e.idx < idx)
      if (prev) appendXml(prev.at, frag, false, group)
      else prefixFrags.push(frag)
    }
    if (prefixFrags.length > 0) appendXml(sorted[0].at, prefixFrags.join(''), true, group)
  }
  for (const [group, emits] of sdtGroupEmits) {
    const shell = sdtGroupShells.get(group)!
    if (shell.openXml && !emits.some((e) => e.open)) attachShell(emits[0], 'open')
    if (shell.closeXml && !emits.some((e) => e.close)) attachShell(emits[emits.length - 1], 'close')
  }

  const deletedCount = [...originalByIndex.keys()].filter((i) => !usedIndexes.has(i)).length
  return {
    saveBlocks,
    chartPatches,
    saveBlockIndexByDocx,
    changedCount,
    deletedCount,
    totalOriginal: originalByIndex.size,
  }
}

/** rebuild a pasted image copy from its preview bytes; null when not materializable */
function imageFromProtectedAttrs(node: PmNode): NewImage | null {
  if (node.attrs?.blockType !== 'image') return null
  const src = String(node.attrs?.imageDataUrl ?? '')
  const lazy = parseLazyMediaUrl(src)
  const m = lazy ? null : /^data:(image\/(?:png|jpeg|gif));base64,(.+)$/.exec(src)
  const widthPx = Number(node.attrs?.imageWidthPx)
  const heightPx = Number(node.attrs?.imageHeightPx)
  if ((!m && !lazy) || !widthPx || !heightPx) return null
  const image: NewImage = lazy
    ? { base64: '', mime: 'image/png', sourcePart: lazy.partPath, widthPx, heightPx }
    : { base64: m![2], mime: m![1] as NewImage['mime'], widthPx, heightPx }
  const align = node.attrs?.imageAlign as NewImage['align'] | null
  if (align) image.align = align
  const wrap = node.attrs?.imageWrap as ImageWrap | null
  if (wrap) image.wrap = wrap
  if (node.attrs?.imageZOrder != null) image.zOrder = Number(node.attrs.imageZOrder)
  if (node.attrs?.imageRotDeg) image.rotDeg = Number(node.attrs.imageRotDeg)
  if (node.attrs?.imageFlipH) image.flipH = true
  if (node.attrs?.imageFlipV) image.flipV = true
  return image
}

/** cloned XML must not repeat the anchor's bookmark/comment ids */
function stripAnchorMarkers(xml: string): string {
  return xml.replace(
    /<w:(?:bookmarkStart|bookmarkEnd|commentRangeStart|commentRangeEnd|commentReference)\b[^>]*\/>/g,
    '',
  )
}

/** chart data changes vs the parsed model; null when untouched or unpatchable */
function chartPatchOf(
  node: PmNode,
  original: Block,
): { partPath: string; patch: ChartPatch } | null {
  const current = node.attrs?.chartDisplay as ChartDisplay | null
  const initial = original.chartDisplay
  if (!current || !initial || current.partPath !== initial.partPath) return null
  if (
    current.series.length !== initial.series.length ||
    current.categories.length !== initial.categories.length
  ) {
    return null
  }
  const patch: ChartPatch = {}
  if (initial.title !== undefined && (current.title ?? '') !== initial.title) {
    patch.title = current.title ?? ''
  }
  let catChanged = false
  const categories = current.categories.map((cat, i) => {
    if (cat === initial.categories[i]) return null
    catChanged = true
    return cat
  })
  if (catChanged) patch.categories = categories
  let serChanged = false
  const series = current.series.map((ser, i): ChartSeriesPatch | null => {
    const orig = initial.series[i]
    if (ser.values.length !== orig.values.length) return null
    const serPatch: ChartSeriesPatch = {}
    if (orig.name !== undefined && (ser.name ?? '') !== orig.name) serPatch.name = ser.name ?? ''
    let valChanged = false
    const values = ser.values.map((value, j) => {
      // gaps in the cache have no pt to patch; they stay read-only
      if (value === orig.values[j] || orig.values[j] === null) return null
      valChanged = true
      return value
    })
    if (valChanged) serPatch.values = values
    if (Object.keys(serPatch).length === 0) return null
    serChanged = true
    return serPatch
  })
  if (serChanged) patch.series = series
  if (Object.keys(patch).length === 0) return null
  return { partPath: initial.partPath, patch }
}

interface ImageBlockPatch {
  widthPx?: number
  heightPx?: number
  align?: 'left' | 'center' | 'right' | null
  /** null = back to inline; undefined = wrap unchanged */
  wrap?: ImageWrap | null
  /** stacking rank among overlapping anchors (bring-forward/send-back); undefined = keep */
  zOrder?: number
  /** posOffset in EMU for free-position floating images; undefined = keep */
  posOffsetX?: number
  posOffsetY?: number
  /** margin-relative align pair (Word position-gallery presets); undefined = keep */
  posH?: 'left' | 'center' | 'right'
  posV?: 'top' | 'center' | 'bottom'
  /** rotation (deg clockwise, 0 removes) / mirror flips; undefined = keep */
  rotDeg?: number
  flipH?: boolean
  flipV?: boolean
}

/** size/align/wrap changes on an original image block; null when untouched */
function imagePatchOf(node: PmNode, original: Block): ImageBlockPatch | null {
  if (original.type !== 'image' || node.attrs?.blockType !== 'image') return null
  const patch: ImageBlockPatch = {}
  const w = node.attrs?.imageWidthPx ? Number(node.attrs.imageWidthPx) : null
  const h = node.attrs?.imageHeightPx ? Number(node.attrs.imageHeightPx) : null
  if (w && h && (w !== (original.imageWidthPx ?? null) || h !== (original.imageHeightPx ?? null))) {
    patch.widthPx = w
    patch.heightPx = h
  }
  const align = (node.attrs?.imageAlign as 'left' | 'center' | 'right' | null) ?? null
  if (align !== (original.imageAlign ?? null)) {
    patch.align = align
  }
  const wrap = (node.attrs?.imageWrap as ImageWrap | null) ?? null
  if (wrap !== (original.imageWrap ?? null)) {
    patch.wrap = wrap
  }
  const zOrder = node.attrs?.imageZOrder != null ? Number(node.attrs.imageZOrder) : undefined
  if (zOrder !== undefined && zOrder !== (original.imageZOrder ?? undefined)) {
    patch.zOrder = zOrder
  }
  // posOffset for free-position floating images
  const posOffsetX =
    node.attrs?.imageOffsetXEmu != null ? Number(node.attrs.imageOffsetXEmu) : undefined
  const posOffsetY =
    node.attrs?.imageOffsetYEmu != null ? Number(node.attrs.imageOffsetYEmu) : undefined
  if (posOffsetX !== undefined && posOffsetX !== (original.imageOffsetXEmu ?? undefined)) {
    patch.posOffsetX = posOffsetX
  }
  if (posOffsetY !== undefined && posOffsetY !== (original.imageOffsetYEmu ?? undefined)) {
    patch.posOffsetY = posOffsetY
  }
  const rotDeg = node.attrs?.imageRotDeg != null ? Number(node.attrs.imageRotDeg) : 0
  if (rotDeg !== (original.imageRotDeg ?? 0)) patch.rotDeg = rotDeg
  const flipH = !!node.attrs?.imageFlipH
  const flipV = !!node.attrs?.imageFlipV
  if (flipH !== (original.imageFlipH ?? false)) patch.flipH = flipH
  if (flipV !== (original.imageFlipV ?? false)) patch.flipV = flipV
  const posH = (node.attrs?.imagePosH as ImageBlockPatch['posH'] | null) ?? null
  const posV = (node.attrs?.imagePosV as ImageBlockPatch['posV'] | null) ?? null
  if (
    posH &&
    posV &&
    (posH !== (original.imagePosH ?? null) || posV !== (original.imagePosV ?? null))
  ) {
    patch.posH = posH
    patch.posV = posV
    // the position pair is written by applyImageWrap, which only runs when wrap is in the patch
    if (patch.wrap === undefined) patch.wrap = wrap
  }
  return Object.keys(patch).length > 0 ? patch : null
}

/** gridGap placeholders have no w:tc in the original XML: drop them so the
 *  patch grid's indexes line up with document-order w:tc segments */
function patchableCells(row: TableCell[] | undefined): TableCell[] {
  return (row ?? []).filter((cell) => !cell.gridGap)
}

/** per-cell text changes on an original table block; null when untouched */
function tableTextsPatch(node: PmNode, original: Block): (string[] | null)[][] | null {
  if (original.type !== 'table' || node.attrs?.blockType !== 'table') return null
  const current = node.attrs?.table as TableModel | null
  const orig = original.table
  if (!current || !orig) return null
  let changed = false
  const texts = current.rows.map((row, r) => {
    const origRow = patchableCells(orig.rows[r])
    return patchableCells(row).map((cell, c) => {
      const originalCell = origRow[c]
      if (!originalCell || cell.vMerge === 'continue') return null
      if (cell.paras.join('\n') === originalCell.paras.join('\n')) return null
      changed = true
      return cell.paras
    })
  })
  return changed ? texts : null
}

function tableTextsPatchFromModel(
  current: TableModel,
  original: Block,
): (CellTextsPatch | null)[][] | null {
  if (original.type !== 'table' || !original.table) return null
  let changed = false
  const texts = current.rows.map((row, r) => {
    const origRow = patchableCells(original.table!.rows[r])
    return patchableCells(row).map((cell, c): CellTextsPatch | null => {
      const originalCell = origRow[c]
      // the PM model never carries continuation-cell content: their original bytes stay
      if (!originalCell || cell.vMerge === 'continue') return null
      // nested tables: diff cell texts per table; changes go through the recursive surgical patch
      if (cell.nestedTables?.length && originalCell.nestedTables?.length) {
        const nested = cell.nestedTables.map((nt, i) => {
          const origNt = originalCell.nestedTables![i]
          if (!origNt) return null
          const grid = nestedTextsDiff(nt, origNt)
          return grid
        })
        if (nested.some((g) => g !== null)) {
          changed = true
          return { nested }
        }
        return null
      }
      const origTexts = parsedCellParaTexts(originalCell)
      if (cell.paras.join('\n') === origTexts.join('\n')) return null
      changed = true
      const rich = cell.richParas
      // rich paragraph patches keep per-run formatting, hyperlinks and images;
      // untouched paragraphs (per-index text match) keep their original bytes
      if (rich && rich.length === cell.paras.length) {
        if (rich.length === origTexts.length) {
          return rich.map((p, i): CellParaPatch =>
            p.runs.map((run) => run.text).join('') === origTexts[i] ? null : { runs: p.runs },
          )
        }
        return rich.map((p): CellParaPatch => ({ runs: p.runs }))
      }
      return cell.paras
    })
  })
  return changed ? texts : null
}

/**
 * Original-cell paragraph texts in the same encoding the PM model uses (run
 * texts with tabs/breaks/symbols as control/decoded chars). The parse-side
 * paras cache is built with a plain text extractor that drops them, so
 * comparing against it flags every break/tab/symbol cell as edited and sends
 * the whole untouched table through the lossy cell-text patch.
 */
function parsedCellParaTexts(cell: TableCell): string[] {
  return cell.richParas?.length
    ? cell.richParas.map((p) => p.runs.map((run) => run.text).join(''))
    : cell.paras
}

/** per-cell text diff of one nested table; null = untouched cell, null entry = untouched paragraph */
function nestedTextsDiff(
  current: TableModel,
  original: TableModel,
): (CellParaPatch[] | null)[][] | null {
  if (current.rows.length !== original.rows.length) return null
  let changed = false
  const grid = current.rows.map((row, r) => {
    const origRow = patchableCells(original.rows[r])
    return patchableCells(row).map((cell, c): CellParaPatch[] | null => {
      const originalCell = origRow[c]
      if (!originalCell || cell.vMerge === 'continue') return null
      const text = cell.paras.join('\n')
      const richTexts = parsedCellParaTexts(originalCell)
      if (text === originalCell.paras.join('\n') || text === richTexts.join('\n')) return null
      changed = true
      if (cell.paras.length !== originalCell.paras.length) return cell.paras
      return cell.paras.map((p, i) =>
        p === originalCell.paras[i] || p === richTexts[i] ? null : p,
      )
    })
  })
  return changed ? grid : null
}

/** rich-run signature of one textbox paragraph, for per-paragraph change detection */
export function textboxParaSignature(para: TextboxDisplay['paras'][number]): string {
  return JSON.stringify([para.align ?? null, normalizedRuns(para.runs)])
}

/** all-empty sub-editor content counts as "still no text" for a paras:[] box */
function boxStillEmpty(paras: TextboxDisplay['paras']): boolean {
  return paras.every((p) => p.runs.every((r) => r.text === ''))
}

/**
 * Per-box, per-paragraph rich-run patches on an original textbox block,
 * addressed by the box's txbxIndex (or shapeId for a box gaining its first
 * text); null when untouched. Unchanged paragraphs stay null so the engine
 * keeps their original bytes.
 */
function textboxParasPatch(node: PmNode, original: Block): TextboxParasPatchSet | null {
  const current = node.attrs?.textboxes as TextboxDisplay[] | null
  const orig = original.textboxes
  if (!current || !orig || current.length !== orig.length) return null
  const byIndex: (TextboxParaPatch | null)[][] = []
  const inject: Array<{ shapeId: string; paras: TextboxParaPatch[] }> = []
  current.forEach((box, i) => {
    const origBox = orig[i]
    const origParas = origBox.paras
    const same =
      box.paras.length === origParas.length &&
      box.paras.every((p, j) => textboxParaSignature(p) === textboxParaSignature(origParas[j]))
    if (same || (origParas.length === 0 && boxStillEmpty(box.paras))) return
    if (origBox.txbxIndex !== undefined) {
      byIndex[origBox.txbxIndex] = box.paras.map((p, j) => {
        if (
          j < origParas.length &&
          textboxParaSignature(p) === textboxParaSignature(origParas[j])
        ) {
          return null
        }
        // align is always explicit so a removed alignment also clears w:jc
        return { runs: mergeRuns(p.runs), align: p.align ?? null }
      })
    } else if (origBox.shapeId) {
      inject.push({
        shapeId: origBox.shapeId,
        paras: box.paras.map((p) => ({ runs: mergeRuns(p.runs), align: p.align ?? null })),
      })
    }
  })
  if (byIndex.length === 0 && inject.length === 0) return null
  return {
    ...(byIndex.length > 0 ? { byIndex } : {}),
    ...(inject.length > 0 ? { inject } : {}),
  }
}

function textboxSizesPatch(node: PmNode, original: Block): (TextboxSizePatch | null)[] | null {
  const current = node.attrs?.textboxes as TextboxDisplay[] | null
  const initial = original.textboxes
  if (!current || !initial || current.length !== initial.length) return null
  let changed = false
  const sizes = current.map((box, index) => {
    // a resize on an autofit box (no initial heightPx) pins its height:
    // patchTextboxSizes drops spAutoFit so Word honors the fixed extent
    const wPx = box.widthPx && box.widthPx !== initial[index].widthPx ? box.widthPx : null
    const hPx = box.heightPx && box.heightPx !== initial[index].heightPx ? box.heightPx : null
    if (wPx == null && hPx == null) return null
    changed = true
    return { wPx, hPx }
  })
  return changed ? sizes : null
}

function textboxStylesPatch(node: PmNode, original: Block): (ShapeStylePatch | null)[] | null {
  const current = node.attrs?.textboxes as TextboxDisplay[] | null
  const initial = original.textboxes
  if (!current || !initial || current.length !== initial.length) return null
  let changed = false
  const styles = current.map((box, index) => {
    const fillHex =
      (box.fill ?? null) !== (initial[index].fill ?? null) ? (box.fill ?? null) : undefined
    const borderHex =
      (box.borderColor ?? null) !== (initial[index].borderColor ?? null)
        ? (box.borderColor ?? null)
        : undefined
    if (fillHex === undefined && borderHex === undefined) return null
    changed = true
    return { fillHex, borderHex }
  })
  return changed ? styles : null
}

function fieldTextPatch(node: PmNode, original: Block): FieldTextPatch | null {
  const current = node.attrs?.fieldDisplay as FieldDisplay | null
  const initial = original.fieldDisplay
  if (!current || !initial || current.kind !== initial.kind) return null
  const patch: FieldTextPatch = {}
  if ((current.left ?? '') !== (initial.left ?? '')) patch.left = current.left ?? ''
  if (current.kind === 'tocLine' && (current.right ?? '') !== (initial.right ?? '')) {
    patch.right = current.right ?? ''
  }
  return Object.keys(patch).length > 0 ? patch : null
}

function formulaTokensPatch(node: PmNode, original: Block): string[] | null {
  const current = node.attrs?.formulaDisplay as FormulaDisplay | null
  const initial = original.formulaDisplay
  if (!current || !initial || current.tokens.length !== initial.tokens.length) return null
  return current.tokens.some((token, i) => token !== initial.tokens[i]) ? current.tokens : null
}

/**
 * High-fidelity pPr passthrough for in-place paragraph edits: text-only edits
 * reuse the original <w:pPr> bytes verbatim; format edits merge the six
 * format-model children into it (everything else keeps its bytes). Structure
 * changes (type / style / list) fall back to a full rebuild.
 */
function applyRawPPr(generated: GeneratedBlock, original: Block): void {
  const structureSame =
    original.type === generated.type &&
    (original.level ?? null) === (generated.level ?? null) &&
    (original.styleId ?? null) === (generated.styleId ?? null) &&
    JSON.stringify(original.list ?? null) === JSON.stringify(generated.list ?? null)
  if (original.rawPPr === undefined) {
    // a pPr-less paragraph laid out with the Normal style's character-unit
    // indents: an indent edit must also write their cancel attributes
    // (w:firstLineChars="0"…), which only mergePPrFormat knows how to do —
    // the generator's plain rebuild would let the style indent win on reload
    if (
      !original.format?.charIndents ||
      !structureSame ||
      generated.type !== 'paragraph' ||
      indentSame(original.format, generated.format)
    )
      return
    let built = mergePPrFormat('', generated.format, original.format)
    if (generated.pPrChange) built = setPPrChange(built, generated.pPrChange)
    generated.rawPPr = built
    return
  }
  if (!structureSame) return
  const formatSame =
    JSON.stringify(normalizedFormat(original.format)) ===
    JSON.stringify(normalizedFormat(generated.format))
  let rawPPr = formatSame
    ? original.rawPPr
    : mergePPrFormat(original.rawPPr, generated.format, original.format)
  // Strip pPrChange when the user accepted/rejected it (pPrChange: null in generated block)
  if (generated.pPrChange === null && original.pPrChangeInfo !== undefined) {
    rawPPr = stripPPrChange(rawPPr)
  } else if (generated.pPrChange) {
    rawPPr = setPPrChange(rawPPr, generated.pPrChange)
  }
  generated.rawPPr = rawPPr
}

const PAGINATION_FLAG_KEYS = [
  'keepNext',
  'keepLines',
  'widowControl',
  'suppressLineNumbers',
] as const

function nodeFormat(node: PmNode): ParaFormat | undefined {
  const format: ParaFormat = {}
  if (node.attrs?.align) format.align = node.attrs.align as ParaFormat['align']
  if (node.attrs?.lineSpacing) format.lineSpacing = Number(node.attrs.lineSpacing)
  if (node.attrs?.lineRule) format.lineRule = node.attrs.lineRule as ParaFormat['lineRule']
  if (node.attrs?.lineRawTwips) format.lineRawTwips = Number(node.attrs.lineRawTwips)
  // 0 kept: an explicit w:left="0" overrides a numbering-level indent
  if (node.attrs?.indentLeft != null) format.indentLeft = Number(node.attrs.indentLeft)
  if (node.attrs?.indentRight != null) format.indentRight = Number(node.attrs.indentRight)
  if (node.attrs?.indentFirstLine != null)
    format.indentFirstLine = Number(node.attrs.indentFirstLine)
  if (node.attrs?.spaceBefore != null) format.spaceBefore = Number(node.attrs.spaceBefore)
  if (node.attrs?.spaceAfter != null) format.spaceAfter = Number(node.attrs.spaceAfter)
  if (node.attrs?.spaceBeforeAuto != null)
    format.spaceBeforeAuto = node.attrs.spaceBeforeAuto as boolean
  if (node.attrs?.spaceAfterAuto != null)
    format.spaceAfterAuto = node.attrs.spaceAfterAuto as boolean
  if (node.attrs?.contextualSpacing != null)
    format.contextualSpacing = node.attrs.contextualSpacing as boolean
  for (const key of PAGINATION_FLAG_KEYS)
    if (node.attrs?.[key] != null) format[key] = node.attrs[key] as boolean
  if (node.attrs?.pageBreakBefore) format.pageBreakBefore = true
  if (node.attrs?.bidi) format.bidi = true
  if (node.attrs?.autoSpace != null) format.autoSpace = node.attrs.autoSpace as boolean
  if (node.attrs?.wordWrap != null) format.wordWrap = node.attrs.wordWrap as boolean
  if (node.attrs?.overflowPunct != null) {
    format.overflowPunct = node.attrs.overflowPunct as boolean
  }
  if (node.attrs?.eaLang) format.eastAsiaLang = String(node.attrs.eaLang)
  if (node.attrs?.shadingFill) format.shadingFill = String(node.attrs.shadingFill)
  if (node.attrs?.borders) format.borders = String(node.attrs.borders)
  if (node.attrs?.borderLines) {
    try {
      const parsed = JSON.parse(String(node.attrs.borderLines))
      if (parsed && typeof parsed === 'object' && Object.keys(parsed).length > 0) {
        format.borderLines = parsed
      }
    } catch {
      /* ignore malformed */
    }
  }
  if (node.attrs?.tabStops) {
    try {
      const parsed = JSON.parse(String(node.attrs.tabStops))
      if (Array.isArray(parsed) && parsed.length > 0) format.tabStops = parsed
    } catch {
      /* ignore malformed */
    }
  }
  if (node.attrs?.dropCap) {
    try {
      const parsed = JSON.parse(String(node.attrs.dropCap))
      if (parsed && typeof parsed === 'object') format.dropCap = parsed
    } catch {
      /* ignore malformed */
    }
  }
  if (node.attrs?.frame) {
    try {
      const parsed = JSON.parse(String(node.attrs.frame))
      if (parsed && typeof parsed === 'object' && parsed.wTwips > 0) format.frame = parsed
    } catch {
      /* ignore malformed */
    }
  }
  if (node.attrs?.textDirection) {
    const dir = String(node.attrs.textDirection)
    if (dir === 'tbRl' || dir === 'tbRlV' || dir === 'btLr') format.textDirection = dir
  }
  if (node.attrs?.emptyRunSize) format.emptyRunSizeHalfPoints = Number(node.attrs.emptyRunSize)
  if (node.attrs?.emptyRunFont) format.emptyRunFontFamily = String(node.attrs.emptyRunFont)
  if (node.attrs?.markSize) format.markSizeHalfPoints = Number(node.attrs.markSize)
  return Object.keys(format).length > 0 ? format : undefined
}

export function pmNodeToGeneratedBlock(node: PmNode): GeneratedBlock {
  const runs = inlineToRuns(node.content ?? [])
  const format = nodeFormat(node)
  const bookmarks = nodeBookmarks(node, 'bookmarks')
  const hiddenBookmarks = nodeBookmarks(node, 'hiddenBookmarks')
  const commentStarts = nodeBookmarks(node, 'commentStarts')
  const commentEnds = nodeBookmarks(node, 'commentEnds')
  if (node.type === 'docHeading') {
    const pPrChange = node.attrs?.pPrChange as string | null | undefined
    const blockRevision = node.attrs?.blockRevision as GeneratedBlock['blockRevision']
    return {
      type: 'heading',
      level: Number(node.attrs?.level) || 1,
      ...(node.attrs?.outlineOnly ? { outlineOnly: true } : {}),
      styleId: (node.attrs?.styleId as string) ?? undefined,
      format,
      bookmarks,
      hiddenBookmarks,
      commentStarts,
      commentEnds,
      runs,
      ...(pPrChange !== undefined ? { pPrChange: pPrChange ?? null } : {}),
      ...(blockRevision !== undefined ? { blockRevision: blockRevision ?? null } : {}),
    }
  }
  if (node.type === 'docListItem') {
    const numId = node.attrs?.numId as string | null
    if (numId) {
      const pPrChange = node.attrs?.pPrChange as string | null | undefined
      const blockRevision = node.attrs?.blockRevision as GeneratedBlock['blockRevision']
      return {
        type: 'listItem',
        styleId: (node.attrs?.styleId as string) ?? undefined,
        list: {
          kind: (node.attrs?.kind as 'bullet' | 'ordered') ?? 'bullet',
          numId,
          ilvl: Number(node.attrs?.ilvl) || 0,
        },
        format,
        bookmarks,
        hiddenBookmarks,
        commentStarts,
        commentEnds,
        runs,
        ...(pPrChange !== undefined ? { pPrChange: pPrChange ?? null } : {}),
        ...(blockRevision !== undefined ? { blockRevision: blockRevision ?? null } : {}),
      }
    }
    // list item without a docx numbering id degrades to a plain paragraph with marker
    return {
      type: 'paragraph',
      format,
      bookmarks,
      hiddenBookmarks,
      commentStarts,
      commentEnds,
      runs: [{ text: '• ' }, ...runs],
    }
  }
  // Extract sdtShell from node attrs (serialized as JSON)
  const sdtShell = node.attrs?.sdtShell
    ? (() => {
        try {
          const parsed = JSON.parse(String(node.attrs.sdtShell))
          return parsed && typeof parsed === 'object' ? parsed : undefined
        } catch {
          return undefined
        }
      })()
    : undefined
  // pPrChange: propagate to GeneratedBlock so applyRawPPr can strip it when accepted
  const pPrChange = node.attrs?.pPrChange as string | null | undefined
  const blockRevision = node.attrs?.blockRevision as GeneratedBlock['blockRevision']
  return {
    type: 'paragraph',
    styleId: (node.attrs?.styleId as string) ?? undefined,
    format,
    bookmarks,
    hiddenBookmarks,
    commentStarts,
    commentEnds,
    runs,
    ...(sdtShell ? { sdtShell } : {}),
    // carry pPrChange through: null means "user accepted/rejected, strip from rawPPr"
    ...(pPrChange !== undefined ? { pPrChange: pPrChange ?? null } : {}),
    ...(blockRevision !== undefined ? { blockRevision: blockRevision ?? null } : {}),
  }
}

function nodeBookmarks(
  node: PmNode,
  attr: 'bookmarks' | 'hiddenBookmarks' | 'commentStarts' | 'commentEnds',
): string[] | undefined {
  const value = node.attrs?.[attr] as string[] | null | undefined
  return Array.isArray(value) && value.length > 0 ? value : undefined
}

export function inlineToRuns(content: PmNode[]): Run[] {
  const runs: Run[] = []
  for (const node of content) {
    if (node.type === 'hardBreak') {
      const ch = node.attrs?.pageBreak
        ? '\f'
        : node.attrs?.colBreak
          ? '\v'
          : node.attrs?.wrapBreak
            ? '\u001e'
            : '\n'
      const prev = runs[runs.length - 1]
      const prevAtomic =
        prev &&
        (prev.noteRef ||
          prev.xeTerm !== undefined ||
          prev.math ||
          prev.sym ||
          prev.ruby ||
          prev.image)
      if (node.marks?.length) runs.push(runFromMarks(ch, node.marks))
      else if (prev && !prevAtomic) prev.text += ch
      else runs.push({ text: ch })
      continue
    }
    if (node.type === 'docNoteRef') {
      runs.push({
        text: String(node.attrs?.num ?? 1),
        noteRef: {
          kind: (node.attrs?.kind as 'footnote' | 'endnote') ?? 'footnote',
          id: String(node.attrs?.id ?? ''),
        },
      })
      continue
    }
    if (node.type === 'docXeMark') {
      runs.push({ text: '', xeTerm: String(node.attrs?.term ?? '') })
      continue
    }
    if (node.type === 'docInlineMath') {
      const omml = String(node.attrs?.omml ?? '')
      if (omml) runs.push({ text: String(node.attrs?.text ?? ''), math: { omml } })
      continue
    }
    if (node.type === 'docRuby') {
      const base = String(node.attrs?.base ?? '')
      const xml = String(node.attrs?.xml ?? '')
      if (base) {
        runs.push(
          xml ? { text: base, ruby: { rt: String(node.attrs?.rt ?? ''), xml } } : { text: base },
        )
      }
      continue
    }
    if (node.type === 'docInlineImage') {
      const dataUrl = String(node.attrs?.dataUrl ?? '')
      const xml = String(node.attrs?.xml ?? '')
      const rule = node.attrs?.rule as NonNullable<Run['image']>['rule'] | null | undefined
      if ((dataUrl || rule) && xml) {
        runs.push({
          text: '',
          ...(node.attrs?.rawRPr ? { rawRPr: String(node.attrs.rawRPr) } : {}),
          image: {
            dataUrl,
            xml,
            ...(rule ? { rule } : {}),
            ...(node.attrs?.widthPx ? { widthPx: Number(node.attrs.widthPx) } : {}),
            ...(node.attrs?.heightPx ? { heightPx: Number(node.attrs.heightPx) } : {}),
            ...(node.attrs?.rotDeg ? { rotDeg: Number(node.attrs.rotDeg) } : {}),
            ...(node.attrs?.flipH ? { flipH: true } : {}),
            ...(node.attrs?.flipV ? { flipV: true } : {}),
          },
        })
      }
      continue
    }
    if (node.type !== 'text' || !node.text) continue
    runs.push(...symRuns(runFromMarks(node.text, node.marks ?? [])))
  }
  return mergeRuns(runs)
}

/** an installed symbol font draws the w:sym glyph itself (the exact Word shape), so the
 *  editor text is the private-use code the font's cmap indexes; the Unicode stand-in is
 *  for machines without the font. Display only: runFromMarks turns it back. */
function symDisplayText(run: Run): string {
  if (!run.sym) return run.text
  const pua = symbolPuaChar(run.sym.char)
  return pua && run.text !== pua && symbolFontCovers(run.sym.font, pua) ? pua : run.text
}

/** a w:sym run is one glyph; text edited or merged under its mark goes back as plain runs */
function symRuns(run: Run): Run[] {
  if (!run.sym) return [run]
  const glyph = symbolGlyph(run.sym.font, run.sym.char)
  if (run.text === glyph) return [run]
  const { sym, ...plain } = run
  return [...run.text].map((ch) => ({ ...plain, text: ch, ...(ch === glyph ? { sym } : {}) }))
}

function runFromMarks(text: string, marks: PmMark[]): Run {
  const run: Run = { text }
  for (const mark of marks) {
    if (mark.type === 'bold') run.bold = true
    else if (mark.type === 'italic') run.italic = true
    else if (mark.type === 'underline') run.underline = true
    else if (mark.type === 'strike') run.strike = true
    else if (mark.type === 'link') {
      run.link = {
        href: String(mark.attrs?.href ?? ''),
        rId: (mark.attrs?.rId as string) ?? undefined,
        tooltip: (mark.attrs?.tooltip as string) ?? undefined,
        ...(mark.attrs?.plain ? { plain: true } : {}),
      }
    } else if (mark.type === 'refField') {
      run.refField = String(mark.attrs?.name ?? '')
      if (mark.attrs?.instr) run.refInstr = String(mark.attrs.instr)
      if (mark.attrs?.dirty === true) run.fldDirty = true
    } else if (mark.type === 'docSym') {
      run.sym = { font: String(mark.attrs?.font ?? ''), char: String(mark.attrs?.char ?? '') }
      if (run.text === symbolPuaChar(run.sym.char))
        run.text = symbolGlyph(run.sym.font, run.sym.char) ?? run.text
    } else if (mark.type === 'instrField') {
      run.instrField = String(mark.attrs?.instr ?? '')
      if (mark.attrs?.beginXml) run.fldBeginXml = String(mark.attrs.beginXml)
      if (mark.attrs?.dirty === true) run.fldDirty = true
      if (/^\s*(?:ADDIN\s+)?(?:ZOTERO_|CSL_)/i.test(run.instrField)) {
        const fieldId = Number(mark.attrs?.fieldId)
        if (Number.isSafeInteger(fieldId) && fieldId > 0) run.zoteroFieldId = fieldId
        const fieldPart = mark.attrs?.fieldPart
        if (
          fieldPart === 'single' ||
          fieldPart === 'begin' ||
          fieldPart === 'inside' ||
          fieldPart === 'end'
        ) {
          run.zoteroFieldPart = fieldPart
        }
      }
    } else if (mark.type === 'ctrlCheckbox') {
      run.sdtCheckboxXml = String(mark.attrs?.sdtPr ?? '')
    } else if (mark.type === 'comment') {
      const ids = String(mark.attrs?.ids ?? '')
        .split(' ')
        .filter(Boolean)
      if (ids.length > 0) run.commentIds = ids
    } else if (mark.type === 'ins' || mark.type === 'del') {
      const info: NonNullable<Run['ins']> = { author: String(mark.attrs?.author ?? '') }
      if (mark.attrs?.date) info.date = String(mark.attrs.date)
      if (mark.attrs?.id) info.id = String(mark.attrs.id)
      if (mark.type === 'ins') run.ins = info
      else run.del = info
    } else if (mark.type === 'docTextStyle') {
      if (mark.attrs?.color) run.color = String(mark.attrs.color)
      if (mark.attrs?.sizeHalfPoints) run.sizeHalfPoints = Number(mark.attrs.sizeHalfPoints)
      if (mark.attrs?.font) run.font = String(mark.attrs.font)
      if (mark.attrs?.fontAscii) run.fontAscii = String(mark.attrs.fontAscii)
      if (mark.attrs?.eastAsiaFont) run.eastAsiaFont = String(mark.attrs.eastAsiaFont)
      if (mark.attrs?.csFont) run.csFont = String(mark.attrs.csFont)
      if (mark.attrs?.charSpacingTwips != null)
        run.charSpacingTwips = Number(mark.attrs.charSpacingTwips)
      if (mark.attrs?.charScalePct) run.charScalePct = Number(mark.attrs.charScalePct)
      if (mark.attrs?.kernHalfPoints != null) run.kernHalfPoints = Number(mark.attrs.kernHalfPoints)
      if (mark.attrs?.highlight) run.highlight = String(mark.attrs.highlight)
      if (mark.attrs?.shading) run.shading = String(mark.attrs.shading)
      if (mark.attrs?.textOutline) {
        const outline = parseTextOutlineAttr(String(mark.attrs.textOutline))
        if (outline) run.textOutline = outline
      }
      const effect = parseTextEffectAttr(mark.attrs?.textEffect)
      if (effect) run.textEffect = effect
      if (mark.attrs?.dstrike != null) run.dstrike = Boolean(mark.attrs.dstrike)
      if (mark.attrs?.glow) {
        const glow = parseGlowAttr(String(mark.attrs.glow))
        if (glow) run.glow = glow
      }
      if (mark.attrs?.positionHalfPoints)
        run.positionHalfPoints = Number(mark.attrs.positionHalfPoints)
      if (mark.attrs?.bdr) {
        const bdr = parseRunBorderAttr(String(mark.attrs.bdr))
        if (bdr) run.bdr = bdr
      }
      if (mark.attrs?.vertAlign === 'superscript' || mark.attrs?.vertAlign === 'subscript') {
        run.vertAlign = mark.attrs.vertAlign
      }
      if (mark.attrs?.em) run.em = mark.attrs.em as NonNullable<Run['em']>
      if (mark.attrs?.cs) run.cs = true
      if (mark.attrs?.vanish) run.vanish = true
      if (mark.attrs?.vanishOwn != null) run.vanishOwn = Boolean(mark.attrs.vanishOwn)
      if (mark.attrs?.caps) run.caps = mark.attrs.caps as Run['caps']
      if (mark.attrs?.eaLang) run.eastAsiaLang = String(mark.attrs.eaLang)
      if (mark.attrs?.rtl != null) run.rtl = Boolean(mark.attrs.rtl)
      if (mark.attrs?.styleId) run.styleId = String(mark.attrs.styleId)
      if (mark.attrs?.rawRPr) run.rawRPr = String(mark.attrs.rawRPr)
      if (mark.attrs?.themeRFonts) {
        run.themeRFonts = JSON.parse(String(mark.attrs.themeRFonts)) as Run['themeRFonts']
      }
      if (mark.attrs?.themeColor) run.themeColor = String(mark.attrs.themeColor)
    } else if (mark.type === 'rprChange') {
      run.rPrChange = {
        author: String(mark.attrs?.author ?? ''),
        ...(mark.attrs?.date ? { date: String(mark.attrs.date) } : {}),
        ...(mark.attrs?.id ? { id: String(mark.attrs.id) } : {}),
        ...(mark.attrs?.old ? { old: mark.attrs.old as NonNullable<Run['rPrChange']>['old'] } : {}),
      }
    }
  }
  return run
}

function mergeRuns(runs: Run[]): Run[] {
  const merged: Run[] = []
  for (const run of runs) {
    const prev = merged[merged.length - 1]
    // reference markers / index entries / inline math / fields are atomic and
    // never merge — two identical FORMCHECKBOX runs must stay two fields
    const atomic =
      run.noteRef ||
      run.xeTerm !== undefined ||
      run.instrField !== undefined ||
      run.sdtCheckboxXml !== undefined ||
      run.math ||
      run.sym ||
      run.ruby ||
      run.image ||
      prev?.noteRef ||
      prev?.xeTerm !== undefined ||
      prev?.instrField !== undefined ||
      prev?.sdtCheckboxXml !== undefined ||
      prev?.math ||
      prev?.sym ||
      prev?.ruby ||
      prev?.image
    if (prev && !atomic && runStyleKey(prev) === runStyleKey(run)) prev.text += run.text
    else merged.push({ ...run })
  }
  return merged
}

/** Font-dialog fields that save from the model (mergeRPrModel rebuilds them when edited) */
function charFormatKey(run: Run): unknown[] {
  return [
    run.charSpacingTwips ?? null,
    run.charScalePct ?? null,
    run.kernHalfPoints ?? null,
    run.positionHalfPoints ?? null,
    run.caps ?? null,
    run.dstrike ?? null,
    run.vanishOwn ?? null,
  ]
}

function runStyleKey(run: Run): string {
  return JSON.stringify([
    run.rawRPr ?? null,
    run.styleId ?? null,
    !!run.bold,
    !!run.italic,
    !!run.underline,
    !!run.strike,
    run.color ?? null,
    run.sizeHalfPoints ?? null,
    run.font ?? null,
    run.fontAscii ?? null,
    run.eastAsiaFont ?? null,
    run.highlight ?? null,
    run.shading ?? null,
    ...charFormatKey(run),
    run.textOutline ? JSON.stringify(run.textOutline) : null,
    run.vertAlign ?? null,
    run.link?.href ?? null,
    run.link?.rId ?? null,
    run.commentIds?.join(' ') ?? null,
    revisionKey(run),
    run.noteRef ? [run.noteRef.kind, run.noteRef.id] : null,
    run.xeTerm ?? null,
    run.refField ?? null,
    run.instrField ?? null,
    run.fldBeginXml ?? null,
    run.fldDirty ?? null,
    run.sdtCheckboxXml ?? null,
    run.math?.omml ?? null,
    run.sym ? [run.sym.font, run.sym.char] : null,
    run.ruby?.xml ?? null,
  ])
}

function revisionKey(run: Run): string | null {
  if (!run.ins && !run.del) return null
  return JSON.stringify([
    run.ins?.author ?? null,
    run.ins?.date ?? null,
    run.ins?.id ?? null,
    run.del?.author ?? null,
    run.del?.date ?? null,
    run.del?.id ?? null,
  ])
}

// ---- signatures: "did the user actually change this block?" ----

function normalizedRuns(runs: Run[]): unknown[] {
  return mergeRuns(runs).map((r) =>
    // a picture/rule run's rPr-derived display fields do not survive the image
    // node round trip; its identity is the fragment plus the raw rPr it re-emits
    r.text === '' && r.image
      ? ['image', r.rawRPr ?? null, r.image.xml, r.link?.href ?? null, revisionKey(r)]
      : [
          r.text,
          r.rawRPr ?? null,
          r.styleId ?? null,
          !!r.bold,
          !!r.italic,
          !!r.underline,
          !!r.strike,
          r.color ?? null,
          r.sizeHalfPoints ?? null,
          r.font ?? null,
          r.fontAscii ?? null,
          r.eastAsiaFont ?? null,
          r.highlight ?? null,
          ...charFormatKey(r),
          r.vertAlign ?? null,
          r.link?.href ?? null,
          r.commentIds?.join(' ') ?? null,
          revisionKey(r),
          r.noteRef ? [r.noteRef.kind, r.noteRef.id] : null,
          r.xeTerm ?? null,
          r.refField ?? null,
          r.instrField ?? null,
          r.zoteroFieldPart ?? null,
          r.fldBeginXml ?? null,
          r.fldDirty ?? null,
          r.sdtCheckboxXml ?? null,
          r.math?.omml ?? null,
          r.sym ? [r.sym.font, r.sym.char] : null,
          r.ruby?.xml ?? null,
          r.image?.xml ?? null,
        ],
  )
}

/** same indent twips in two paragraph formats (an explicit 0 differs from unset, as emitted) */
function indentSame(a: ParaFormat | undefined, b: ParaFormat | undefined): boolean {
  const norm = (v: number | undefined) => (v !== undefined ? Math.round(v) : null)
  return (
    norm(a?.indentLeft) === norm(b?.indentLeft) &&
    norm(a?.indentRight) === norm(b?.indentRight) &&
    norm(a?.indentFirstLine) === norm(b?.indentFirstLine)
  )
}

function normalizedFormat(format: ParaFormat | undefined): unknown {
  if (!format) return null
  return [
    format.align ?? null,
    format.lineSpacing ?? null,
    format.lineRule ?? null,
    format.lineRawTwips ?? null,
    format.bidi ?? false,
    format.indentLeft ?? null,
    format.indentRight ?? null,
    format.indentFirstLine ?? null,
    format.spaceBefore ?? null,
    format.spaceAfter ?? null,
    format.pageBreakBefore ?? false,
    format.keepNext ?? null,
    format.keepLines ?? null,
    format.widowControl ?? null,
    format.suppressLineNumbers ?? null,
    format.contextualSpacing ?? null,
    format.shadingFill ?? null,
    format.borders ?? null,
    format.borderLines ? JSON.stringify(format.borderLines) : null,
    format.tabStops ? JSON.stringify(format.tabStops) : null,
    format.dropCap ? JSON.stringify(format.dropCap) : null,
    format.autoSpace ?? null,
    format.wordWrap ?? null,
    format.overflowPunct ?? null,
    format.emptyRunSizeHalfPoints ?? null,
  ]
}

export function signatureOfBlock(block: Block): string {
  return JSON.stringify({
    type: block.type,
    level: block.level ?? null,
    styleId: block.styleId ?? null,
    list: block.list ? [block.list.kind, block.list.numId, block.list.ilvl] : null,
    format: normalizedFormat(block.format),
    bookmarks: block.bookmarks ?? null,
    hiddenBookmarks: block.hiddenBookmarks ?? null,
    commentStarts: block.commentStarts ?? null,
    commentEnds: block.commentEnds ?? null,
    runs: normalizedRuns(block.runs ?? []),
    // include pPrChangeInfo in signature: present = has pPrChange, absent = clean
    pPrChange: block.pPrChangeInfo ? JSON.stringify(block.pPrChangeInfo) : null,
    blockRevision: block.blockRevision ?? null,
  })
}

export function signatureOfGenerated(block: GeneratedBlock): string {
  return JSON.stringify({
    type: block.type,
    level: block.level ?? null,
    styleId: block.styleId ?? null,
    list: block.list ? [block.list.kind, block.list.numId, block.list.ilvl] : null,
    format: normalizedFormat(block.format),
    bookmarks: block.bookmarks ?? null,
    hiddenBookmarks: block.hiddenBookmarks ?? null,
    commentStarts: block.commentStarts ?? null,
    commentEnds: block.commentEnds ?? null,
    runs: normalizedRuns(block.runs),
    // include pPrChange in signature: null (accepted/rejected) != non-null (has pPrChange)
    pPrChange: block.pPrChange ?? null,
    blockRevision: block.blockRevision ?? null,
  })
}
