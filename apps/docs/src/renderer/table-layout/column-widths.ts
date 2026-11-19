import type { TableModel } from '@genoffice/docx-engine'
import type { FontMetricsProvider } from '../line-metrics'
import { defaultTextMeasure, textMeasureOf, type RunDefaults, type TextMeasure } from './measure'
import { DEFAULT_CELL_MAR, MIN_COL_TWIPS } from './constants'
import {
  maxContentColTwips,
  minContentColTwips,
  sideMarginTwips,
  tableIndentTwips,
} from './content-widths'
import {
  distributeContentAutofit,
  nestedFloorTwips,
  reclaimFloorTwips,
  reclaimSurplus,
  scaleToFit,
} from './distribute'

/** twips the table may grow into (`avail`) and the text column it should fit (`fit`) */
export interface ColumnBudget {
  avail: number
  fit: number
}

export interface ColumnLayoutOptions {
  /** laid out inside a cell: fits the cell, floors at the side margins, never holds a dxa width */
  nested?: boolean
  /** compat < 15: the box hangs by the side cell margins (w:tblInd measured to the cell text) */
  legacy?: boolean
  defaults?: RunDefaults
}

export type ColumnPhase = 'pct-resolve' | 'content-autofit' | 'grow-min' | 'reclaim' | 'scale'

/** pure output of layoutTableColumns */
export interface ColumnWidths {
  /** the working grid: the model's own, a pct table's resolved shares, or the autofit result */
  widths: number[] | undefined
  /** `widths` differ from the model's grid and must be drawn */
  changed: boolean
  /** a pct table was resolved to absolute widths (drawn as such only when `changed`) */
  resolvedPct: boolean
  /** growth bound less the table indent (nested tables are laid out within it) */
  budget: number
  phases: ColumnPhase[]
}

/**
 * Word autofit (tblW auto/absent, no fixed tblLayout): a column grows to its
 * min-content — the widest unbreakable word — so narrow declared widths don't
 * shatter words char by char (heuristic estimate, see maxWordWidthPx). Growth
 * past the fit width, like a declared total already past it, is reclaimed
 * proportionally from columns with surplus over their own min-content; any
 * remainder spills and is bounded by clampColumnWidths. Phases:
 *   1 declared grid + autofit gate, 2 pct resolution (+ content autofit without tcW),
 *   3 min-content, 4 kept total (dxa held / shrunk to the column / nested cell),
 *   5 grow to min-content and reclaim the overflow, 6 scale when the minimums alone overflow.
 */
export function layoutTableColumns(
  model: TableModel,
  budgetIn: ColumnBudget,
  measure: TextMeasure,
  opts: ColumnLayoutOptions = {},
): ColumnWidths {
  const { nested = false, legacy = false, defaults } = opts
  const availTwips = budgetIn.avail
  const fitTwips = budgetIn.fit
  const indent = tableIndentTwips(model)
  const budget = availTwips - indent
  const phases: ColumnPhase[] = []
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
  const sideMar = sideMarginTwips(model)
  // compat < 15: a pct table spans the text column plus both side cell margins,
  // whatever its jc or w:tblInd (probe 2026-11-09: pct 5000 -> 9026 + 216, pct 2500
  // -> half of it, tblCellMar 200 -> 9026 + 400; compat 15 stays on the column)
  const legacyPctHang = legacy && !nested && !model.floatSide ? sideMar : 0
  if (autofit && model.widthPct && model.colWidthsPct?.length && budget > 0) {
    const tableW = ((Math.min(fitTwips, availTwips) + legacyPctHang) * model.widthPct) / 100
    if (tableW > 0) {
      widths = model.colWidthsPct.map((p) => Math.round((p / 100) * tableW))
      resolvedPct = true
      phases.push('pct-resolve')
      // the hung width only reaches the render path as absolute columns
      if (legacyPctHang > 0) widthsChanged = true
      if (model.contentAutofit) {
        const n = widths.length
        const mins = minContentColTwips(model, n, measure, false, defaults)
        const maxes = maxContentColTwips(model, n, measure, defaults)
        if (maxes.some((m) => m > 0)) {
          // an empty column still keeps its side margins
          widths = distributeContentAutofit(
            Math.round(tableW),
            mins.map((m) => Math.max(m, Math.min(sideMar, Math.round(tableW) / n))),
            maxes.map((m, i) => Math.max(m, mins[i])),
          )
          widthsChanged = true
          phases.push('content-autofit')
        }
      }
    }
  }
  if (autofit && (resolvedPct || !model.widthPct) && widths?.length && budget > 0) {
    const minsRaw = minContentColTwips(model, widths.length, measure, !!model.layoutGrid, defaults)
    const mins = minsRaw.map((m) => Math.min(m, budget))
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
    // cell padding, so it never hangs. An explicit dxa tblW is honoured
    // (measured: landscape and portrait forms 7-22% wider than the column,
    // negative tblInd, drawn at full width; a left-aligned 531 pt grid on a
    // 451 pt column ran 8 pt past the paper edge, clipped, rows single-line).
    const leftAligned = model.align !== 'center' && model.align !== 'right'
    const legacyHang = !legacy
      ? 0
      : leftAligned && !model.floatSide
        ? (model.cellMarTwips?.right ?? DEFAULT_CELL_MAR)
        : sideMar
    const box = fit - (leftAligned ? (model.indentTwips ?? 0) : 0) + legacyHang
    const hang = leftAligned ? Math.min(model.indentTwips ?? 0, 0) : 0
    // an explicit dxa tblW keeps its declared total whatever the jc: a centred
    // 12000/13000-twip table on a 9360 column is centred on the margin box and
    // clipped by the paper, a right-aligned one pins its right edge to the
    // margin and runs off the paper on the left (probe 2026-10-29)
    const explicitDxa =
      !model.autoLayout &&
      !nested &&
      ((model.dxaWidth && !model.floatSide) || declared + hang <= budget)
    // a resolved pct table keeps the width it resolved to (the legacy hang included)
    const kept =
      resolvedPct || nested
        ? target + (resolvedPct ? legacyPctHang : 0)
        : Math.max(target, explicitDxa ? declared : Math.min(declared, box))
    if (mins.some((m, i) => m > widths![i]) || (!resolvedPct && declared > kept)) {
      let grown = widths.map((w, i) => Math.max(w, mins[i]))
      phases.push('grow-min')
      const overflow = grown.reduce((a, b) => a + b, 0) - kept
      if (overflow > 0) {
        const floors = grown.map((w, i) =>
          nested ? nestedFloorTwips(w, mins[i], sideMar) : reclaimFloorTwips(w, mins[i]),
        )
        grown = reclaimSurplus(grown, floors, overflow)
        phases.push('reclaim')
        // min-contents alone can exceed the budget (one huge unbreakable
        // token): the text widths compress over the side margins like Word
        // instead of handing the excess to clampColumnWidths, which zeroes
        // trailing columns
        // the scale works from the uncapped min-contents: a token wider than
        // the column still weighs its full width (probe: 60 m / a / a)
        const uncapped = grown.map((w, i) => Math.max(w, minsRaw[i]))
        const scaled = scaleToFit(
          uncapped,
          kept,
          grown.map(() => sideMar),
        )
        if (scaled !== uncapped) {
          phases.push('scale')
          grown = scaled
        }
      }
      widths = grown
      widthsChanged = true
    }
  }
  return { widths, changed: widthsChanged, resolvedPct, budget, phases }
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
 * garbage gridCol pushes content megapixels off page. Returns null when the
 * grid fits.
 */
export function clampColumnWidths(
  model: TableModel,
  capTwips: number,
  measure: TextMeasure,
): { widths: number[]; budget: number } | null {
  const budget = capTwips - tableIndentTwips(model)
  let widths = model.colWidthsTwips
  if (!widths || budget <= 0 || widths.reduce((a, b) => a + b, 0) <= budget) return null
  const mins = minContentColTwips(model, widths.length, measure).map((m, i) =>
    Math.min(Math.max(m, MIN_COL_TWIPS), budget, widths![i]),
  )
  const overflow = widths.reduce((a, b) => a + b, 0) - budget
  widths = reclaimSurplus(widths, mins, overflow)
  // min-contents alone can exceed the cap: compress proportionally
  widths = scaleToFit(widths, budget)
  return { widths, budget }
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
 * Model-level autofit: layoutTableColumns on the table and, recursively, on
 * every nested table within its cell width. Display-only, like the clamp:
 * untouched tables still save their original bytes, structural rebuilds
 * regenerate the grid from these display widths (same as user column resizes).
 */
export function expandAutofitColWidths(
  model: TableModel,
  availTwips: number,
  fitTwips: number = availTwips,
  metrics: FontMetricsProvider | null | undefined = undefined,
  nested = false,
  legacy = false,
  defaults?: RunDefaults,
): TableModel {
  const measure = metrics === undefined ? defaultTextMeasure() : textMeasureOf(metrics)
  return expandWithMeasure(model, { avail: availTwips, fit: fitTwips }, measure, {
    nested,
    legacy,
    defaults,
  })
}

function expandWithMeasure(
  model: TableModel,
  budget: ColumnBudget,
  measure: TextMeasure,
  opts: ColumnLayoutOptions,
): TableModel {
  const laid = layoutTableColumns(model, budget, measure, opts)
  // nested tables fit the cells of the working grid (a pct table's resolved shares
  // bound them even when nothing grew)
  const { rows, changed: rowsChanged } = mapNestedTables(
    model,
    laid.widths,
    laid.budget,
    (nt, avail) =>
      expandWithMeasure(nt, { avail, fit: avail }, measure, {
        nested: true,
        defaults: opts.defaults,
      }),
  )
  if (!laid.changed && !rowsChanged) return model
  const result = withColWidths(model, rows, laid.changed ? laid.widths : undefined)
  if (laid.changed && laid.resolvedPct) delete result.widthPct
  return result
}

/** model-level paper clamp (clampColumnWidths), nested tables clamped to their cell */
export function clampTableColWidths(
  model: TableModel,
  capTwips: number,
  metrics: FontMetricsProvider | null | undefined = undefined,
): TableModel {
  const measure = metrics === undefined ? defaultTextMeasure() : textMeasureOf(metrics)
  return clampWithMeasure(model, capTwips, measure)
}

function clampWithMeasure(model: TableModel, capTwips: number, measure: TextMeasure): TableModel {
  const clamped = clampColumnWidths(model, capTwips, measure)
  const budget = clamped?.budget ?? capTwips - tableIndentTwips(model)
  const widths = clamped?.widths ?? model.colWidthsTwips
  const { rows, changed: rowsChanged } = mapNestedTables(model, widths, budget, (nt, cap) =>
    clampWithMeasure(nt, cap, measure),
  )
  if (!clamped && !rowsChanged) return model
  return withColWidths(model, rows, clamped?.widths)
}
