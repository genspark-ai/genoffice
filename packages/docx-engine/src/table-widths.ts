// Table column widths as the document declares them: the one place in the engine
// that reads w:tblGrid / w:tcW / w:tblW / w:tblLayout and arbitrates between them.
// The renderer's table-layout/ module takes the result from the table model.
import {
  CELL_SPACING_GAPS_PER_COLUMN,
  GRID_ROUNDING_TWIPS_PER_COLUMN,
  SAVED_GRID_TOLERANCE,
  TCW_GRID_PCT_TOLERANCE,
  UNIFORM_GRID_MIN_TWIPS,
  UNIFORM_GRID_TOLERANCE,
} from './table-width-thresholds'
import type { CellMargins, SectionSettings, TableCell, TableModel } from './types'
import { attrsOf, childrenThroughSdt, findChild, findChildren, type XNode } from './xml-utils'

/**
 * Per-column widths reconstructed from cell w:tcW (dxa) across all rows: the widest
 * un-spanned cell per grid slot wins (Word widens a column to fit later rows, never
 * narrows it). Undefined unless every column got a value — partial data would skew
 * the ratio worse than the tblGrid fallback.
 */
function tcwColumnWidths(tbl: XNode, pctTableTwips = 0): number[] | undefined {
  const cols: number[] = []
  let colCount = 0
  for (const tr of childrenThroughSdt(tbl, 'w:tr')) {
    const edges = rowGridEdges(tr)
    let idx = edges.before
    for (const tc of childrenThroughSdt(tr, 'w:tc')) {
      const tcPr = findChild(tc, 'w:tcPr')
      const span = Math.max(
        1,
        Number(attrsOf(findChild(tcPr ?? {}, 'w:gridSpan') ?? {})['w:val']) || 1,
      )
      // duplicated w:tcW: Word keeps the last occurrence (generators leave stale first values)
      const a = attrsOf(findChildren(tcPr ?? {}, 'w:tcW').at(-1) ?? {})
      // pct cell widths (1/50 of a percent) of a pct table resolve against its width
      // (probe 2026-10-29: tcW pct 15/40/30/15 drawn as those shares, grid ignored)
      const w =
        !a['w:type'] || a['w:type'] === 'dxa'
          ? Number(a['w:w']) || 0
          : a['w:type'] === 'pct' && pctTableTwips > 0
            ? Math.round(((Number(a['w:w']) || 0) / 5000) * pctTableTwips)
            : 0
      if (span === 1 && w > 0) cols[idx] = Math.max(cols[idx] || 0, w)
      idx += span
    }
    colCount = Math.max(colCount, idx + edges.after)
  }
  if (colCount === 0) return undefined
  for (let i = 0; i < colCount; i++) if (!(cols[i] > 0)) return undefined
  return cols.slice(0, colCount)
}

/**
 * Column widths w:noWrap cells pin: a noWrap cell holds its column at its w:tcW
 * (probe 2026-11-08: 1530-twip noWrap cell drawn 1530 wide beside a 12915-twip
 * neighbour; both noWrap and the sum over the column scales 1530 -> 956 and the
 * cell wraps). A spanning cell raises its columns to its tcW when they pin less.
 */
function noWrapPinnedTwips(
  rows: TableCell[][],
  rowTcws: (number | undefined)[][],
  grid: number[] | undefined,
): number {
  const pinned: number[] = []
  const spans: { from: number; to: number; w: number }[] = []
  rows.forEach((cells, i) => {
    let col = 0
    cells.forEach((cell, j) => {
      const span = cell.colSpan ?? 1
      const w = rowTcws[i]?.[j]
      if (cell.noWrap && w !== undefined && w > 0) {
        if (span === 1) pinned[col] = Math.max(pinned[col] ?? 0, w)
        else spans.push({ from: col, to: col + span, w })
      }
      col += span
    })
  })
  for (const { from, to, w } of spans) {
    let have = 0
    for (let c = from; c < to; c++) have += pinned[c] ?? 0
    if (have >= w) continue
    const share = (c: number) => grid?.[c] ?? 1
    let total = 0
    for (let c = from; c < to; c++) total += share(c)
    for (let c = from; c < to; c++)
      pinned[c] = Math.max(pinned[c] ?? 0, Math.round((w * share(c)) / (total || 1)))
  }
  return pinned.reduce((a, b) => a + (b ?? 0), 0)
}

/** true when no direct cell of the table declares w:tcW (generator output) */
function tableLacksTcW(tbl: XNode): boolean {
  for (const tr of childrenThroughSdt(tbl, 'w:tr'))
    for (const tc of childrenThroughSdt(tr, 'w:tc'))
      if (findChild(findChild(tc, 'w:tcPr') ?? {}, 'w:tcW')) return false
  return true
}

/** trPr w:gridBefore/w:gridAfter column counts plus their w:wBefore/w:wAfter widths (twips) */
export function rowGridEdges(tr: XNode): {
  before: number
  after: number
  wBefore?: number
  wAfter?: number
} {
  const trPr = findChild(tr, 'w:trPr')
  if (!trPr) return { before: 0, after: 0 }
  const count = (name: string) => {
    const v = Number(attrsOf(findChild(trPr, name) ?? {})['w:val'])
    return Number.isFinite(v) && v > 0 ? Math.floor(v) : 0
  }
  const width = (name: string) => {
    const a = attrsOf(findChild(trPr, name) ?? {})
    const w = !a['w:type'] || a['w:type'] === 'dxa' ? Number(a['w:w']) : NaN
    return w > 0 ? w : undefined
  }
  return {
    before: count('w:gridBefore'),
    after: count('w:gridAfter'),
    wBefore: width('w:wBefore'),
    wAfter: width('w:wAfter'),
  }
}

/** boundary positions within TOL twips fuse into one grid line (rounding drift across generator-written rows) */
const GRID_SNAP_TOL = 20

/**
 * Repair grids whose rows do not all span the declared column count (legacy
 * generators emit gridSpan/tblGrid pairs that disagree with the cells' w:tcW).
 * Word lays each such row out from the cells' preferred widths, so the true
 * grid is the union of every row's tcW boundaries; without this, a cell mapped
 * onto a leftover sliver column collapses to one character per line.
 * Mutates the cells' colSpan and returns the rebuilt column widths, or
 * undefined when the grid is already consistent or a width is unresolvable.
 */
export function reconcileGridColumns(
  rows: TableCell[][],
  rowTcws: Array<Array<number | undefined>>,
  gridCols: number[] | undefined,
): number[] | undefined {
  const spanSums = rows.map((row) => row.reduce((sum, c) => sum + (c.colSpan ?? 1), 0))
  let widestSpan = 0
  for (const sum of spanSums) {
    if (sum > widestSpan) widestSpan = sum
  }
  const colCount = gridCols?.length ?? widestSpan
  if (spanSums.every((sum) => sum === colCount)) return undefined
  const rowBounds: number[][] = []
  for (let r = 0; r < rows.length; r++) {
    const bounds: number[] = []
    let x = 0
    let pos = 0
    for (let c = 0; c < rows[r].length; c++) {
      const span = rows[r][c].colSpan ?? 1
      let w = rowTcws[r][c]
      if (!(w !== undefined && w > 0)) {
        w = gridCols?.slice(pos, pos + span).reduce((sum, v) => sum + v, 0)
      }
      if (!(w !== undefined && w > 0)) return undefined
      x += w
      bounds.push(x)
      pos += span
    }
    rowBounds.push(bounds)
  }
  const sorted = rowBounds.flat().sort((a, b) => a - b)
  const reps: number[] = []
  for (const b of sorted) {
    if (reps.length === 0 || b - reps[reps.length - 1] > GRID_SNAP_TOL) reps.push(b)
  }
  // Word caps table grids at 63 columns; a wider union means garbage input
  if (reps.length === 0 || reps.length > 96) return undefined
  const repIndex = (b: number) => {
    let lo = 0
    let hi = reps.length - 1
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1
      if (reps[mid] <= b) lo = mid
      else hi = mid - 1
    }
    return reps[lo] <= b && b - reps[lo] <= GRID_SNAP_TOL ? lo : -1
  }
  const newSpans: number[][] = []
  for (const bounds of rowBounds) {
    const spans: number[] = []
    let prev = -1
    for (const b of bounds) {
      const idx = repIndex(b)
      if (idx <= prev) return undefined
      spans.push(idx - prev)
      prev = idx
    }
    newSpans.push(spans)
  }
  rows.forEach((row, r) =>
    row.forEach((cell, c) => {
      if (newSpans[r][c] > 1) cell.colSpan = newSpans[r][c]
      else delete cell.colSpan
    }),
  )
  return reps.map((v, i) => v - (i > 0 ? reps[i - 1] : 0))
}

/**
 * Widths Word shrinks an over-wide autofit table to: the text column, one
 * newspaper column in a multi-column section, each less the signed table indent
 * (a negative indent widens the box to the left, the right edge stays on the
 * margin); compat < 15 measures the indent to the cell text, so the border box hangs by
 * the side cell margins (probe 2026-09-17: compat 14 tblInd 558 in an 8550
 * column -> 8208 = 8550 - 558 + 216; compat 15 -> 7992).
 */
function usableTableWidths(sect: SectionSettings, tblIndTwips: number, hangTwips = 0): number[] {
  const text = sect.pageWidth - sect.marginLeft - sect.marginRight
  const perColumn =
    sect.columns > 1
      ? sect.colWidths?.length === sect.columns
        ? sect.colWidths
        : [(text - (sect.colSpace ?? 720) * (sect.columns - 1)) / sect.columns]
      : []
  const indent = Number.isFinite(tblIndTwips) ? tblIndTwips : 0
  const bases = [text, ...perColumn]
  const spans = [...bases, ...bases.map((w) => w - indent)]
  if (hangTwips > 0) spans.push(...spans.map((w) => w + hangTwips))
  return [...new Set(spans)].filter((w) => w > 0)
}

/** side cell margin of a table that declares none (twips) */
export const DEFAULT_TABLE_CELL_MAR = 108

/** every gridCol within UNIFORM_GRID_TOLERANCE of the mean: a generator placeholder, not a laid-out grid */
function isUniformGrid(widths: number[] | undefined): boolean {
  if (!widths) return true
  const total = widths.reduce((a, b) => a + b, 0)
  return widths.every(
    (w) =>
      Math.abs(w * widths.length - total) <=
      Math.max(UNIFORM_GRID_MIN_TWIPS, total * UNIFORM_GRID_TOLERANCE),
  )
}

/** what the parser hands over: the section the table sits in and the table-level
 *  properties the width rules read besides the widths themselves */
export interface TableWidthContext {
  sect: SectionSettings | undefined
  tblIndTwips: number
  /** effective table cell margins (document level, else the table style) */
  cellMar: CellMargins | undefined
  /** compatibilityMode < 15: the indent is measured to the cell text */
  legacy: boolean
  /** w:tblCellSpacing (twips) of the table or its first row */
  tblSpacing: number | undefined
}

/** the column grid a w:tbl declares, after the gridCol-versus-tcW arbitration,
 *  and the facts the later width flags are decided from */
export interface DeclaredTableWidths {
  colWidthsPct: number[] | undefined
  colWidthsTwips: number[] | undefined
  gridWidthsRaw: number[] | undefined
  widthPct: number | undefined
  fixedLayout: boolean
  uniformGrid: boolean
  gridMatchesPct: boolean
  tcwWins: boolean
  hasTblW: boolean
  tblWType: unknown
  tblWValue: number
}

export function declaredTableWidths(
  tbl: XNode,
  tblPrNode: XNode | undefined,
  { sect, tblIndTwips, cellMar, legacy, tblSpacing }: TableWidthContext,
): DeclaredTableWidths {
  const grid = findChild(tbl, 'w:tblGrid')
  let colWidthsPct: number[] | undefined
  let colWidthsTwips: number[] | undefined
  let gridWidthsRaw: number[] | undefined
  if (grid) {
    const widths = findChildren(grid, 'w:gridCol').map((c) => Number(attrsOf(c)['w:w']) || 0)
    const total = widths.reduce((a, b) => a + b, 0)
    if (total > 0) {
      gridWidthsRaw = widths
      colWidthsPct = widths.map((w) => (w / total) * 100)
      if (widths.every((w) => w > 0)) colWidthsTwips = widths
    }
  }
  const fixedLayout = attrsOf(findChild(tblPrNode ?? {}, 'w:tblLayout') ?? {})['w:type'] === 'fixed'
  const tblWNode = findChild(tblPrNode ?? {}, 'w:tblW')
  const tblW = attrsOf(tblWNode ?? {})
  const legacyHang = legacy
    ? (cellMar?.left ?? DEFAULT_TABLE_CELL_MAR) + (cellMar?.right ?? DEFAULT_TABLE_CELL_MAR)
    : 0
  const uniformGrid = isUniformGrid(colWidthsTwips)
  let widthPct: number | undefined
  if (tblW['w:type'] === 'pct') {
    const raw = String(tblW['w:w'] ?? '')
    // The pct unit is 1/50 of a percentage point; some generators write a literal "NN%"
    const pct = raw.endsWith('%') ? parseFloat(raw) : Number(raw) / 50
    if (Number.isFinite(pct) && pct > 0 && pct <= 100) widthPct = pct
  }
  const gridTotalRaw = (colWidthsTwips ?? []).reduce((a, b) => a + b, 0)
  // a pct table's saved grid is Word's own layout only when it sums to the pct
  // width (SAVED_GRID_TOLERANCE); any other non-uniform grid is a generator artefact and yields
  // to tcW (probe 2026-10-29)
  const pctWidths =
    widthPct && sect ? usableTableWidths(sect, tblIndTwips).map((u) => (u * widthPct!) / 100) : []
  const gridMatchesPct =
    !!widthPct &&
    gridTotalRaw > 0 &&
    !tableLacksTcW(tbl) &&
    (pctWidths.length === 0 ||
      pctWidths.some((w) => Math.abs(gridTotalRaw - w) <= w * SAVED_GRID_TOLERANCE))
  // Cell-level w:tcW is Word's actual layout input for auto tables; generators often
  // leave a stale evenly-split tblGrid behind. When the two disagree, tcW wins.
  const tcwWidths = tcwColumnWidths(tbl, widthPct ? (pctWidths[0] ?? gridTotalRaw) : 0)
  let tcwWins = false
  if (tcwWidths) {
    const tcwTotal = tcwWidths.reduce((a, b) => a + b, 0)
    const tcwPct = tcwWidths.map((w) => (w / tcwTotal) * 100)
    // Fixed layout draws the saved grid when its total agrees with tcW (Word
    // keeps gridCol over disagreeing cell widths); a matching ratio with a
    // different absolute sum still means the grid is stale. So does a table with
    // no declared width (tblW auto): Word sizes it from the cells' preferred widths,
    // unless those overflow the text column - then Word shrinks the table to the
    // column and the saved grid already is that shrunk layout, so a grid spanning
    // the column stays (a narrower same-ratio grid is still a stale placeholder).
    const gridTotal = (colWidthsTwips ?? []).reduce((a, b) => a + b, 0)
    // an autofit grid with unequal columns is Word's finished layout (Word
    // draws it on open even when tcW disagree); only an evenly split
    // placeholder grid yields to tcW
    const tblWAuto =
      tblW['w:type'] === 'auto' ||
      ((!tblW['w:type'] || tblW['w:type'] === 'dxa') && !(Number(tblW['w:w']) > 0))
    const gridIsShrunkLayout = (sect ? usableTableWidths(sect, tblIndTwips, legacyHang) : []).some(
      (usable) =>
        tcwTotal > usable + tcwWidths.length * GRID_ROUNDING_TWIPS_PER_COLUMN &&
        Math.abs(gridTotal - usable) <= usable * SAVED_GRID_TOLERANCE,
    )
    // with w:tblCellSpacing the grid also holds the gaps that w:tcW leaves out;
    // Word draws that wider grid, so it is not a stale one
    const gridExcess = gridTotal - tcwTotal
    const gridRounding = tcwWidths.length * GRID_ROUNDING_TWIPS_PER_COLUMN
    const spacingRoom = tblSpacing
      ? CELL_SPACING_GAPS_PER_COLUMN * tblSpacing * tcwWidths.length
      : 0
    const disagree =
      !colWidthsPct ||
      colWidthsPct.length !== tcwPct.length ||
      (uniformGrid &&
        colWidthsPct.some((w, i) => Math.abs(w - tcwPct[i]) > TCW_GRID_PCT_TOLERANCE)) ||
      ((fixedLayout || (tblWAuto && !gridIsShrunkLayout)) &&
        (gridExcess < -gridRounding || gridExcess > gridRounding + spacingRoom)) ||
      (!!widthPct && !gridMatchesPct)
    if (disagree) {
      colWidthsPct = tcwPct
      colWidthsTwips = tcwWidths
      tcwWins = true
    }
  }
  // Placeholder grids (generators emit 100-twip gridCols without computing a
  // layout) collapse the table to a strip. Word lays autofit tables out from
  // tblW/tcW, so a grid summing below the declared dxa width stretches to it.
  if (!fixedLayout && colWidthsTwips) {
    const tblWDxa = !tblW['w:type'] || tblW['w:type'] === 'dxa' ? Number(tblW['w:w']) : NaN
    const gridTotal = colWidthsTwips.reduce((a, b) => a + b, 0)
    const rounding = colWidthsTwips.length * GRID_ROUNDING_TWIPS_PER_COLUMN
    if (tblWDxa > 0 && gridTotal > 0 && gridTotal < tblWDxa - rounding) {
      const scale = tblWDxa / gridTotal
      colWidthsTwips = colWidthsTwips.map((w) => Math.round(w * scale))
    }
  }
  return {
    colWidthsPct,
    colWidthsTwips,
    gridWidthsRaw,
    widthPct,
    fixedLayout,
    uniformGrid,
    gridMatchesPct,
    tcwWins,
    hasTblW: !!tblWNode,
    tblWType: tblW['w:type'],
    tblWValue: Number(tblW['w:w']),
  }
}

type TableWidthFields = Pick<
  TableModel,
  | 'colWidthsPct'
  | 'colWidthsTwips'
  | 'noWrapPinnedTwips'
  | 'widthPct'
  | 'autoLayout'
  | 'layoutGrid'
  | 'contentAutofit'
  | 'autoFit'
  | 'fixedLayout'
  | 'dxaWidth'
>

/** the width fields of the table model, in the order the model has always carried
 *  them; reconciles the cells' colSpan against the rows' own widths first */
export function tableWidthFields(
  tbl: XNode,
  declared: DeclaredTableWidths,
  rows: Parameters<typeof reconcileGridColumns>[0],
  rowTcws: Parameters<typeof reconcileGridColumns>[1],
): TableWidthFields {
  let { colWidthsPct, colWidthsTwips } = declared
  const { widthPct, fixedLayout, gridMatchesPct, tcwWins, uniformGrid, tblWType, tblWValue } =
    declared
  const reconciled = reconcileGridColumns(rows, rowTcws, declared.gridWidthsRaw)
  if (reconciled) {
    colWidthsTwips = reconciled
    const total = reconciled.reduce((a, b) => a + b, 0)
    colWidthsPct = reconciled.map((w) => (w / total) * 100)
  }
  const model: TableWidthFields = { colWidthsPct }
  if (colWidthsTwips) model.colWidthsTwips = colWidthsTwips
  const pinned = noWrapPinnedTwips(rows, rowTcws, colWidthsTwips)
  if (pinned > 0) model.noWrapPinnedTwips = pinned
  if (widthPct) model.widthPct = widthPct
  const autoWidth =
    !declared.hasTblW ||
    tblWType === 'auto' ||
    ((!tblWType || tblWType === 'dxa') && !(tblWValue > 0))
  // pct widths still lay out with the autofit algorithm (only w:tblLayout
  // fixed switches it off), so their columns keep the min-content floor
  if (!fixedLayout && (autoWidth || widthPct)) model.autoLayout = true
  // Word lays an autofit table (tblW auto or dxa) out from tcW/content and saves
  // the result as the grid, so an unequal grid it kept is Word's own measure of
  // every word (probe: a dxa table drawn at its saved grid, not its tcW)
  if (!fixedLayout && (!widthPct || gridMatchesPct) && colWidthsTwips && !tcwWins && !uniformGrid)
    model.layoutGrid = true
  // no w:tcW at all: Word lays the pct table out from the cell content and never
  // consults the generator's grid (probe 2026-10-20: uniform and skewed grids
  // give identical columns)
  if (!fixedLayout && widthPct && tableLacksTcW(tbl)) model.contentAutofit = true
  model.autoFit =
    fixedLayout || (!autoWidth && !widthPct) ? 'fixed' : widthPct === 100 ? 'window' : 'contents'
  if (fixedLayout) model.fixedLayout = true
  else if ((!tblWType || tblWType === 'dxa') && tblWValue > 0 && !widthPct) model.dxaWidth = true
  return model
}

/** a cell's own preferred width (the last w:tcW wins), dxa only */
export function cellTcwTwips(tc: XNode): number | undefined {
  const a = attrsOf(findChildren(findChild(tc, 'w:tcPr') ?? {}, 'w:tcW').at(-1) ?? {})
  const rawW = !a['w:type'] || a['w:type'] === 'dxa' ? Number(a['w:w']) : NaN
  return rawW > 0 ? rawW : undefined
}

/** header/footer table: its saved grid columns, the table width they sum to (a dxa
 *  w:tblW when there is no grid), and whether the table is sized as a percentage */
export function hfTableGrid(
  tbl: XNode,
  tblPr: XNode | undefined,
): { gridCols: number[]; tblWidth: number; pctWidth: boolean } {
  const tblW = attrsOf(findChild(tblPr ?? {}, 'w:tblW') ?? {})
  const grid = findChild(tbl, 'w:tblGrid')
  const gridCols = grid
    ? findChildren(grid, 'w:gridCol').map((g) => Number(attrsOf(g)['w:w']) || 0)
    : []
  const gridWidth = gridCols.reduce((sum, w) => sum + w, 0)
  const tblWidth =
    gridWidth || (tblW['w:type'] === 'dxa' || !tblW['w:type'] ? Number(tblW['w:w']) || 0 : 0)
  return { gridCols, tblWidth, pctWidth: tblW['w:type'] === 'pct' }
}

/** header/footer cell: its declared dxa w:tcW, else the grid columns it spans */
export function hfCellWidth(
  tcPr: XNode | undefined,
  gridCols: number[],
  col: number,
  span: number,
): { widthTwips: number; widthIsPct: boolean } {
  const a = attrsOf(findChild(tcPr ?? {}, 'w:tcW') ?? {})
  const v = Number(a['w:w'])
  const widthIsPct = a['w:type'] === 'pct'
  const declared = !widthIsPct && Number.isFinite(v) && v > 0 ? v : 0
  const fromGrid = gridCols.slice(col, col + span).reduce((s, w) => s + w, 0)
  return { widthTwips: declared || fromGrid, widthIsPct }
}
