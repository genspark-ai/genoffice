/**
 * Lattice table detection (pdf2docx rule 2, vector domain): strokes grouped by
 * contact → per group an implied outer frame is added, horizontal/vertical
 * lines are clustered into row/column boundaries, and MISSING borders between
 * grid cells become merged cells (gridSpan / vMerge). Fills color cells; page
 * chars falling inside the region are routed through the regular char pipeline
 * per cell and removed from the paragraph flow.
 *
 * Suppression bias is "miss rather than misfire": grids under 2×2, groups
 * without both line directions, tiny regions and low border coverage are all
 * rejected.
 */
import type { Rect } from '../geometry'
import { overlapRatio, rectArea, rectCenterX, rectCenterY, rectUnionAll } from '../geometry'
import type { Fill, PageShapes, PdfChar, Stroke, TableBlock, TableCellBlock } from '../ir'
import { groupIntoBlocks } from './blocks'
import { analyzeChars } from './chars'

/** strokes whose boxes come this close (pt) are touching → same group */
const CONNECT_TOL = 2.5
/** collinear lines within this distance (pt) cluster into one boundary */
const POS_TOL = 2.0
/** boundaries closer than this (pt) merge — cells thinner than this are noise */
const MIN_CELL_DIM = 3
/** a border "exists" over a cell edge when its strokes cover this share of it */
const BORDER_COVER_MIN = 0.5
/** minimum table region size (pt) — smaller stroke groups are decorations */
const MIN_TABLE_W = 24
const MIN_TABLE_H = 16
/** minimum share of all grid boundary length covered by real strokes */
const GRID_COVER_MIN = 0.25
/** a closed text box's four outer edges must each be near-fully drawn */
const BOX_EDGE_COVER_MIN = 0.85
/** a low-rank box must hold at least this many visible chars (else it is
 * extractEmptyFrames territory) */
const BOX_MIN_CHARS = 4
/** …and must not span most of the page height (page frames, certificates) */
const BOX_MAX_PAGE_SHARE = 0.55
/** a short strip (a title banner) renders better as flow text than a table */
const BOX_MIN_H_PT = 38
/** fills larger than this multiple of the table area are page background */
const FILL_MAX_AREA_RATIO = 1.5
/** one stroke color must cover this share of grid line length to become the
 * table's border color (mixed-color grids keep the default) */
const BORDER_COLOR_DOMINANCE = 0.8

const touches = (a: Rect, b: Rect, tol: number): boolean =>
  a.x0 <= b.x1 + tol && b.x0 <= a.x1 + tol && a.y0 <= b.y1 + tol && b.y0 <= a.y1 + tol

/** Union-find grouping of strokes by contact (within CONNECT_TOL). */
export function groupStrokes(strokes: readonly Stroke[]): Stroke[][] {
  const parent = strokes.map((_, i) => i)
  const find = (i: number): number => {
    while (parent[i] !== i) {
      parent[i] = parent[parent[i]!]!
      i = parent[i]!
    }
    return i
  }
  const union = (a: number, b: number): void => {
    const ra = find(a)
    const rb = find(b)
    if (ra !== rb) parent[rb] = ra
  }
  for (let i = 0; i < strokes.length; i++) {
    for (let j = i + 1; j < strokes.length; j++) {
      if (touches(strokes[i]!.box, strokes[j]!.box, CONNECT_TOL)) union(i, j)
    }
  }
  const byRoot = new Map<number, Stroke[]>()
  for (let i = 0; i < strokes.length; i++) {
    const root = find(i)
    let list = byRoot.get(root)
    if (!list) byRoot.set(root, (list = []))
    list.push(strokes[i]!)
  }
  return [...byRoot.values()]
}

/** one grid boundary: clustered collinear line segments at `pos` */
interface BoundaryLine {
  pos: number
  /** covered intervals along the perpendicular axis */
  segments: Array<[number, number]>
  virtual: boolean
}

export interface TableGrid {
  box: Rect
  /** column boundaries, left → right */
  xs: number[]
  /** row boundaries, top → bottom (descending y) */
  ys: number[]
  /** parallel to ys / xs */
  hLines: BoundaryLine[]
  vLines: BoundaryLine[]
  /** under-2×2 grid accepted as a closed text box (P18 C) — needs vetting */
  lowRank?: boolean
}

/** cluster parallel strokes into boundary lines (positions within tol merge) */
function clusterLines(
  strokes: Stroke[],
  pos: (s: Stroke) => number,
  span: (s: Stroke) => [number, number],
  tol: number,
): BoundaryLine[] {
  const sorted = [...strokes].sort((a, b) => pos(a) - pos(b))
  const clusters: BoundaryLine[] = []
  for (const s of sorted) {
    const p = pos(s)
    const last = clusters[clusters.length - 1]
    if (last && p - last.pos <= tol) {
      last.segments.push(span(s))
      // drift toward the running mean so long tolerance chains stay anchored
      last.pos = (last.pos + p) / 2
    } else {
      clusters.push({ pos: p, segments: [span(s)], virtual: false })
    }
  }
  return clusters
}

/** total length of [lo,hi] covered by the line's (merged) segments */
function coverage(line: BoundaryLine, lo: number, hi: number): number {
  if (hi <= lo) return 0
  const clipped = line.segments
    .map(([a, b]): [number, number] => [Math.max(Math.min(a, b), lo), Math.min(Math.max(a, b), hi)])
    .filter(([a, b]) => b > a)
    .sort((a, b) => a[0] - b[0])
  let total = 0
  let cursor = lo
  for (const [a, b] of clipped) {
    const start = Math.max(a, cursor)
    if (b > start) {
      total += b - start
      cursor = b
    }
  }
  return total
}

const borderPresent = (line: BoundaryLine, lo: number, hi: number): boolean =>
  line.virtual || coverage(line, lo, hi) >= BORDER_COVER_MIN * (hi - lo)

/**
 * Solve one stroke group into a rectangular grid. Returns null when the group
 * fails any suppression rule (not a table).
 */
export function solveGrid(group: readonly Stroke[], allowLowRank = false): TableGrid | null {
  const h = group.filter((s) => s.orientation === 'h')
  const v = group.filter((s) => s.orientation === 'v')
  // underlines / separators have one direction only — never a table
  if (h.length === 0 || v.length === 0) return null

  const box = rectUnionAll(group.map((s) => s.box))
  if (box.x1 - box.x0 < MIN_TABLE_W || box.y1 - box.y0 < MIN_TABLE_H) return null

  let hLines = clusterLines(
    h,
    (s) => rectCenterY(s.box),
    (s) => [s.box.x0, s.box.x1],
    POS_TOL,
  )
  let vLines = clusterLines(
    v,
    (s) => rectCenterX(s.box),
    (s) => [s.box.y0, s.box.y1],
    POS_TOL,
  )

  // implied outer frame: missing edge lines are added as virtual boundaries
  const frame = (
    lines: BoundaryLine[],
    lo: number,
    hi: number,
    seg: [number, number],
  ): BoundaryLine[] => {
    const out = [...lines]
    if (!out.some((l) => Math.abs(l.pos - lo) <= MIN_CELL_DIM)) {
      out.unshift({ pos: lo, segments: [seg], virtual: true })
    }
    if (!out.some((l) => Math.abs(l.pos - hi) <= MIN_CELL_DIM)) {
      out.push({ pos: hi, segments: [seg], virtual: true })
    }
    return out.sort((a, b) => a.pos - b.pos)
  }
  hLines = frame(hLines, box.y0, box.y1, [box.x0, box.x1])
  vLines = frame(vLines, box.x0, box.x1, [box.y0, box.y1])

  // boundaries closer than a plausible cell dimension merge (noise / double strokes)
  const dedupe = (lines: BoundaryLine[]): BoundaryLine[] => {
    const out: BoundaryLine[] = []
    for (const line of lines) {
      const last = out[out.length - 1]
      if (last && line.pos - last.pos < MIN_CELL_DIM) {
        last.segments.push(...line.segments)
        last.virtual = last.virtual && line.virtual
      } else {
        out.push({ ...line, segments: [...line.segments] })
      }
    }
    return out
  }
  hLines = dedupe(hLines)
  vLines = dedupe(vLines)

  const rows = hLines.length - 1
  const cols = vLines.length - 1
  let lowRank = false
  if (rows < 2 || cols < 2) {
    // an under-2×2 grid is normally a decoration — EXCEPT a fully drawn
    // closed box (a prompt/answer frame on a form, a stack of boxed form
    // sections): all four outer edges real and near-complete (P18 C). The
    // caller opts in and vets the result against page text and other grids.
    if (!allowLowRank || rows < 1 || cols < 1) return null
    const closed =
      [hLines[0]!, hLines[hLines.length - 1]!].every(
        (l) => !l.virtual && coverage(l, box.x0, box.x1) >= BOX_EDGE_COVER_MIN * (box.x1 - box.x0),
      ) &&
      [vLines[0]!, vLines[vLines.length - 1]!].every(
        (l) => !l.virtual && coverage(l, box.y0, box.y1) >= BOX_EDGE_COVER_MIN * (box.y1 - box.y0),
      )
    if (!closed) return null
    lowRank = true
  }

  // real-stroke coverage across all boundaries (virtual frame counts as gap)
  const width = box.x1 - box.x0
  const height = box.y1 - box.y0
  let covered = 0
  for (const line of hLines) if (!line.virtual) covered += coverage(line, box.x0, box.x1)
  for (const line of vLines) if (!line.virtual) covered += coverage(line, box.y0, box.y1)
  const totalLen = hLines.length * width + vLines.length * height
  if (covered < GRID_COVER_MIN * totalLen) return null

  return {
    box,
    xs: vLines.map((l) => l.pos),
    ys: hLines.map((l) => l.pos).reverse(), // top → bottom
    hLines: [...hLines].reverse(),
    vLines,
    ...(lowRank ? { lowRank: true } : {}),
  }
}

/** anchor coordinates of the merged cell covering each grid position */
export interface CellLayout {
  rows: number
  cols: number
  /** [r][c] → anchor {r,c} of the covering cell */
  anchorOf: Array<Array<{ r: number; c: number }>>
  /** spans keyed by `r:c` anchor */
  spanOf: Map<string, { colSpan: number; rowSpan: number }>
}

/**
 * Resolve merged cells: a missing border between neighbouring grid positions
 * merges them (missing right border → gridSpan grows; missing bottom border →
 * vMerge grows, provided the lower row repeats the same internal structure).
 */
export function layoutCells(grid: TableGrid): CellLayout {
  const rows = grid.ys.length - 1
  const cols = grid.xs.length - 1
  // row r band: y in [ys[r+1], ys[r]] (ys descend)
  const rowBand = (r: number): [number, number] => [grid.ys[r + 1]!, grid.ys[r]!]
  const vPresent = (boundary: number, r: number): boolean => {
    const [lo, hi] = rowBand(r)
    return borderPresent(grid.vLines[boundary]!, lo, hi)
  }
  const hPresent = (boundary: number, x0: number, x1: number): boolean =>
    borderPresent(grid.hLines[boundary]!, x0, x1)

  const anchorOf: CellLayout['anchorOf'] = Array.from({ length: rows }, () =>
    Array.from({ length: cols }, () => ({ r: -1, c: -1 })),
  )
  const spanOf = new Map<string, { colSpan: number; rowSpan: number }>()
  const taken = Array.from({ length: rows }, () => Array.from({ length: cols }, () => false))

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (taken[r]![c]!) continue
      let colSpan = 1
      while (c + colSpan < cols && !taken[r]![c + colSpan]! && !vPresent(c + colSpan, r)) colSpan++
      let rowSpan = 1
      expand: while (r + rowSpan < rows) {
        const below = r + rowSpan
        // the shared horizontal border must be absent across the whole cell width
        if (hPresent(below, grid.xs[c]!, grid.xs[c + colSpan]!)) break
        // and the lower row may not re-introduce internal vertical borders
        for (let k = 1; k < colSpan; k++) if (vPresent(c + k, below)) break expand
        for (let k = 0; k < colSpan; k++) if (taken[below]![c + k]!) break expand
        rowSpan++
      }
      spanOf.set(`${r}:${c}`, { colSpan, rowSpan })
      for (let dr = 0; dr < rowSpan; dr++) {
        for (let dc = 0; dc < colSpan; dc++) {
          taken[r + dr]![c + dc] = true
          anchorOf[r + dr]![c + dc] = { r, c }
        }
      }
    }
  }
  return { rows, cols, anchorOf, spanOf }
}

// ── ghost edge columns (P13 B) ──
/** an internal row boundary "reaches into" an edge column when it covers this share of it */
const GHOST_H_REACH_RATIO = 0.3
/** …or at least this many points (caps/overshoot of neighbouring segments stay noise) */
const GHOST_H_REACH_MIN_PT = 2

/**
 * Drop ghost edge columns (P13 B): HTML-sourced PDFs (MaxGen et al.) draw a
 * container frame around the table that is wider than the grid itself, and
 * its full-height edge becomes an extra column boundary — the converted table
 * gains an empty trailing (or leading) column the source never had. Row
 * border overshoot creates the same artifact via the implied (virtual) frame.
 *
 * An edge column is a ghost only when nothing at all acknowledges it: no
 * internal row boundary reaches into it, no text sits in it, no fill covers
 * it, AND the geometry says the grid ends before it — either the outer
 * boundary was never drawn (virtual), or the inner boundary carries per-row
 * border segments (cell-level drawing provably stops there). Legit empty
 * columns (form answer boxes, cross-page vMerge continuations) are bordered
 * by single full-height lines and keep their column; tables under 3 rows
 * have no internal row boundaries to consult and are never trimmed (miss
 * rather than misfire).
 */
export function trimGhostEdgeColumns(
  grid: TableGrid,
  chars: readonly PdfChar[],
  fills: readonly Fill[],
): TableGrid {
  const tableArea = rectArea(grid.box)
  // top/bottom frame lines span the container too — only internal ones
  // discriminate, and a single one is too weak a signal (a 2-row table's
  // merged empty trailing cell would read as a ghost): require ≥ 3 rows
  const internalH = grid.hLines.slice(1, -1)
  if (internalH.length < 2) return grid

  // per-row border drawing: segments breaking at internal row boundaries
  const rowStructured = (line: BoundaryLine): boolean =>
    line.segments.length >= 2 &&
    line.segments.some(([a, b]) =>
      internalH.some(
        (h) =>
          Math.abs(Math.min(a, b) - h.pos) <= POS_TOL ||
          Math.abs(Math.max(a, b) - h.pos) <= POS_TOL,
      ),
    )

  const isGhost = (inner: BoundaryLine, outer: BoundaryLine, x0: number, x1: number): boolean => {
    const bandW = x1 - x0
    if (bandW <= 0) return false
    if (!outer.virtual && !rowStructured(inner)) return false
    const reach = Math.max(GHOST_H_REACH_MIN_PT, GHOST_H_REACH_RATIO * bandW)
    for (const line of internalH) {
      if (line.virtual || coverage(line, x0, x1) >= reach) return false
    }
    for (const c of chars) {
      if (c.code <= 0x20) continue
      const cx = rectCenterX(c.box)
      const cy = rectCenterY(c.box)
      if (cx >= x0 && cx <= x1 && cy >= grid.box.y0 && cy <= grid.box.y1) return false
    }
    for (const fill of fills) {
      if (WHITE.test(fill.color)) continue
      if (rectArea(fill.box) > FILL_MAX_AREA_RATIO * tableArea) continue
      const ix = Math.min(fill.box.x1, x1) - Math.max(fill.box.x0, x0)
      const iy = Math.min(fill.box.y1, grid.box.y1) - Math.max(fill.box.y0, grid.box.y0)
      if (ix >= 0.5 * bandW && iy >= MIN_CELL_DIM) return false
    }
    return true
  }

  const xs = [...grid.xs]
  const vLines = [...grid.vLines]
  while (
    xs.length > 3 &&
    isGhost(
      vLines[vLines.length - 2]!,
      vLines[vLines.length - 1]!,
      xs[xs.length - 2]!,
      xs[xs.length - 1]!,
    )
  ) {
    xs.pop()
    vLines.pop()
  }
  while (xs.length > 3 && isGhost(vLines[1]!, vLines[0]!, xs[0]!, xs[1]!)) {
    xs.shift()
    vLines.shift()
  }
  if (xs.length === grid.xs.length) return grid
  return {
    ...grid,
    xs,
    vLines,
    box: { ...grid.box, x0: xs[0]!, x1: xs[xs.length - 1]! },
  }
}

// ── cell vertical alignment (P13 A) ──
/** the cell must be this much taller than its text for alignment to mean anything */
const VALIGN_MIN_SLACK = 1.25
/** centre/bottom tolerance: max(this, 8% of the cell height) in points */
const VALIGN_TOL_PT = 1.5
const VALIGN_TOL_RATIO = 0.08

/**
 * Detect vertical content alignment inside one cell (P13 A): compare the text
 * bbox's vertical centre against the cell's geometric centre. Only cells with
 * real head-room (height > text height × 1.25) carry alignment semantics —
 * tightly packed cells render the same regardless and stay unmarked (top).
 */
export function detectCellVAlign(
  cellBox: Rect,
  chars: readonly PdfChar[],
): 'center' | 'bottom' | undefined {
  let y0 = Infinity
  let y1 = -Infinity
  for (const c of chars) {
    if (c.code <= 0x20) continue
    y0 = Math.min(y0, c.box.y0)
    y1 = Math.max(y1, c.box.y1)
  }
  if (y1 <= y0) return undefined
  const cellH = cellBox.y1 - cellBox.y0
  if (cellH <= (y1 - y0) * VALIGN_MIN_SLACK) return undefined
  const tol = Math.max(VALIGN_TOL_PT, VALIGN_TOL_RATIO * cellH)
  if (Math.abs((y0 + y1) / 2 - (cellBox.y0 + cellBox.y1) / 2) < tol) return 'center'
  if (y0 - cellBox.y0 < tol) return 'bottom'
  return undefined
}

// ── cell horizontal alignment ──
/** the cell must be this much wider than its widest line for alignment to mean anything */
const HALIGN_MIN_SLACK = 1.15
/** centre/right tolerance: max(this, 4% of the cell width) in points */
const HALIGN_TOL_PT = 1.5
const HALIGN_TOL_RATIO = 0.04

/**
 * Detect horizontal content alignment inside one cell: every line must agree
 * (centres on the cell's centre, or right edges on the cell's right), and the
 * cell needs real slack — a line that fills the cell renders the same under
 * any w:jc and stays unmarked. The block analyser cannot make this call: its
 * reference frame is the text's own extent, which for a lone short line
 * carries no alignment information; only the cell box does.
 */
export function detectCellHAlign(
  cellBox: Rect,
  lines: readonly { box: Rect }[],
): 'center' | 'right' | undefined {
  if (lines.length === 0) return undefined
  const cellW = cellBox.x1 - cellBox.x0
  const tol = Math.max(HALIGN_TOL_PT, HALIGN_TOL_RATIO * cellW)
  const cx = (cellBox.x0 + cellBox.x1) / 2
  let widest = 0
  let centered = true
  let righted = true
  for (const line of lines) {
    widest = Math.max(widest, line.box.x1 - line.box.x0)
    if (Math.abs((line.box.x0 + line.box.x1) / 2 - cx) >= tol) centered = false
    if (cellBox.x1 - line.box.x1 >= tol || line.box.x0 - cellBox.x0 < tol) righted = false
  }
  if (cellW <= widest * HALIGN_MIN_SLACK) return undefined
  if (centered) return 'center'
  if (righted) return 'right'
  return undefined
}

const WHITE = /^F[EF]F[EF]F[EF]$/i

/**
 * Dominant stroke color across the grid's lines (length-weighted). White or
 * colored rulings (zebra decks separate their fills with white lines) would
 * otherwise render as Word's default black and repaint the whole table dark.
 * Black stays undefined — the default borders already render black.
 */
function latticeBorderColor(gridBox: Rect, strokes: readonly Stroke[]): string | undefined {
  const byColor = new Map<string, number>()
  let total = 0
  for (const s of strokes) {
    if (s.fromForm) continue
    const cx = rectCenterX(s.box)
    const cy = rectCenterY(s.box)
    if (cx < gridBox.x0 - CONNECT_TOL || cx > gridBox.x1 + CONNECT_TOL) continue
    if (cy < gridBox.y0 - CONNECT_TOL || cy > gridBox.y1 + CONNECT_TOL) continue
    const len = Math.max(s.box.x1 - s.box.x0, s.box.y1 - s.box.y0)
    total += len
    byColor.set(s.color, (byColor.get(s.color) ?? 0) + len)
  }
  if (total <= 0) return undefined
  let best: string | undefined
  let bestLen = 0
  for (const [color, len] of byColor) {
    if (len > bestLen) {
      best = color
      bestLen = len
    }
  }
  if (best === undefined || bestLen / total < BORDER_COLOR_DOMINANCE) return undefined
  return best === '000000' ? undefined : best
}

/** last-painted non-white fill covering most of the cell (paint order = array order) */
function cellFill(cellBox: Rect, fills: readonly Fill[], tableArea: number): string | undefined {
  let color: string | undefined
  for (const fill of fills) {
    if (rectArea(fill.box) > FILL_MAX_AREA_RATIO * tableArea) continue
    if (WHITE.test(fill.color)) continue
    if (overlapRatio(cellBox, fill.box) > 0.5) color = fill.color
  }
  return color
}

export interface DetectedTables {
  tables: TableBlock[]
  /** chars that stay in the regular paragraph flow */
  remainingChars: PdfChar[]
}

/** form-XObject stroke groups must solve into a grid at least this rich (P27) */
const FORM_GRID_MIN_ROWS = 3
const FORM_GRID_MIN_COLS = 3
/** a grid spanning this share of BOTH page dimensions is a page frame, not a table */
const PAGE_FRAME_COVER = 0.96

/**
 * Solve a page's strokes into candidate lattice grids. Top-level strokes work
 * exactly as before (P14 C: form-XObject dividers minting a bordered 2×2 stay
 * out). Form-XObject strokes get a SECOND, stricter chance (P27): whole
 * tables arrive wrapped in forms (tabula twotables' financial tables), so a
 * form-stroke group that solves into a rich ≥3×3 non-lowRank grid is a real
 * table — a lone divider cross can never reach that rank.
 */
export function solveLatticeGrids(strokes: readonly Stroke[]): TableGrid[] {
  const grids: TableGrid[] = []
  for (const group of groupStrokes(strokes.filter((s) => !s.fromForm))) {
    const grid = solveGrid(group, true)
    if (grid) grids.push(grid)
  }
  for (const group of groupStrokes(strokes.filter((s) => s.fromForm))) {
    const grid = solveGrid(group, false)
    if (
      grid &&
      grid.ys.length - 1 >= FORM_GRID_MIN_ROWS &&
      grid.xs.length - 1 >= FORM_GRID_MIN_COLS
    ) {
      grids.push(grid)
    }
  }
  return grids
}

/** Full lattice pass over one page's shapes + chars. */
/** share of a char's width that must overhang BOTH sides of a boundary to count as crossing */
const SPLIT_CROSS_SHARE = 0.3

/** minimum whitespace gap (in em, via char height) required at a column boundary to split a merged run */
const SPLIT_BOUNDARY_GAP_EM = 0.5

/** vertical strokes that must break at the same y to imply a row boundary (P27) */
const JUNCTION_MIN_STROKES = 3

/**
 * Harvest implied row boundaries from vline junctions (P27): some tables
 * draw their cell sides as per-row vertical segments but omit the horizontal
 * rules between data rows (us-008's Head Start grid) — the solved grid then
 * has 2 giant rows and every data row's text piles into one band. Where at
 * least JUNCTION_MIN_STROKES vertical lines break at the same y, and no char
 * straddles that y, insert a virtual row boundary (virtual ⇒ borderPresent,
 * so layoutCells keeps the rows apart — same contract as the implied frame).
 */
function augmentGridRowBoundaries(grid: TableGrid, chars: readonly PdfChar[]): void {
  const ends: number[] = []
  for (const line of grid.vLines) {
    if (line.virtual) continue
    for (const [a, b] of line.segments) {
      ends.push(Math.min(a, b))
      ends.push(Math.max(a, b))
    }
  }
  if (ends.length === 0) return
  ends.sort((a, b) => a - b)
  const clusters: Array<{ pos: number; n: number }> = []
  for (const e of ends) {
    const last = clusters[clusters.length - 1]
    if (last && e - last.pos <= POS_TOL) {
      last.pos = (last.pos + e) / 2
      last.n++
    } else {
      clusters.push({ pos: e, n: 1 })
    }
  }
  const inner: number[] = []
  for (const c of clusters) {
    if (c.n < JUNCTION_MIN_STROKES) continue
    if (c.pos <= grid.box.y0 + MIN_CELL_DIM || c.pos >= grid.box.y1 - MIN_CELL_DIM) continue
    if (grid.ys.some((y) => Math.abs(y - c.pos) < MIN_CELL_DIM)) continue
    // text must respect the boundary — a straddling char refutes the row
    // split (and kills dash-pattern phantom junctions in one stroke)
    let crossed = false
    for (const ch of chars) {
      if (ch.code <= 0x20) continue
      const over = SPLIT_CROSS_SHARE * (ch.box.y1 - ch.box.y0)
      if (ch.box.y0 < c.pos - over && ch.box.y1 > c.pos + over) {
        crossed = true
        break
      }
    }
    if (!crossed) inner.push(c.pos)
  }
  if (inner.length === 0) return
  for (const pos of inner) {
    grid.hLines.push({ pos, segments: [[grid.box.x0, grid.box.x1]], virtual: true })
  }
  grid.hLines.sort((a, b) => b.pos - a.pos) // top → bottom, like solveGrid
  grid.ys = grid.hLines.map((l) => l.pos)
}

/**
 * Split a merged horizontal run back into per-column cells (P27): tables in
 * the "vertical rules only in the header" style (EFSA country tables) merge
 * every data row into one full-width cell, and the whole row's text lands in
 * a single cell. When none of the run's own characters straddles an interior
 * column boundary, every populated boundary shows a real whitespace gap, and
 * at least two of the resulting columns hold text, the merge is a drawing
 * omission rather than a colspan — return the per-column char partition.
 * Genuine spanning cells (titles crossing a boundary, content confined to
 * one column, or continuous prose flowing across boundaries — CJK sentences
 * have no straddling char yet must never be chopped into columns) return
 * null and keep their merge.
 */
function splitRunByColumns(
  xs: readonly number[],
  c0: number,
  colSpan: number,
  chars: readonly PdfChar[],
): PdfChar[][] | null {
  const parts: PdfChar[][] = Array.from({ length: colSpan }, () => [])
  for (const ch of chars) {
    const w = ch.box.x1 - ch.box.x0
    if (ch.code > 0x20) {
      for (let b = 1; b < colSpan; b++) {
        const bx = xs[c0 + b]!
        const over = SPLIT_CROSS_SHARE * w
        if (ch.box.x0 < bx - over && ch.box.x1 > bx + over) return null
      }
    }
    const cx = rectCenterX(ch.box)
    let k = colSpan - 1
    for (let b = 1; b < colSpan; b++) {
      if (cx < xs[c0 + b]!) {
        k = b - 1
        break
      }
    }
    parts[k]!.push(ch)
  }
  let populated = 0
  for (const p of parts) if (p.some((ch) => ch.code > 0x20)) populated++
  if (populated < 2) return null
  // continuous prose guard: a populated boundary with no whitespace gap means
  // the text flows across it (CJK chars never straddle, so the crossing test
  // alone would happily shred a sentence into per-column fragments)
  let emSum = 0
  let emN = 0
  for (const ch of chars) {
    if (ch.code > 0x20) {
      emSum += ch.box.y1 - ch.box.y0
      emN++
    }
  }
  const minGap = SPLIT_BOUNDARY_GAP_EM * (emN > 0 ? emSum / emN : 0)
  for (let b = 1; b < colSpan; b++) {
    const bx = xs[c0 + b]!
    let leftEdge = -Infinity
    let rightEdge = Infinity
    for (const ch of chars) {
      if (ch.code <= 0x20) continue
      if (rectCenterX(ch.box) < bx) leftEdge = Math.max(leftEdge, ch.box.x1)
      else rightEdge = Math.min(rightEdge, ch.box.x0)
    }
    if (leftEdge === -Infinity || rightEdge === Infinity) continue // one-sided boundary
    if (rightEdge - leftEdge < minGap) return null
  }
  return parts
}

export function detectTables(
  shapes: PageShapes,
  chars: readonly PdfChar[],
  pageHeightPt?: number,
  pageWidthPt?: number,
): DetectedTables {
  let grids: TableGrid[] = []
  for (const grid of solveLatticeGrids(shapes.strokes)) {
    grids.push(grid.lowRank ? grid : trimGhostEdgeColumns(grid, chars, shapes.fills))
  }
  // a grid spanning ~the whole page is a page frame / certificate border, not
  // a table — page-edge rules grouped into one grid otherwise mint a ghost
  // whole-page table that swallows every char on the page (camelot assam, P27)
  if (pageWidthPt !== undefined && pageHeightPt !== undefined) {
    grids = grids.filter(
      (g) =>
        !(
          g.box.x1 - g.box.x0 >= PAGE_FRAME_COVER * pageWidthPt &&
          g.box.y1 - g.box.y0 >= PAGE_FRAME_COVER * pageHeightPt
        ),
    )
  }
  // vet the low-rank boxes (P18 C): a frame around another grid is a page
  // decoration (fdo80097's outer frames), a near-page-tall box is a page
  // frame, and a box without text is extractEmptyFrames territory — all of
  // those keep today's degrade path
  const inside = (cx: number, cy: number, r: Rect): boolean =>
    cx >= r.x0 && cx <= r.x1 && cy >= r.y0 && cy <= r.y1
  grids = grids.filter((g) => {
    if (!g.lowRank) return true
    // landscape pages are slide territory: their outlined stat cards live in
    // pinned column layouts, and a flow table there shatters the slide
    // (the baijiu-trends sample's stat cards -0.097). Portrait forms keep the box rule.
    if (pageWidthPt !== undefined && pageHeightPt !== undefined && pageWidthPt > pageHeightPt)
      return false
    if (pageHeightPt !== undefined && g.box.y1 - g.box.y0 > BOX_MAX_PAGE_SHARE * pageHeightPt)
      return false
    if (g.box.y1 - g.box.y0 < BOX_MIN_H_PT) return false
    if (grids.some((o) => o !== g && inside(rectCenterX(o.box), rectCenterY(o.box), g.box)))
      return false
    let n = 0
    for (const c of chars) {
      if (c.code > 0x20 && inside(rectCenterX(c.box), rectCenterY(c.box), g.box)) n++
      if (n >= BOX_MIN_CHARS) return true
    }
    return false
  })
  if (grids.length === 0) return { tables: [], remainingChars: [...chars] }

  const remainingChars: PdfChar[] = []
  const perGrid = grids.map(() => [] as PdfChar[])
  // smallest grid first: with nested/overlapping grids a char belongs to the
  // innermost one, or an outer frame grid swallows the inner table's text (P27)
  const routeOrder = [...grids.entries()].sort(([, a], [, b]) => rectArea(a.box) - rectArea(b.box))
  outer: for (const c of chars) {
    const cx = rectCenterX(c.box)
    const cy = rectCenterY(c.box)
    for (const [i, grid] of routeOrder) {
      if (cx >= grid.box.x0 && cx <= grid.box.x1 && cy >= grid.box.y0 && cy <= grid.box.y1) {
        perGrid[i]!.push(c)
        continue outer
      }
    }
    remainingChars.push(c)
  }

  const tables: TableBlock[] = []
  for (const [i, grid] of grids.entries()) {
    // per-row cell sides without horizontal rules imply the rows (P27)
    if (!grid.lowRank) augmentGridRowBoundaries(grid, perGrid[i]!)
    const layout = layoutCells(grid)
    const tableArea = rectArea(grid.box)

    // route each char to its anchor cell (clamped into the grid)
    const cellChars = new Map<string, PdfChar[]>()
    for (const c of perGrid[i]!) {
      const cy = rectCenterY(c.box)
      const cx = rectCenterX(c.box)
      let r = grid.ys.findIndex((y, idx) => idx > 0 && cy >= y) - 1
      if (r < 0) r = cy > grid.ys[0]! ? 0 : layout.rows - 1
      r = Math.min(Math.max(r, 0), layout.rows - 1)
      let col = grid.xs.findIndex((x, idx) => idx > 0 && cx <= x) - 1
      if (col < 0) col = cx < grid.xs[0]! ? 0 : layout.cols - 1
      col = Math.min(Math.max(col, 0), layout.cols - 1)
      const anchor = layout.anchorOf[r]![col]!
      const key = `${anchor.r}:${anchor.c}`
      let list = cellChars.get(key)
      if (!list) cellChars.set(key, (list = []))
      list.push(c)
    }

    const rows: TableCellBlock[][] = []
    for (let r = 0; r < layout.rows; r++) {
      const rowCells: TableCellBlock[] = []
      for (let c = 0; c < layout.cols; c++) {
        const anchor = layout.anchorOf[r]![c]!
        if (anchor.c !== c) continue // horizontally covered by the cell to the left
        const span = layout.spanOf.get(`${anchor.r}:${anchor.c}`)!
        const box: Rect = {
          x0: grid.xs[c]!,
          x1: grid.xs[c + span.colSpan]!,
          y0: grid.ys[anchor.r + span.rowSpan]!,
          y1: grid.ys[anchor.r]!,
        }
        const cell: TableCellBlock = { box, gridSpan: span.colSpan, blocks: [] }
        if (anchor.r === r) {
          if (span.rowSpan > 1) cell.vMerge = 'restart'
          const inside = cellChars.get(`${r}:${c}`) ?? []
          // header-only vertical rules (P27): a single-row merged run whose
          // chars respect the column boundaries splits back into columns
          if (span.rowSpan === 1 && span.colSpan > 1 && inside.length > 0) {
            const parts = splitRunByColumns(grid.xs, c, span.colSpan, inside)
            if (parts) {
              for (let k = 0; k < span.colSpan; k++) {
                const partBox: Rect = {
                  x0: grid.xs[c + k]!,
                  x1: grid.xs[c + k + 1]!,
                  y0: box.y0,
                  y1: box.y1,
                }
                const part: TableCellBlock = {
                  box: partBox,
                  gridSpan: 1,
                  blocks: [],
                  softEdges: {
                    ...(k > 0 ? { left: true } : {}),
                    ...(k < span.colSpan - 1 ? { right: true } : {}),
                  },
                }
                const partChars = parts[k]!
                if (partChars.length > 0) {
                  part.blocks = groupIntoBlocks(analyzeChars(partChars))
                  const vAlign = detectCellVAlign(partBox, partChars)
                  if (vAlign) part.vAlign = vAlign
                }
                const partFill = cellFill(partBox, shapes.fills, tableArea)
                if (partFill) part.fill = partFill
                rowCells.push(part)
              }
              continue
            }
          }
          if (inside.length > 0) {
            cell.blocks = groupIntoBlocks(analyzeChars(inside))
            const vAlign = detectCellVAlign(box, inside)
            if (vAlign) cell.vAlign = vAlign
          }
          const fill = cellFill(box, shapes.fills, tableArea)
          if (fill) cell.fill = fill
        } else {
          cell.vMerge = 'continue' // covered row of a vertical merge
        }
        rowCells.push(cell)
      }
      rows.push(rowCells)
    }

    const colWidthsPt: number[] = []
    for (let c = 0; c < layout.cols; c++) colWidthsPt.push(grid.xs[c + 1]! - grid.xs[c]!)
    const borderColor = latticeBorderColor(grid.box, shapes.strokes)
    tables.push({
      kind: 'table',
      box: grid.box,
      colWidthsPt,
      rows,
      ...(borderColor !== undefined ? { borderColor } : {}),
    })
  }

  // top-down page order
  tables.sort((a, b) => b.box.y1 - a.box.y1)
  return { tables, remainingChars }
}
