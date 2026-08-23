/**
 * Band-row splitting (P40, cell-data). Rule-less statement tables solve into
 * a header row plus ONE giant body band — every column's transactions glued
 * into a single cell. Real rows are recoverable from text geometry: a data
 * row starts where MOST line-bearing columns start a line at the same y
 * (a wrapped description's second line starts alone and never votes).
 * Splitting is emission-side only: the IR grid, and the docx path, keep the
 * drawn 2-row truth.
 */
import type { Line, TableBlock, TableCellBlock, TextBlock } from '../ir'

/** line tops within this distance cluster into one row-start vote (pt) */
const ROW_TOP_TOL_PT = 2
/** a boundary needs line starts from this share of the line-bearing columns */
const ROW_BOUNDARY_COL_SHARE = 0.5
/** a band must split into at least this many rows — wrapped 2-line cells stay */
const MIN_SPLIT_ROWS = 3
/** the fullest cell must carry at least this many lines to call the row a band */
const BAND_MIN_LINES = 4

interface CellLine {
  line: Line
  src: TextBlock
  col: number
}

const linesOf = (cell: TableCellBlock): Array<{ line: Line; src: TextBlock }> =>
  cell.blocks.flatMap((b) =>
    b.lines
      .filter((l) => l.spans.some((s) => s.text.trim() !== ''))
      .map((line) => ({ line, src: b })),
  )

/** boundaries closer than this collapse into one row (staggered baselines
 * across columns must not mint 2-4pt phantom rows) */
const MIN_ROW_PITCH_PT = 6

interface BoundaryCluster {
  top: number
  cols: Set<number>
  members: CellLine[]
}

/** row-start boundaries: clustered line tops backed by enough distinct
 * columns, then thinned to a minimum row pitch */
function rowBoundaries(all: CellLine[], lineBearingCols: number): BoundaryCluster[] {
  const sorted = [...all].sort((a, b) => b.line.box.y1 - a.line.box.y1)
  const clusters: BoundaryCluster[] = []
  for (const entry of sorted) {
    const last = clusters[clusters.length - 1]
    if (last && last.top - entry.line.box.y1 <= ROW_TOP_TOL_PT) {
      last.cols.add(entry.col)
      last.members.push(entry)
    } else {
      clusters.push({ top: entry.line.box.y1, cols: new Set([entry.col]), members: [entry] })
    }
  }
  const need = Math.max(2, Math.ceil(lineBearingCols * ROW_BOUNDARY_COL_SHARE))
  const bounds: BoundaryCluster[] = []
  for (const c of clusters) {
    if (c.cols.size < need) continue
    const last = bounds[bounds.length - 1]
    if (last && last.top - c.top < MIN_ROW_PITCH_PT) continue
    bounds.push(c)
  }
  return bounds
}

function subCell(
  cell: TableCellBlock,
  rowLines: Array<{ line: Line; src: TextBlock }>,
  edge: { top: boolean; bottom: boolean },
  band: { y0: number; y1: number },
): TableCellBlock {
  const softEdges = {
    ...(cell.softEdges ?? {}),
    ...(edge.top ? {} : { top: true }),
    ...(edge.bottom ? {} : { bottom: true }),
  }
  const box = { x0: cell.box.x0, x1: cell.box.x1, y0: band.y0, y1: band.y1 }
  const base: TableCellBlock = {
    box,
    gridSpan: cell.gridSpan,
    ...(cell.fill ? { fill: cell.fill } : {}),
    ...(cell.vAlign ? { vAlign: cell.vAlign } : {}),
    ...(Object.keys(softEdges).length > 0 ? { softEdges } : {}),
    blocks: [],
  }
  if (rowLines.length === 0) return base
  const src = rowLines[0]!.src
  const lines = rowLines.map((r) => r.line)
  const lineBox = {
    x0: Math.min(...lines.map((l) => l.box.x0)),
    x1: Math.max(...lines.map((l) => l.box.x1)),
    y0: Math.min(...lines.map((l) => l.box.y0)),
    y1: Math.max(...lines.map((l) => l.box.y1)),
  }
  base.blocks = [
    {
      kind: 'text',
      lines,
      box: lineBox,
      align: src.align,
      firstLineIndentPt: 0,
      dir: src.dir,
    },
  ]
  return base
}

/** split one band row into geometry-derived data rows; null = leave as-is */
function splitRow(row: TableCellBlock[]): TableCellBlock[][] | null {
  if (row.some((c) => c.gridSpan > 1 || c.vMerge !== undefined)) return null
  const perCell = row.map(linesOf)
  const bearing = perCell.filter((l) => l.length > 0).length
  if (bearing < 2) return null
  if (Math.max(...perCell.map((l) => l.length)) < BAND_MIN_LINES) return null
  const all: CellLine[] = perCell.flatMap((ls, col) => ls.map((l) => ({ ...l, col })))
  const bounds = rowBoundaries(all, bearing)
  if (bounds.length < MIN_SPLIT_ROWS) return null
  // IR y grows upward: data row i spans from just above its boundary line
  // down to just above the next one. Boundary voters keep their cluster's
  // row directly — a voter can sit up to ROW_TOP_TOL_PT below the cluster
  // top and interval math alone could push it into the next row. Content
  // above the first boundary (an opening-balance preamble too sparse to
  // vote) becomes its own lead row.
  const memberRow = new Map<Line, number>()
  bounds.forEach((c, i) => {
    for (const m of c.members) memberRow.set(m.line, i)
  })
  const rowIdx = (line: Line): number => {
    const member = memberRow.get(line)
    if (member !== undefined) return member
    const top = line.box.y1
    for (let i = 0; i < bounds.length; i++) {
      if (top > bounds[i]!.top + ROW_TOP_TOL_PT) return i - 1
    }
    return bounds.length - 1
  }
  const hasLead = all.some(({ line }) => rowIdx(line) < 0)
  const lead = hasLead ? 1 : 0
  const total = bounds.length + lead
  const rowY = (i: number): { y0: number; y1: number } => ({
    y1: i === 0 ? Math.max(...row.map((c) => c.box.y1)) : bounds[i - lead]!.top,
    y0:
      i - lead + 1 < bounds.length
        ? bounds[i - lead + 1]!.top
        : Math.min(...row.map((c) => c.box.y0)),
  })
  const out: TableCellBlock[][] = []
  for (let i = 0; i < total; i++) {
    const edge = { top: i === 0, bottom: i === total - 1 }
    out.push(
      row.map((cell, col) =>
        subCell(
          cell,
          perCell[col]!.filter(({ line }) => rowIdx(line) + lead === i),
          edge,
          rowY(i),
        ),
      ),
    )
  }
  return out
}

/** split every qualifying band row of the table; returns the block unchanged
 * when nothing splits */
export function splitBandRows(block: TableBlock): TableBlock {
  let changed = false
  const rows: TableCellBlock[][] = []
  for (const row of block.rows) {
    const split = splitRow(row)
    if (split) {
      rows.push(...split)
      changed = true
    } else {
      rows.push(row)
    }
  }
  return changed ? { ...block, rows } : block
}
