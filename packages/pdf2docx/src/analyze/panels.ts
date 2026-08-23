/**
 * Side-by-side lattice panels → one unified grid (P28). Dense forms draw
 * several ruled panels beside each other with hair gaps (4–6pt) — far below
 * any column gutter, so the flow rebuild stacked them vertically and the page
 * exploded (the JA civil-record form rebuilt 2 pages as 3). Merging keeps the
 * panels beside each other inside ONE table: column boundaries concatenate
 * (the gap becomes a borderless filler column), row boundaries take the
 * union, cells spanning several union rows become vMerge runs, and bands a
 * shorter panel never reaches fill with edge-suppressed empty cells.
 */
import type { Rect } from '../geometry'
import { rectHeight, rectUnionAll } from '../geometry'
import type { TableBlock, TableCellBlock } from '../ir'

/** panels sit this close or closer (pt); real column gutters stay apart */
const MAX_GAP_PT = 16
/** minimum y-overlap share of the SHORTER panel */
const MIN_OVERLAP_SHARE = 0.5
/** a panel is a real grid, not a stray strip */
const MIN_PANEL_ROWS = 2
/** boundary snap tolerance (pt) */
const SNAP_PT = 1.0

interface PanelRows {
  table: TableBlock
  /** row boundaries, top → bottom (y-up: descending) */
  ys: number[]
  /** column boundaries, left → right */
  xs: number[]
}

/** recover a table's row boundaries from its cell boxes; null = malformed */
function rowBoundaries(t: TableBlock): number[] | null {
  const ys: number[] = [t.box.y1]
  for (let r = 1; r < t.rows.length; r++) {
    const tops = t.rows[r]!.filter((c) => c.vMerge !== 'continue').map((c) => c.box.y1)
    if (tops.length > 0) {
      ys.push(Math.max(...tops))
      continue
    }
    const bottomsAbove = t.rows[r - 1]!.filter((c) => c.vMerge === undefined).map((c) => c.box.y0)
    if (bottomsAbove.length === 0) return null
    ys.push(Math.min(...bottomsAbove))
  }
  ys.push(t.box.y0)
  for (let i = 1; i < ys.length; i++) if (ys[i]! >= ys[i - 1]! - 0.1) return null
  return ys
}

function colBoundaries(t: TableBlock): number[] {
  const xs = [t.box.x0]
  for (const w of t.colWidthsPt) xs.push(xs[xs.length - 1]! + w)
  return xs
}

const overlapShare = (a: Rect, b: Rect): number => {
  const overlap = Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0)
  return overlap / Math.max(1, Math.min(rectHeight(a), rectHeight(b)))
}

/** chain tables left → right while each joint is a hair gap with tall overlap */
function groupPanels(tables: TableBlock[]): TableBlock[][] {
  const sorted = [...tables].sort((a, b) => a.box.x0 - b.box.x0)
  const groups: TableBlock[][] = []
  for (const t of sorted) {
    const g = groups.find((group) => {
      const last = group[group.length - 1]!
      const gap = t.box.x0 - last.box.x1
      return gap > -0.5 && gap <= MAX_GAP_PT && overlapShare(t.box, last.box) >= MIN_OVERLAP_SHARE
    })
    if (g) g.push(t)
    else groups.push([t])
  }
  return groups.filter((g) => g.length >= 2)
}

const snapIndex = (bounds: number[], v: number): number => {
  let best = 0
  for (let i = 1; i < bounds.length; i++) {
    if (Math.abs(bounds[i]! - v) < Math.abs(bounds[best]! - v)) best = i
  }
  return best
}

function mergeGroup(group: TableBlock[]): TableBlock | null {
  const panels: PanelRows[] = []
  for (const t of group) {
    const ys = rowBoundaries(t)
    if (ys === null || t.rows.length < MIN_PANEL_ROWS) return null
    panels.push({ table: t, ys, xs: colBoundaries(t) })
  }

  // merged column boundaries: concatenated panels, hair gaps become columns
  const xs: number[] = [panels[0]!.xs[0]!]
  for (const p of panels) {
    for (const x of p.xs) if (x > xs[xs.length - 1]! + SNAP_PT) xs.push(x)
  }
  // merged row boundaries: union, snapped
  const allYs = panels.flatMap((p) => p.ys).sort((a, b) => b - a)
  const ys: number[] = []
  for (const y of allYs) {
    if (ys.length === 0 || y < ys[ys.length - 1]! - SNAP_PT) ys.push(y)
  }
  if (xs.length < 3 || ys.length < 3) return null

  // per-panel merged-grid column bands
  const bandOf = panels.map((p) => ({
    lo: snapIndex(xs, p.xs[0]!),
    hi: snapIndex(xs, p.xs[p.xs.length - 1]!),
  }))

  type Placed = { cell: TableCellBlock; c0: number; c1: number; r0: number; r1: number }
  const placed: Placed[] = []
  for (const [pi, p] of panels.entries()) {
    const off = bandOf[pi]!
    for (const row of p.table.rows) {
      for (const cell of row) {
        if (cell.vMerge === 'continue') continue
        const c0 = snapIndex(xs, cell.box.x0)
        const c1 = snapIndex(xs, cell.box.x1)
        const r0 = snapIndex(ys, cell.box.y1)
        const r1 = snapIndex(ys, cell.box.y0)
        if (c1 <= c0 || r1 <= r0 || c0 < off.lo || c1 > off.hi) return null
        placed.push({ cell, c0, c1, r0, r1 })
      }
    }
  }

  // paint the merged grid; unclaimed slots become border-suppressed fillers
  const owner: Array<Array<Placed | null>> = Array.from({ length: ys.length - 1 }, () =>
    Array.from({ length: xs.length - 1 }, () => null),
  )
  for (const pl of placed) {
    for (let r = pl.r0; r < pl.r1; r++) {
      for (let c = pl.c0; c < pl.c1; c++) {
        if (owner[r]![c] !== null) return null // overlapping cells — bail out
        owner[r]![c] = pl
      }
    }
  }

  const rows: TableCellBlock[][] = []
  for (let r = 0; r < ys.length - 1; r++) {
    const rowCells: TableCellBlock[] = []
    for (let c = 0; c < xs.length - 1;) {
      const pl = owner[r]![c]
      if (pl !== null) {
        if (c !== pl.c0) return null // mid-cell entry — geometry disagrees
        const span = pl.c1 - pl.c0
        if (r === pl.r0) {
          const cell: TableCellBlock = { ...pl.cell, gridSpan: span }
          if (pl.r1 - pl.r0 > 1) cell.vMerge = 'restart'
          else delete cell.vMerge
          rowCells.push(cell)
        } else {
          rowCells.push({ box: pl.cell.box, gridSpan: span, vMerge: 'continue', blocks: [] })
        }
        c = pl.c1
        continue
      }
      // greedy run of unclaimed slots → one filler cell
      let cEnd = c
      while (cEnd < xs.length - 1 && owner[r]![cEnd] === null) cEnd++
      rowCells.push({
        box: { x0: xs[c]!, x1: xs[cEnd]!, y0: ys[r + 1]!, y1: ys[r]! },
        gridSpan: cEnd - c,
        blocks: [],
        softEdges: { left: true, right: true, top: true, bottom: true },
      })
      c = cEnd
    }
    rows.push(rowCells)
  }

  // every union row must carry a height-bearing cell (vMerge undefined): the
  // rebuild measures row heights off plain cells only, and a row of pure merge
  // placeholders (adjacent panels, no gap filler) would leave trHeight null
  // for Word to auto-size past the source
  for (const row of rows) {
    if (!row.some((cell) => cell.vMerge === undefined)) return null
  }

  const colWidthsPt: number[] = []
  for (let c = 1; c < xs.length; c++) colWidthsPt.push(xs[c]! - xs[c - 1]!)
  const borderColor = group.find((t) => t.borderColor !== undefined)?.borderColor
  return {
    kind: 'table',
    box: rectUnionAll(group.map((t) => t.box)),
    colWidthsPt,
    rows,
    ...(borderColor !== undefined ? { borderColor } : {}),
  }
}

/**
 * Merge hair-gap side-by-side lattice tables into unified grids. Returns the
 * new table list (merged grids replace their members) plus a note per merge.
 */
export function mergeSideBySidePanels(tables: TableBlock[]): {
  tables: TableBlock[]
  notes: string[]
} {
  const lattice = tables.filter((t) => t.confidence === undefined)
  if (lattice.length < 2) return { tables, notes: [] }
  const notes: string[] = []
  const replaced = new Map<TableBlock, TableBlock | null>()
  for (const group of groupPanels(lattice)) {
    const merged = mergeGroup(group)
    if (merged === null) continue
    replaced.set(group[0]!, merged)
    for (const t of group.slice(1)) replaced.set(t, null)
    notes.push(`${group.length} side-by-side panels merged into one grid`)
  }
  if (replaced.size === 0) return { tables, notes: [] }
  const out: TableBlock[] = []
  for (const t of tables) {
    if (!replaced.has(t)) {
      out.push(t)
      continue
    }
    const sub = replaced.get(t)
    if (sub !== null && sub !== undefined) out.push(sub)
  }
  return { tables: out, notes }
}
