/** Side-by-side lattice panel merge (P28): hand-built tables, no wasm. */
import { describe, expect, it } from 'vitest'
import { mergeSideBySidePanels } from '../src/analyze/panels'
import type { Rect } from '../src/geometry'
import type { TableBlock, TableCellBlock } from '../src/ir'

/** uniform lattice grid: box + row/col boundary lists (y descending, x ascending) */
function grid(xs: number[], ys: number[]): TableBlock {
  const rows: TableCellBlock[][] = []
  for (let r = 0; r + 1 < ys.length; r++) {
    const row: TableCellBlock[] = []
    for (let c = 0; c + 1 < xs.length; c++) {
      const box: Rect = { x0: xs[c]!, x1: xs[c + 1]!, y0: ys[r + 1]!, y1: ys[r]! }
      row.push({ box, gridSpan: 1, blocks: [] })
    }
    rows.push(row)
  }
  const colWidthsPt = xs.slice(1).map((x, i) => x - xs[i]!)
  return {
    kind: 'table',
    box: { x0: xs[0]!, x1: xs[xs.length - 1]!, y0: ys[ys.length - 1]!, y1: ys[0]! },
    colWidthsPt,
    rows,
  }
}

describe('mergeSideBySidePanels', () => {
  it('merges three hair-gap panels into one grid with union rows', () => {
    // panels like the JA civil-record form: same top, third panel shorter
    const a = grid([15, 150, 279], [429, 300, 100, 15])
    const b = grid([283, 400, 547], [429, 250, 15])
    const c = grid([552, 700, 826], [429, 280, 148])
    const { tables, notes } = mergeSideBySidePanels([a, b, c])
    expect(tables).toHaveLength(1)
    expect(notes).toEqual(['3 side-by-side panels merged into one grid'])
    const merged = tables[0]!
    expect(merged.box).toEqual({ x0: 15, x1: 826, y0: 15, y1: 429 })
    // columns: 2+gap+2+gap+2 = 8
    expect(merged.colWidthsPt).toHaveLength(8)
    expect(merged.colWidthsPt[2]!).toBeCloseTo(4, 5) // 279→283 gap
    // union rows: 429, 300, 280, 250, 148, 100, 15 → 6 rows
    expect(merged.rows).toHaveLength(6)
    const row0 = merged.rows[0]!
    // full grid width is covered on every row
    for (const row of merged.rows) {
      const width = row.reduce((t, cell) => t + cell.gridSpan, 0)
      expect(width).toBe(8)
    }
    // panel C is absent below y=148: bottom rows carry border-suppressed fillers there
    const bottomRow = merged.rows[5]!
    const filler = bottomRow[bottomRow.length - 1]!
    expect(filler.softEdges).toEqual({ left: true, right: true, top: true, bottom: true })
    expect(filler.blocks).toHaveLength(0)
    // spanning cells open with restart on their top row
    expect(row0.some((cell) => cell.vMerge === 'restart')).toBe(true)
  })

  it('keeps tables apart across a real column gutter', () => {
    const a = grid([15, 100, 200], [400, 300, 200])
    const b = grid([260, 350, 450], [400, 300, 200]) // 60pt gap = layout, not panels
    const { tables, notes } = mergeSideBySidePanels([a, b])
    expect(tables).toHaveLength(2)
    expect(notes).toHaveLength(0)
  })

  it('keeps vertically stacked tables apart', () => {
    const a = grid([15, 100, 200], [400, 300, 250])
    const b = grid([15, 100, 200], [200, 150, 100])
    const { tables } = mergeSideBySidePanels([a, b])
    expect(tables).toHaveLength(2)
  })

  it('leaves stream tables (confidence set) alone', () => {
    const a = grid([15, 100, 200], [400, 300, 200])
    const b = { ...grid([204, 300, 400], [400, 300, 200]), confidence: 0.8 }
    const { tables } = mergeSideBySidePanels([a, b])
    expect(tables).toHaveLength(2)
  })

  it('preserves cell payloads and fills through the merge', () => {
    const a = grid([15, 100, 200], [400, 300, 200])
    a.rows[0]![0]!.fill = 'ddeeff'
    a.rows[0]![0]!.vAlign = 'center'
    const b = grid([204, 300, 400], [400, 300, 200])
    const { tables } = mergeSideBySidePanels([a, b])
    expect(tables).toHaveLength(1)
    const cell = tables[0]!.rows[0]![0]!
    expect(cell.fill).toBe('ddeeff')
    expect(cell.vAlign).toBe('center')
  })
})

describe('mergeSideBySidePanels row-height safety', () => {
  it('declines when a union row would hold only merge placeholders', () => {
    // adjacent panels (gap < snap, no filler column) with offset row edges:
    // the middle union row is covered by spans from both sides only
    const a = grid([0, 100, 200], [400, 300, 0])
    const b = grid([200.3, 300, 400], [400, 280, 0])
    const { tables, notes } = mergeSideBySidePanels([a, b])
    expect(tables).toHaveLength(2)
    expect(notes).toHaveLength(0)
  })
})
