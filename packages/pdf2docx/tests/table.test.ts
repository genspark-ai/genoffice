/** Lattice table detection unit tests: hand-built strokes/fills/chars, no wasm. */
import { describe, expect, it } from 'vitest'
import {
  detectCellHAlign,
  detectCellVAlign,
  detectTables,
  groupStrokes,
  layoutCells,
  solveGrid,
  trimGhostEdgeColumns,
} from '../src/analyze'
import type { PageShapes, Stroke, TableBlock } from '../src/ir'
import { mkText } from './helpers/chars'

const h = (x0: number, x1: number, y: number, w = 1): Stroke => ({
  box: { x0, x1, y0: y - w / 2, y1: y + w / 2 },
  orientation: 'h',
  widthPt: w,
  color: '000000',
})
const v = (y0: number, y1: number, x: number, w = 1): Stroke => ({
  box: { x0: x - w / 2, x1: x + w / 2, y0, y1 },
  orientation: 'v',
  widthPt: w,
  color: '000000',
})

/** full 2-row × 3-col grid: x 100|200|300|400, y 700|650|600 */
const fullGrid = (): Stroke[] => [
  h(100, 400, 700),
  h(100, 400, 650),
  h(100, 400, 600),
  v(600, 700, 100),
  v(600, 700, 200),
  v(600, 700, 300),
  v(600, 700, 400),
]

const shapesOf = (strokes: Stroke[], fills: PageShapes['fills'] = []): PageShapes => ({
  strokes,
  fills,
  ignoredPaths: 0,
})

describe('groupStrokes', () => {
  it('groups touching strokes and separates distant ones', () => {
    const far = h(100, 200, 100)
    const groups = groupStrokes([...fullGrid(), far])
    expect(groups).toHaveLength(2)
    expect(groups.find((g) => g.length === 7)).toBeDefined()
    expect(groups.find((g) => g.length === 1)).toContain(far)
  })
})

describe('solveGrid', () => {
  it('solves a complete grid into row/column boundaries', () => {
    const grid = solveGrid(fullGrid())!
    expect(grid).not.toBeNull()
    expect(grid.xs).toEqual([100, 200, 300, 400])
    expect(grid.ys).toEqual([700, 650, 600])
  })

  it('adds the implied outer frame when edge lines are missing', () => {
    // only the inner lines: one h separator + two inner v lines
    const grid = solveGrid([h(100, 400, 650), v(600, 700, 200), v(600, 700, 300)])!
    expect(grid).not.toBeNull()
    expect(grid.xs).toEqual([100, 200, 300, 400])
    expect(grid.ys).toEqual([700, 650, 600])
  })

  it('rejects one-directional groups (underlines are not tables)', () => {
    expect(solveGrid([h(100, 400, 700), h(100, 400, 650), h(100, 400, 600)])).toBeNull()
  })

  it('rejects 1×N grids', () => {
    // 2 h lines + 4 v lines = 1 row × 3 cols
    const oneRow = [
      h(100, 400, 700),
      h(100, 400, 650),
      v(650, 700, 100),
      v(650, 700, 200),
      v(650, 700, 300),
      v(650, 700, 400),
    ]
    expect(solveGrid(oneRow)).toBeNull()
  })

  it('rejects tiny stroke groups (decorations, checkbox glyphs)', () => {
    const tiny = [h(0, 10, 8), h(0, 10, 4), h(0, 10, 0), v(0, 8, 0), v(0, 8, 5), v(0, 8, 10)]
    expect(solveGrid(tiny)).toBeNull()
  })

  it('clusters slightly misaligned collinear segments into one boundary', () => {
    const strokes = [
      h(100, 250, 700.6),
      h(250, 400, 699.6), // same top line, drawn in two segments 1pt apart
      h(100, 400, 650),
      h(100, 400, 600),
      v(600, 700, 100),
      v(600, 700, 200),
      v(600, 700, 300),
      v(600, 700, 400),
    ]
    const grid = solveGrid(strokes)!
    expect(grid.ys).toHaveLength(3)
  })
})

describe('detectTables: lattice border color (P20)', () => {
  it('carries a dominant non-black stroke color (white zebra rulings)', () => {
    const white = fullGrid().map((s) => ({ ...s, color: 'FFFFFF' }))
    const { tables } = detectTables(shapesOf(white), [])
    expect(tables[0]!.borderColor).toBe('FFFFFF')
  })

  it('black grids keep the default borders (no borderColor)', () => {
    const { tables } = detectTables(shapesOf(fullGrid()), [])
    expect(tables[0]!.borderColor).toBeUndefined()
  })

  it('mixed-color grids below dominance keep the default', () => {
    const mixed = fullGrid().map((s, i) => ({ ...s, color: i % 2 === 0 ? 'FFFFFF' : 'FF0000' }))
    const { tables } = detectTables(shapesOf(mixed), [])
    expect(tables[0]!.borderColor).toBeUndefined()
  })

  it('ignores strokes outside the grid box (a nearby decor line has no vote)', () => {
    const white = fullGrid().map((s) => ({ ...s, color: 'FFFFFF' }))
    const decor = h(100, 400, 750, 2)
    const { tables } = detectTables(shapesOf([...white, decor]), [])
    expect(tables[0]!.borderColor).toBe('FFFFFF')
  })
})

describe('detectTables: closed text boxes (P18 C)', () => {
  // a fully drawn rectangle holding text (a prompt/answer box on a form)
  const box1x1 = (): Stroke[] => [
    h(100, 400, 700),
    h(100, 400, 620),
    v(620, 700, 100),
    v(620, 700, 400),
  ]

  it('accepts a 1×1 box holding text as a one-cell table', () => {
    const inside = mkText('prompt text', 110, { y: 680 }).chars
    const { tables, remainingChars } = detectTables(shapesOf(box1x1()), inside, 792)
    expect(tables).toHaveLength(1)
    expect(tables[0]!.rows).toHaveLength(1)
    expect(tables[0]!.rows[0]).toHaveLength(1)
    expect(remainingChars).toHaveLength(0)
  })

  it('rejects an empty box (extractEmptyFrames territory)', () => {
    const { tables } = detectTables(shapesOf(box1x1()), [], 792)
    expect(tables).toHaveLength(0)
  })

  it('rejects a box that frames another grid (page decoration)', () => {
    const frame = [h(80, 420, 720), h(80, 420, 580), v(580, 720, 80), v(580, 720, 420)]
    const text = mkText('cell', 110, { y: 680 }).chars
    const { tables } = detectTables(shapesOf([...frame, ...fullGrid()]), text, 792)
    // the inner 2×3 lattice survives; the outer frame does not become a box table
    expect(tables).toHaveLength(1)
    expect(tables[0]!.rows).toHaveLength(2)
  })

  it('rejects a near-page-tall box (page frame)', () => {
    const tall = [h(100, 400, 700), h(100, 400, 100), v(100, 700, 100), v(100, 700, 400)]
    const text = mkText('content', 110, { y: 400 }).chars
    const { tables } = detectTables(shapesOf(tall), text, 792)
    expect(tables).toHaveLength(0)
  })

  it('rejects an open box (missing edge stays a decoration)', () => {
    const open = [h(100, 400, 700), h(100, 400, 620), v(620, 700, 100)]
    const text = mkText('quote', 110, { y: 680 }).chars
    const { tables } = detectTables(shapesOf(open), text, 792)
    expect(tables).toHaveLength(0)
  })

  it('accepts a stacked 3×1 run of boxed sections as a table', () => {
    const stack = [
      h(100, 400, 700),
      h(100, 400, 650),
      h(100, 400, 600),
      v(600, 700, 100),
      v(600, 700, 400),
    ]
    const a = mkText('section one', 110, { y: 680 }).chars
    const b = mkText('section two', 110, { y: 630 }).chars
    const { tables } = detectTables(shapesOf(stack), [...a, ...b], 792)
    expect(tables).toHaveLength(1)
    expect(tables[0]!.rows).toHaveLength(2)
    expect(tables[0]!.rows[0]).toHaveLength(1)
  })
})

describe('layoutCells (merged cells from missing borders)', () => {
  it('a full grid yields no merges', () => {
    const layout = layoutCells(solveGrid(fullGrid())!)
    expect(layout.rows).toBe(2)
    expect(layout.cols).toBe(3)
    expect(layout.spanOf.size).toBe(6)
    for (const span of layout.spanOf.values()) {
      expect(span).toEqual({ colSpan: 1, rowSpan: 1 })
    }
  })

  it('missing vertical border → gridSpan grows', () => {
    // v line at x=200 only exists in the top row → bottom row merges cols 0-1
    const strokes = fullGrid().filter(
      (s) => !(s.orientation === 'v' && s.box.x0 < 201 && s.box.x0 > 199),
    )
    strokes.push(v(650, 700, 200))
    const layout = layoutCells(solveGrid(strokes)!)
    expect(layout.spanOf.get('1:0')).toEqual({ colSpan: 2, rowSpan: 1 })
    expect(layout.anchorOf[1]![1]).toEqual({ r: 1, c: 0 })
  })

  it('missing horizontal border → vertical merge', () => {
    // middle h line stops at x=300 → col 2 merges rows 0-1
    const strokes = fullGrid().filter(
      (s) => !(s.orientation === 'h' && s.box.y0 < 651 && s.box.y0 > 649),
    )
    strokes.push(h(100, 300, 650))
    const layout = layoutCells(solveGrid(strokes)!)
    expect(layout.spanOf.get('0:2')).toEqual({ colSpan: 1, rowSpan: 2 })
    expect(layout.anchorOf[1]![2]).toEqual({ r: 0, c: 2 })
  })
})

describe('detectTables', () => {
  it('routes chars into their cells and removes them from the page flow', () => {
    const inA1 = mkText('A1', 110, { y: 670 }).chars
    const inB2 = mkText('B2', 210, { y: 620 }).chars
    const outside = mkText('caption', 100, { y: 560 }).chars
    const { tables, remainingChars } = detectTables(shapesOf(fullGrid()), [
      ...inA1,
      ...inB2,
      ...outside,
    ])
    expect(tables).toHaveLength(1)
    const table = tables[0]!
    expect(table.rows).toHaveLength(2)
    expect(table.rows[0]).toHaveLength(3)
    expect(table.colWidthsPt).toEqual([100, 100, 100])
    const textOf = (r: number, c: number) =>
      table.rows[r]![c]!.blocks.map((b) =>
        b.lines.map((l) => l.spans.map((s) => s.text).join('')).join(' '),
      ).join('\n')
    expect(textOf(0, 0)).toBe('A1')
    expect(textOf(1, 1)).toBe('B2')
    expect(textOf(0, 1)).toBe('')
    expect(remainingChars.map((c) => c.text).join('')).toBe('caption')
  })

  it('emits gridSpan and vMerge placeholders for merged cells', () => {
    // bottom row merges cols 0-1; col 2 merges rows 0-1
    const strokes = [
      h(100, 400, 700),
      h(100, 300, 650), // stops before col 2 → vertical merge in col 2
      h(100, 400, 600),
      v(600, 700, 100),
      v(650, 700, 200), // only top row → bottom row col 0-1 merge
      v(600, 700, 300),
      v(600, 700, 400),
    ]
    const { tables } = detectTables(shapesOf(strokes), [])
    const table = tables[0]!
    // row 0: three cells, col 2 opens a vertical merge
    expect(table.rows[0]!.map((c) => c.gridSpan)).toEqual([1, 1, 1])
    expect(table.rows[0]![2]!.vMerge).toBe('restart')
    // row 1: merged cell (span 2) + vMerge continuation
    expect(table.rows[1]!.map((c) => c.gridSpan)).toEqual([2, 1])
    expect(table.rows[1]![0]!.vMerge).toBeUndefined()
    expect(table.rows[1]![1]!.vMerge).toBe('continue')
  })

  it('matches fills to cells as shading and ignores white/page-size fills', () => {
    const fills = [
      { box: { x0: 100, y0: 650, x1: 200, y1: 700 }, color: 'FFCC00' }, // cell (0,0)
      { box: { x0: 200, y0: 650, x1: 300, y1: 700 }, color: 'FFFFFF' }, // white → skip
      { box: { x0: 0, y0: 0, x1: 612, y1: 792 }, color: 'EEEEEE' }, // page background → skip
    ]
    const { tables } = detectTables(shapesOf(fullGrid(), fills), [])
    const table = tables[0]!
    expect(table.rows[0]![0]!.fill).toBe('FFCC00')
    expect(table.rows[0]![1]!.fill).toBeUndefined()
    expect(table.rows[1]![0]!.fill).toBeUndefined()
  })

  it('leaves underline-only pages table-free (suppression)', () => {
    const chars = mkText('underlined heading', 100, { y: 705 }).chars
    const { tables, remainingChars } = detectTables(
      shapesOf([h(100, 250, 698), h(100, 400, 660), h(100, 400, 620)]),
      chars,
    )
    expect(tables).toHaveLength(0)
    expect(remainingChars).toHaveLength(chars.length)
  })

  it('detects two separate tables and orders them top-down', () => {
    const lower: Stroke[] = [
      h(100, 300, 400),
      h(100, 300, 360),
      h(100, 300, 320),
      v(320, 400, 100),
      v(320, 400, 200),
      v(320, 400, 300),
    ]
    const { tables } = detectTables(shapesOf([...lower, ...fullGrid()]), [])
    expect(tables).toHaveLength(2)
    expect(tables[0]!.box.y1).toBeGreaterThan(tables[1]!.box.y1)
  })
})

describe('detectCellVAlign (P13 A)', () => {
  // 40pt-tall cell; mkChar boxes for fontSize 10 span y ≈ [y-2.1, y+7.2]
  const cell = { x0: 100, x1: 200, y0: 660, y1: 700 }

  it('detects centred text within tolerance', () => {
    // text box [675.4, 684.7], mid ≈ 680.05 vs cell mid 680
    expect(detectCellVAlign(cell, mkText('中', 110, { y: 677.5 }).chars)).toBe('center')
  })

  it('does not mark top-aligned text', () => {
    // text hugs the cell top: mid ≈ 695 vs cell mid 680
    expect(detectCellVAlign(cell, mkText('中', 110, { y: 690 }).chars)).toBeUndefined()
  })

  it('detects bottom-hugging text', () => {
    // text box bottom ≈ 660.4, within tol of the cell bottom 660
    expect(detectCellVAlign(cell, mkText('中', 110, { y: 662.5 }).chars)).toBe('bottom')
  })

  it('stays quiet without head-room (tight cells carry no alignment semantics)', () => {
    const tight = { x0: 100, x1: 200, y0: 675, y1: 686 }
    expect(detectCellVAlign(tight, mkText('中', 110, { y: 677.5 }).chars)).toBeUndefined()
  })

  it('offset beyond the tolerance is not centred', () => {
    // tol = max(1.5, 0.08×40) = 3.2pt; mid ≈ 687 is 7pt above cell mid
    expect(detectCellVAlign(cell, mkText('中', 110, { y: 682 }).chars)).toBeUndefined()
  })

  it('ignores whitespace-only content', () => {
    expect(detectCellVAlign(cell, mkText(' ', 110, { y: 677.5 }).chars)).toBeUndefined()
  })
})

describe('detectCellHAlign', () => {
  // 100pt-wide cell, centre at 150; tol = max(1.5, 0.04×100) = 4pt
  const cell = { x0: 100, x1: 200, y0: 660, y1: 700 }
  const line = (x0: number, x1: number) => ({ box: { x0, x1, y0: 670, y1: 680 } })

  it('detects a centred short line', () => {
    expect(detectCellHAlign(cell, [line(135, 165)])).toBe('center')
  })

  it('detects a right-hugging line (numeric column)', () => {
    expect(detectCellHAlign(cell, [line(160, 199)])).toBe('right')
  })

  it('leaves a left-hugging line unmarked', () => {
    expect(detectCellHAlign(cell, [line(101, 140)])).toBeUndefined()
  })

  it('stays quiet without slack (full-width lines render the same regardless)', () => {
    expect(detectCellHAlign(cell, [line(102, 198)])).toBeUndefined()
  })

  it('requires every line to agree on centre', () => {
    expect(detectCellHAlign(cell, [line(135, 165), line(120, 150)])).toBeUndefined()
  })

  it('multi-line consensus centres', () => {
    expect(detectCellHAlign(cell, [line(135, 165), line(130, 172)])).toBe('center')
  })

  it('is quiet on empty cells', () => {
    expect(detectCellHAlign(cell, [])).toBeUndefined()
  })

  it('offset beyond the tolerance is not centred', () => {
    // centre 155 vs cell centre 150, tol 4pt
    expect(detectCellHAlign(cell, [line(145, 165)])).toBeUndefined()
  })
})

describe('trimGhostEdgeColumns (P13 B)', () => {
  /**
   * repro shape: a 3-row × 2-col grid (x 100|200|300) wrapped by a container
   * frame whose right edge sits at x=390 — internal row separators stop at
   * x=300, so the frame edge mints an empty [300,390] trailing column. Cell
   * borders are drawn per row (HTML renderer style): the inner boundary at
   * x=300 visibly ends the cell-level grid.
   */
  const perRowV = (x: number): Stroke[] => [v(650, 700, x), v(600, 650, x), v(550, 600, x)]
  const framedGrid = (): Stroke[] => [
    // container frame (full width, full height)
    h(100, 390, 700),
    h(100, 390, 550),
    v(550, 700, 100),
    v(550, 700, 390),
    // table grid proper: rows at 650/600 stop at x=300
    h(100, 300, 650),
    h(100, 300, 600),
    ...perRowV(200),
    ...perRowV(300),
  ]

  it('drops a trailing column no internal row boundary reaches into', () => {
    const grid = solveGrid(framedGrid())!
    expect(grid.xs).toEqual([100, 200, 300, 390])
    const trimmed = trimGhostEdgeColumns(grid, [], [])
    expect(trimmed.xs).toEqual([100, 200, 300])
    expect(trimmed.box.x1).toBe(300)
    expect(trimmed.vLines).toHaveLength(3)
  })

  it('a through-running vertical line with crossing row separators forms a column', () => {
    // legit empty column: row separators span the full width
    const strokes = framedGrid().map((s) =>
      s.orientation === 'h' && s.box.x1 <= 301 ? h(100, 390, (s.box.y0 + s.box.y1) / 2) : s,
    )
    const grid = solveGrid(strokes)!
    const trimmed = trimGhostEdgeColumns(grid, [], [])
    expect(trimmed.xs).toEqual([100, 200, 300, 390])
  })

  it('keeps an edge column that holds text', () => {
    const grid = solveGrid(framedGrid())!
    const chars = mkText('印', 320, { y: 620 }).chars
    expect(trimGhostEdgeColumns(grid, chars, []).xs).toEqual([100, 200, 300, 390])
  })

  it('keeps an edge column covered by a fill', () => {
    const grid = solveGrid(framedGrid())!
    const fills = [{ box: { x0: 300, y0: 550, x1: 390, y1: 700 }, color: '1A365D' }]
    expect(trimGhostEdgeColumns(grid, [], fills).xs).toEqual([100, 200, 300, 390])
  })

  it('ignores white and page-scale fills when judging the column', () => {
    const grid = solveGrid(framedGrid())!
    const fills = [
      { box: { x0: 300, y0: 550, x1: 390, y1: 700 }, color: 'FFFFFF' },
      { box: { x0: 0, y0: 0, x1: 612, y1: 792 }, color: 'EEEEEE' },
    ]
    expect(trimGhostEdgeColumns(grid, [], fills).xs).toEqual([100, 200, 300])
  })

  it('drops a leading ghost column symmetrically', () => {
    const strokes: Stroke[] = [
      h(10, 300, 700),
      h(10, 300, 550),
      v(550, 700, 10),
      v(550, 700, 300),
      h(100, 300, 650),
      h(100, 300, 600),
      ...perRowV(100),
      ...perRowV(200),
    ]
    const grid = solveGrid(strokes)!
    expect(grid.xs).toEqual([10, 100, 200, 300])
    expect(trimGhostEdgeColumns(grid, [], []).xs).toEqual([100, 200, 300])
  })

  it('keeps a merged empty edge column bordered by single full-height lines', () => {
    // cross-page continuation form: the empty left columns are real (their
    // borders are drawn, just with no row separators inside) — single
    // full-height boundary lines are no evidence the grid ends there
    const strokes: Stroke[] = [
      h(10, 300, 700),
      h(10, 300, 550),
      v(550, 700, 10),
      v(550, 700, 100), // single full-height line, not per-row
      h(100, 300, 650),
      h(100, 300, 600),
      ...perRowV(200),
      v(550, 700, 300),
    ]
    const grid = solveGrid(strokes)!
    expect(grid.xs).toEqual([10, 100, 200, 300])
    expect(trimGhostEdgeColumns(grid, [], []).xs).toEqual([10, 100, 200, 300])
  })

  it('drops a virtual-edge sliver minted by row-border overshoot', () => {
    // top/bottom border lines overshoot the last drawn column boundary →
    // the implied frame adds a virtual boundary and a sliver column appears
    const strokes: Stroke[] = [
      h(100, 420, 700),
      h(100, 420, 550),
      h(100, 390, 650),
      h(100, 390, 600),
      v(550, 700, 100),
      v(550, 700, 200),
      v(550, 700, 300),
      v(550, 700, 390),
    ]
    const grid = solveGrid(strokes)!
    expect(grid.xs).toEqual([100, 200, 300, 390, 420])
    expect(grid.vLines[4]!.virtual).toBe(true)
    expect(trimGhostEdgeColumns(grid, [], []).xs).toEqual([100, 200, 300, 390])
  })

  it('never trims below two columns', () => {
    // 3-row × 2-col where the second column is ghost-shaped: stays 2 cols
    const strokes: Stroke[] = [
      h(100, 390, 700),
      h(100, 390, 550),
      v(550, 700, 100),
      v(550, 700, 390),
      h(100, 300, 650),
      h(100, 300, 600),
      ...perRowV(300),
    ]
    const grid = solveGrid(strokes)!
    expect(grid.xs).toEqual([100, 300, 390])
    expect(trimGhostEdgeColumns(grid, [], []).xs).toEqual([100, 300, 390])
  })

  it('2-row tables (no internal boundaries) are never trimmed', () => {
    const strokes: Stroke[] = [
      h(100, 390, 700),
      h(100, 390, 600),
      v(600, 700, 100),
      v(600, 700, 200),
      v(600, 700, 300),
      v(600, 700, 390),
      h(100, 300, 650),
    ]
    const grid = solveGrid(strokes)!
    expect(trimGhostEdgeColumns(grid, [], []).xs).toHaveLength(grid.xs.length)
  })

  it('detectTables end-to-end: ghost column gone, vAlign carried on centred cells', () => {
    // 3 rows × 2 cols + container frame; centred text in cell (0,0): row band
    // [650,700], mkChar(10pt) at y=672.5 → box [670.4,679.7], mid 675.05
    const chars = mkText('中心', 110, { y: 672.5 }).chars
    const { tables } = detectTables(shapesOf(framedGrid()), chars)
    expect(tables).toHaveLength(1)
    const table = tables[0]!
    expect(table.rows[0]).toHaveLength(2)
    expect(table.colWidthsPt).toEqual([100, 100])
    expect(table.rows[0]![0]!.vAlign).toBe('center')
  })
})

/** cell text integrates the full char pipeline (words + spans) */
describe('detectTables cell content pipeline', () => {
  it('keeps word spacing inside cells', () => {
    const chars = mkText('two words', 110, { y: 670 }).chars
    const { tables } = detectTables(shapesOf(fullGrid()), chars)
    const cell = (tables[0] as TableBlock).rows[0]![0]!
    const text = cell.blocks[0]!.lines[0]!.spans.map((s) => s.text).join('')
    expect(text).toBe('two words')
  })
})

// ── P27: form-XObject grids, page frames, innermost routing ──

/** 3-row × 3-col grid at x 100..400, y 550..700 built from the given factory */
const grid3x3 = (mk: { h: typeof h; v: typeof v }): Stroke[] => [
  mk.h(100, 400, 700),
  mk.h(100, 400, 650),
  mk.h(100, 400, 600),
  mk.h(100, 400, 550),
  mk.v(550, 700, 100),
  mk.v(550, 700, 200),
  mk.v(550, 700, 300),
  mk.v(550, 700, 400),
]

const asForm = (s: Stroke): Stroke => ({ ...s, fromForm: true })

describe('P27 form-XObject lattice', () => {
  it('accepts a rich >=3x3 grid drawn inside a form XObject', () => {
    const strokes = grid3x3({ h, v }).map(asForm)
    const { tables } = detectTables(shapesOf(strokes), [])
    expect(tables).toHaveLength(1)
    expect(tables[0]!.rows).toHaveLength(3)
    expect(tables[0]!.colWidthsPt).toHaveLength(3)
  })

  it('still rejects a lone form-stroke divider cross (P14 C)', () => {
    const strokes = [h(100, 400, 650), v(550, 700, 250)].map(asForm)
    const { tables } = detectTables(shapesOf(strokes), [])
    expect(tables).toHaveLength(0)
  })
})

describe('P27 page-frame grids', () => {
  it('drops a grid spanning ~the whole page and keeps the inner table', () => {
    // page frame: full-page border + one interior rule through the middle
    const frame = [h(0, 595, 841), h(0, 595, 421), h(0, 595, 1), v(1, 841, 0.5), v(1, 841, 594.5)]
    const inner = grid3x3({ h, v })
    const chars = mkText('cell', 110, { y: 672.5 }).chars
    const { tables } = detectTables(shapesOf([...frame, ...inner]), chars, 842, 595)
    expect(tables).toHaveLength(1)
    expect(tables[0]!.colWidthsPt).toHaveLength(3)
    const text = tables[0]!.rows[0]![0]!.blocks[0]!.lines[0]!.spans.map((s) => s.text).join('')
    expect(text).toBe('cell')
  })
})

describe('P27 innermost char routing', () => {
  it('routes chars to the smallest containing grid', () => {
    // outer grid 3x3 over (50..450, 500..750) not page-covering; inner 3x3
    const outer: Stroke[] = [
      h(50, 450, 750),
      h(50, 450, 730),
      h(50, 450, 500),
      v(500, 750, 50),
      v(500, 750, 430),
      v(500, 750, 450),
    ]
    const inner = grid3x3({ h, v })
    const chars = mkText('inside', 110, { y: 672.5 }).chars
    const { tables } = detectTables(shapesOf([...outer, ...inner]), chars, 842, 595)
    const innerTable = tables.find((t) => t.colWidthsPt.length === 3)!
    const text = innerTable.rows[0]![0]!.blocks[0]?.lines[0]!.spans.map((s) => s.text).join('')
    expect(text).toBe('inside')
  })
})

describe('P27 merged-run split (header-only vertical rules)', () => {
  /** EFSA style: h-rules between every row, vlines only in the header band +
   * full-height outer edges. x 100|200|300|400, header y 680-700, rows to 600. */
  const headerOnlyVRules = (): Stroke[] => [
    h(100, 400, 700),
    h(100, 400, 680),
    h(100, 400, 660),
    h(100, 400, 640),
    h(100, 400, 600),
    v(600, 700, 100),
    v(600, 700, 400),
    v(680, 700, 200),
    v(680, 700, 300),
  ]

  it('splits a merged data row whose chars respect the column boundaries', () => {
    const chars = [
      ...mkText('Head1', 110, { y: 686 }).chars,
      ...mkText('Head2', 210, { y: 686 }).chars,
      ...mkText('Head3', 310, { y: 686 }).chars,
      ...mkText('Austria', 110, { y: 666 }).chars,
      ...mkText('86.2', 210, { y: 666 }).chars,
      ...mkText('13.8', 310, { y: 666 }).chars,
    ]
    const { tables } = detectTables(shapesOf(headerOnlyVRules()), chars)
    expect(tables).toHaveLength(1)
    const row = tables[0]!.rows[1]!
    expect(row).toHaveLength(3)
    expect(row.map((c) => c.gridSpan)).toEqual([1, 1, 1])
    expect(row[0]!.softEdges).toEqual({ right: true })
    expect(row[1]!.softEdges).toEqual({ left: true, right: true })
    expect(row[2]!.softEdges).toEqual({ left: true })
    const text = (c: (typeof row)[number]): string =>
      c.blocks.map((b) => b.lines.map((l) => l.spans.map((s) => s.text).join('')).join('')).join('')
    expect(text(row[0]!)).toBe('Austria')
    expect(text(row[1]!)).toBe('86.2')
    expect(text(row[2]!)).toBe('13.8')
  })

  it('keeps a genuine colspan whose text crosses a column boundary', () => {
    const chars = [
      ...mkText('Head1', 110, { y: 686 }).chars,
      ...mkText('Head2', 210, { y: 686 }).chars,
      ...mkText('Head3', 310, { y: 686 }).chars,
      // crosses the x=200 boundary: stays one merged cell
      ...mkText('Spanning subtotal title', 150, { y: 666 }).chars,
    ]
    const { tables } = detectTables(shapesOf(headerOnlyVRules()), chars)
    expect(tables).toHaveLength(1)
    const row = tables[0]!.rows[1]!
    expect(row).toHaveLength(1)
    expect(row[0]!.gridSpan).toBe(3)
    expect(row[0]!.softEdges).toBeUndefined()
  })

  it('keeps a merged run whose text sits in a single column', () => {
    const chars = [
      ...mkText('Head1', 110, { y: 686 }).chars,
      ...mkText('Head2', 210, { y: 686 }).chars,
      ...mkText('Head3', 310, { y: 686 }).chars,
      ...mkText('Total', 110, { y: 666 }).chars, // one column only: no split
    ]
    const { tables } = detectTables(shapesOf(headerOnlyVRules()), chars)
    const row = tables[0]!.rows[1]!
    expect(row).toHaveLength(1)
    expect(row[0]!.gridSpan).toBe(3)
  })

  it('keeps continuous CJK prose that flows across a boundary without straddling it', () => {
    // fullwidth glyphs advance edge-to-edge: none straddles x=200, yet there is
    // no whitespace gap at the boundary — chopping the sentence into columns
    // fragments it and reflows the docx (form-gov P0202503… page spill, P27)
    const chars = [
      ...mkText('Head1', 110, { y: 686 }).chars,
      ...mkText('Head2', 210, { y: 686 }).chars,
      ...mkText('Head3', 310, { y: 686 }).chars,
      ...mkText('说明企业应当根据实际情况选择填报', 110, { y: 666 }).chars,
    ]
    const { tables } = detectTables(shapesOf(headerOnlyVRules()), chars)
    const row = tables[0]!.rows[1]!
    expect(row).toHaveLength(1)
    expect(row[0]!.gridSpan).toBe(3)
    expect(row[0]!.softEdges).toBeUndefined()
  })
})

describe('P27 vline-junction row harvest', () => {
  /** us-008 style: h-rules only at top/header/bottom; vlines drawn per row so
   * their segment breaks imply the missing row boundaries. */
  const perRowVSegments = (): Stroke[] => {
    const xs = [100, 200, 300, 400]
    const bands: Array<[number, number]> = [
      [680, 700],
      [660, 680],
      [640, 660],
      [620, 640],
      [600, 620],
    ]
    return [
      h(100, 400, 700),
      h(100, 400, 680),
      h(100, 400, 600),
      ...xs.flatMap((x) => bands.map(([y0, y1]) => v(y0, y1, x))),
    ]
  }

  it('recovers unruled row boundaries where >=3 vlines break at the same y', () => {
    const rows: Array<[string, string, string, number]> = [
      ['Head1', 'Head2', 'Head3', 686],
      ['a', '1', '2', 666],
      ['b', '3', '4', 646],
      ['c', '5', '6', 626],
      ['d', '7', '8', 606],
    ]
    const chars = rows.flatMap(([c1, c2, c3, y]) => [
      ...mkText(c1, 110, { y }).chars,
      ...mkText(c2, 210, { y }).chars,
      ...mkText(c3, 310, { y }).chars,
    ])
    const { tables } = detectTables(shapesOf(perRowVSegments()), chars)
    expect(tables).toHaveLength(1)
    expect(tables[0]!.rows).toHaveLength(5)
    for (let r = 1; r < 5; r++) {
      const row = tables[0]!.rows[r]!
      const texts = row.map((c) =>
        c.blocks
          .map((b) => b.lines.map((l) => l.spans.map((s) => s.text).join('')).join(''))
          .join(''),
      )
      expect(texts).toHaveLength(3)
      expect(texts.every((t) => t.length > 0)).toBe(true)
    }
  })

  it('does not invent boundaries where text straddles the junction y', () => {
    // tall chars crossing y=660 refute the junction there
    const chars = [
      ...mkText('Head1', 110, { y: 686 }).chars,
      ...mkText('Head2', 210, { y: 686 }).chars,
      ...mkText('big', 110, { y: 655, fontSize: 20 }).chars,
    ]
    const { tables } = detectTables(shapesOf(perRowVSegments()), chars)
    expect(tables).toHaveLength(1)
    // y=660 junction refuted by the straddling glyphs; others may survive
    expect(tables[0]!.rows.some((row) => row[0] && row[0].box.y0 === 660)).toBe(false)
  })
})
