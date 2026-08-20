/**
 * TableBlock → pptx grid-table spec (P25). The IR grid (colWidthsPt +
 * row-major cells with gridSpan / vMerge placeholders) maps 1:1 onto
 * buildTableGridXml's model; the only derived data is per-row heights (from
 * cell boxes) and rowSpan counts (from vMerge continue chains).
 */
import type { Rect } from '../geometry'
import { rectUnion } from '../geometry'
import type { TableBlock, TableCellBlock } from '../ir'
import type { NewTableCellSpec, NewTableGridOptions } from '../../../pptx-engine/src/insert'
import { textBlockParagraph } from './text'

/** lattice border weight (pt): the IR keeps no per-table stroke width */
const LATTICE_BORDER_PT = 0.75

export interface PageMapper {
  /** pt rect (PDF y-up) → EMU rect (slide top-left) */
  rect(box: Rect): { x: number; y: number; cx: number; cy: number }
  /** pt length → EMU */
  len(pt: number): number
  /** page→slide uniform scale (fonts ride it) */
  scale: number
}

/** column start index of every cell in one IR row (gridSpan advances) */
function colStarts(row: TableCellBlock[]): number[] {
  const starts: number[] = []
  let col = 0
  for (const cell of row) {
    starts.push(col)
    col += Math.max(1, cell.gridSpan)
  }
  return starts
}

/** rows a vMerge-restart cell spans: 1 + following continue placeholders in its column */
function rowSpanOf(block: TableBlock, rowIdx: number, colStart: number): number {
  let span = 1
  for (let r = rowIdx + 1; r < block.rows.length; r++) {
    const row = block.rows[r]!
    const starts = colStarts(row)
    const idx = starts.indexOf(colStart)
    if (idx < 0 || row[idx]!.vMerge !== 'continue') break
    span++
  }
  return span
}

/** per-row heights from row-top boundaries (first non-continue cell's box top) */
function rowHeightsPt(block: TableBlock): number[] {
  const tops: number[] = []
  for (const [r, row] of block.rows.entries()) {
    const anchor = row.find((c) => c.vMerge !== 'continue') ?? row[0]
    tops.push(anchor ? anchor.box.y1 : block.box.y1 - r)
  }
  const heights: number[] = []
  for (let r = 0; r < tops.length; r++) {
    const bottom = r + 1 < tops.length ? tops[r + 1]! : block.box.y0
    heights.push(Math.max(1, tops[r]! - bottom))
  }
  return heights
}

function cellSpec(
  block: TableBlock,
  rowIdx: number,
  cell: TableCellBlock,
  colStart: number,
  m: PageMapper,
): NewTableCellSpec[] {
  const span = Math.max(1, cell.gridSpan)
  if (cell.vMerge === 'continue') {
    // covered by the restart above; its trailing gridSpan columns are covered both ways
    const specs: NewTableCellSpec[] = [{ vMerge: true }]
    for (let i = 1; i < span; i++) specs.push({ vMerge: true, hMerge: true })
    return specs
  }
  const spec: NewTableCellSpec = {}
  if (span > 1) spec.gridSpan = span
  if (cell.vMerge === 'restart') {
    const rows = rowSpanOf(block, rowIdx, colStart)
    if (rows > 1) spec.rowSpan = rows
  }
  if (cell.fill) spec.fillColor = `#${cell.fill}`
  if (cell.vAlign) spec.anchor = cell.vAlign === 'center' ? 'ctr' : 'b'
  if (cell.blocks.length > 0) {
    spec.paragraphs = cell.blocks.map((b) => textBlockParagraph(b, m.scale))
    // measured content insets: the gap between the cell box and its text ink;
    // vertical margins stay 0 when an anchor already places the content
    const ink = cell.blocks.map((b) => b.box).reduce(rectUnion)
    const vAnchored = spec.anchor !== undefined
    spec.marginsEmu = {
      l: m.len(Math.max(0, ink.x0 - cell.box.x0)),
      t: vAnchored ? 0 : m.len(Math.max(0, cell.box.y1 - ink.y1)),
      r: m.len(Math.max(0, cell.box.x1 - ink.x1)),
      b: vAnchored ? 0 : m.len(Math.max(0, ink.y0 - cell.box.y0)),
    }
  } else {
    spec.marginsEmu = { l: 0, t: 0, r: 0, b: 0 }
  }
  const specs: NewTableCellSpec[] = [spec]
  for (let i = 1; i < span; i++) specs.push({ hMerge: true })
  return specs
}

export function tableGridOptions(block: TableBlock, m: PageMapper): NewTableGridOptions {
  const cells: NewTableCellSpec[][] = block.rows.map((row, r) => {
    const starts = colStarts(row)
    return row.flatMap((cell, i) => cellSpec(block, r, cell, starts[i]!, m))
  })
  const opts: NewTableGridOptions = {
    offset: m.rect(block.box),
    colWidthsEmu: block.colWidthsPt.map((w) => m.len(w)),
    rowHeightsEmu: rowHeightsPt(block).map((h) => m.len(h)),
    cells,
  }
  if (block.confidence === undefined) {
    // lattice: drawn rulings, unconditionally trusted → full grid borders
    opts.border = {
      color: `#${block.borderColor ?? '000000'}`,
      widthEmu: m.len(LATTICE_BORDER_PT),
      scope: 'all',
    }
  } else if (block.sepRule) {
    // rule-separated side-by-side zone (P22): only the drawn separator survives
    opts.border = {
      color: `#${block.borderColor ?? '000000'}`,
      widthEmu: m.len(1),
      scope: 'insideV',
    }
  }
  return opts
}
