/** Checkbox form-table detection unit tests: hand-built chars + strokes, no wasm. */
import { describe, expect, it } from 'vitest'
import { detectCheckboxSquares, detectFormTables } from '../src/analyze/form'
import { clusterCombiningMarks, groupIntoLines } from '../src/analyze/lines'
import { splitIntoUnits } from '../src/analyze/units'
import type { PageShapes, PdfChar, Stroke, TableBlock } from '../src/ir'
import { mkText } from './helpers/chars'

/** the 4 thin strokes of a drawn square outline (side pt, bottom-left x/y) */
function squareStrokes(x: number, y: number, side = 10): Stroke[] {
  const thin = 0.05
  const mk = (box: Stroke['box'], orientation: Stroke['orientation']): Stroke => ({
    box,
    orientation,
    widthPt: 0.1,
    color: '000000',
  })
  return [
    mk({ x0: x, x1: x + side, y0: y + side - thin, y1: y + side + thin }, 'h'),
    mk({ x0: x, x1: x + side, y0: y - thin, y1: y + thin }, 'h'),
    mk({ x0: x - thin, x1: x + thin, y0: y, y1: y + side }, 'v'),
    mk({ x0: x + side - thin, x1: x + side + thin, y0: y, y1: y + side }, 'v'),
  ]
}

const shapesOf = (strokes: Stroke[]): PageShapes => ({ strokes, fills: [], ignoredPaths: 0 })

const unitsOf = (chars: PdfChar[]) => splitIntoUnits(groupIntoLines(clusterCombiningMarks(chars)))

const cellText = (table: TableBlock, r: number, c: number): string =>
  table.rows[r]![c]!.blocks.map((b) =>
    b.lines.map((l) => l.spans.map((s) => s.text).join('')).join(' '),
  ).join(' ')

/**
 * The NIST-style control block: one checkbox row "☐ Implemented ☐ Planned
 * ☐ Not Applicable" (squares at x 64/229/395, baseline 700) plus a tight
 * full-width description line under it and a distant next heading.
 */
function formPage(): { chars: PdfChar[]; strokes: Stroke[] } {
  const strokes = [
    ...squareStrokes(64, 698),
    ...squareStrokes(229, 698),
    ...squareStrokes(395, 698),
  ]
  const chars = [
    ...mkText('Implemented', 80, { y: 700 }).chars,
    ...mkText('Planned to be Implemented', 245, { y: 700 }).chars,
    ...mkText('Not Applicable', 411, { y: 700 }).chars,
    ...mkText('Current implementation or planned implementation details.', 64, { y: 686 }).chars,
    ...mkText('3.1.16. Authorize wireless access prior to allowing.', 58, { y: 660 }).chars,
  ]
  return { chars, strokes }
}

describe('detectCheckboxSquares', () => {
  it('finds small drawn square outlines', () => {
    const squares = detectCheckboxSquares(squareStrokes(64, 698))
    expect(squares).toHaveLength(1)
    expect(squares[0]!.x0).toBeCloseTo(64, 0)
    expect(squares[0]!.x1).toBeCloseTo(74, 0)
  })

  it('rejects big rects, non-squares and lone rules', () => {
    expect(detectCheckboxSquares(squareStrokes(64, 600, 30))).toHaveLength(0)
    // 10×22 outline — not a square
    const tall = squareStrokes(64, 600).map((s) =>
      s.orientation === 'v'
        ? { ...s, box: { ...s.box, y1: s.box.y0 + 22 } }
        : s.box.y0 > 605
          ? { ...s, box: { x0: s.box.x0, x1: s.box.x1, y0: 622 - 0.05, y1: 622 + 0.05 } }
          : s,
    )
    expect(detectCheckboxSquares(tall)).toHaveLength(0)
    expect(detectCheckboxSquares([squareStrokes(64, 600)[0]!])).toHaveLength(0)
  })
})

describe('detectFormTables', () => {
  it('turns a checkbox row + tight description line into a 2-row table', () => {
    const { chars, strokes } = formPage()
    const { tables, remainingUnits } = detectFormTables(unitsOf(chars), shapesOf(strokes))
    expect(tables).toHaveLength(1)
    const table = tables[0]!
    expect(table.rows).toHaveLength(2)
    expect(table.rows[0]).toHaveLength(3)
    expect(cellText(table, 0, 0)).toBe('☐ Implemented')
    expect(cellText(table, 0, 1)).toBe('☐ Planned to be Implemented')
    expect(cellText(table, 0, 2)).toBe('☐ Not Applicable')
    // description joins as one merged full-width row
    expect(table.rows[1]).toHaveLength(1)
    expect(table.rows[1]![0]!.gridSpan).toBe(3)
    expect(cellText(table, 1, 0)).toContain('Current implementation')
    expect(table.confidence).toBeGreaterThanOrEqual(0.5)
    // the next heading stays in the flow (gap too large to be a description)
    expect(remainingUnits).toHaveLength(1)
    expect(remainingUnits[0]!.chars[0]!.text).toBe('3')
  })

  it('requires a checkbox on EVERY unit of the row', () => {
    const { chars } = formPage()
    const strokes = [...squareStrokes(64, 698), ...squareStrokes(229, 698)] // third missing
    const { tables, remainingUnits } = detectFormTables(unitsOf(chars), shapesOf(strokes))
    expect(tables).toHaveLength(0)
    expect(remainingUnits).toHaveLength(unitsOf(chars).length)
  })

  it('a single checkbox-led unit is not a table', () => {
    const chars = mkText('Implemented', 80, { y: 700 }).chars
    const { tables } = detectFormTables(unitsOf(chars), shapesOf(squareStrokes(64, 698)))
    expect(tables).toHaveLength(0)
  })

  it('stacks consecutive same-width checkbox rows into one grid', () => {
    const strokes = [
      ...squareStrokes(64, 698),
      ...squareStrokes(229, 698),
      ...squareStrokes(64, 682),
      ...squareStrokes(229, 682),
    ]
    const chars = [
      ...mkText('Alpha', 80, { y: 700 }).chars,
      ...mkText('Beta', 245, { y: 700 }).chars,
      ...mkText('Gamma', 80, { y: 684 }).chars,
      ...mkText('Delta', 245, { y: 684 }).chars,
    ]
    const { tables } = detectFormTables(unitsOf(chars), shapesOf(strokes))
    expect(tables).toHaveLength(1)
    expect(tables[0]!.rows).toHaveLength(2)
    expect(tables[0]!.rows[1]).toHaveLength(2)
    expect(cellText(tables[0]!, 1, 0)).toBe('☐ Gamma')
  })

  it('ignores squares inside excluded (lattice) regions', () => {
    const { chars, strokes } = formPage()
    const { tables } = detectFormTables(unitsOf(chars), shapesOf(strokes), [
      { x0: 0, y0: 0, x1: 612, y1: 792 },
    ])
    expect(tables).toHaveLength(0)
  })
})
