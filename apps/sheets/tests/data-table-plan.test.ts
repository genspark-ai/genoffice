import { describe, expect, it } from 'vitest'

import { MAX_DATA_TABLE_CELLS, planDataTable } from '../src/renderer/data-table'

const cell = (row: number, column: number) => ({ row, column })
const bounds = (startRow: number, startColumn: number, endRow: number, endColumn: number) => ({
  startRow,
  startColumn,
  endRow,
  endColumn,
})

describe('planDataTable layout', () => {
  it('maps a column-oriented one-variable region: formula top-right, values down the left', () => {
    // B4:C7 → formula C4, values B5:B7, results C5:C7
    const plan = planDataTable(bounds(3, 1, 6, 2), null, cell(0, 3))
    expect(plan).toEqual({
      kind: 'column',
      formulaCell: cell(3, 2),
      rowValueCells: [],
      columnValueCells: [cell(4, 1), cell(5, 1), cell(6, 1)],
      resultStart: cell(4, 2),
      resultRows: 3,
      resultColumns: 1,
    })
  })

  it('maps a row-oriented one-variable region: formula bottom-left, values across the top', () => {
    // B4:D5 → formula B5, values C4:D4, results C5:D5
    const plan = planDataTable(bounds(3, 1, 4, 3), cell(8, 4), null)
    expect(plan).toEqual({
      kind: 'row',
      formulaCell: cell(4, 1),
      rowValueCells: [cell(3, 2), cell(3, 3)],
      columnValueCells: [],
      resultStart: cell(4, 2),
      resultRows: 1,
      resultColumns: 2,
    })
  })

  it('maps a two-variable region: formula corner, top row → row input, left column → column input', () => {
    // B4:D6 → formula B4, row values C4:D4, column values B5:B6, results C5:D6
    const plan = planDataTable(bounds(3, 1, 5, 3), cell(0, 5), cell(1, 5))
    expect(plan).toEqual({
      kind: 'grid',
      formulaCell: cell(3, 1),
      rowValueCells: [cell(3, 2), cell(3, 3)],
      columnValueCells: [cell(4, 1), cell(5, 1)],
      resultStart: cell(4, 2),
      resultRows: 2,
      resultColumns: 2,
    })
  })

  it('accepts the minimal 2×2 shapes for every kind', () => {
    expect(planDataTable(bounds(0, 0, 1, 1), null, cell(5, 5))).toMatchObject({
      kind: 'column',
      resultRows: 1,
      resultColumns: 1,
    })
    expect(planDataTable(bounds(0, 0, 1, 1), cell(5, 5), null)).toMatchObject({
      kind: 'row',
      resultRows: 1,
      resultColumns: 1,
    })
    expect(planDataTable(bounds(0, 0, 1, 1), cell(5, 5), cell(6, 6))).toMatchObject({
      kind: 'grid',
      resultRows: 1,
      resultColumns: 1,
    })
  })
})

describe('planDataTable validation', () => {
  it('rejects when no input cell is given', () => {
    expect(planDataTable(bounds(3, 1, 5, 3), null, null)).toBe('noInputCell')
  })

  it('rejects a two-variable table whose inputs are the same cell', () => {
    expect(planDataTable(bounds(3, 1, 5, 3), cell(0, 5), cell(0, 5))).toBe('sameInputCell')
  })

  it('rejects input cells inside the table region, whichever axis they feed', () => {
    expect(planDataTable(bounds(3, 1, 5, 3), cell(4, 2), null)).toBe('inputInsideTable')
    expect(planDataTable(bounds(3, 1, 5, 3), null, cell(3, 3))).toBe('inputInsideTable')
    expect(planDataTable(bounds(3, 1, 5, 3), cell(5, 1), cell(0, 5))).toBe('inputInsideTable')
  })

  it('rejects wrong shapes: 3 columns or 1 row for column tables, 3 rows or 1 column for row tables', () => {
    expect(planDataTable(bounds(3, 1, 5, 3), null, cell(0, 5))).toBe('badShape')
    expect(planDataTable(bounds(3, 1, 3, 2), null, cell(0, 5))).toBe('badShape')
    expect(planDataTable(bounds(3, 1, 5, 2), cell(0, 5), null)).toBe('badShape')
    expect(planDataTable(bounds(3, 1, 4, 1), cell(0, 5), null)).toBe('badShape')
    // grid needs both dimensions ≥ 2
    expect(planDataTable(bounds(3, 1, 3, 3), cell(0, 5), cell(1, 5))).toBe('badShape')
    expect(planDataTable(bounds(3, 1, 5, 1), cell(0, 5), cell(1, 5))).toBe('badShape')
  })

  it('caps the result area at MAX_DATA_TABLE_CELLS', () => {
    // column table of exactly MAX values passes, one more does not
    expect(planDataTable(bounds(0, 0, MAX_DATA_TABLE_CELLS, 1), null, cell(0, 5))).toMatchObject({
      kind: 'column',
      resultRows: MAX_DATA_TABLE_CELLS,
    })
    expect(planDataTable(bounds(0, 0, MAX_DATA_TABLE_CELLS + 1, 1), null, cell(0, 5))).toBe(
      'tooManyCells',
    )
    // grid: (n-1)² cells — 101×101 results overflows (inputs stay outside)
    expect(planDataTable(bounds(0, 0, 101, 101), cell(0, 200), cell(1, 200))).toBe('tooManyCells')
  })
})
