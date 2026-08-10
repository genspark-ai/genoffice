import { describe, expect, it } from 'vitest'

import { expandAutoHeightRanges } from '../src/renderer/autofit-multi-row'

const COLUMNS = 26
const fullRows = (startRow: number, endRow: number) => ({
  startRow,
  endRow,
  startColumn: 0,
  endColumn: COLUMNS - 1,
})
const clicked = (row: number) => [{ startRow: row, endRow: row, startColumn: 0, endColumn: 5 }]

describe('expandAutoHeightRanges', () => {
  it('grows a single-row autofit to the full multi-row selection', () => {
    expect(expandAutoHeightRanges(clicked(3), [fullRows(1, 5)], COLUMNS)).toEqual([
      { startRow: 1, endRow: 5, startColumn: 0, endColumn: 5 },
    ])
  })

  it('covers Ctrl-selected disjoint row spans', () => {
    expect(expandAutoHeightRanges(clicked(2), [fullRows(1, 2), fullRows(7, 9)], COLUMNS)).toEqual([
      { startRow: 1, endRow: 2, startColumn: 0, endColumn: 5 },
      { startRow: 7, endRow: 9, startColumn: 0, endColumn: 5 },
    ])
  })

  it('leaves the command alone when the clicked row is outside the selection', () => {
    expect(expandAutoHeightRanges(clicked(10), [fullRows(1, 5)], COLUMNS)).toBeNull()
  })

  it('leaves single-row selections and cell-range selections alone', () => {
    expect(expandAutoHeightRanges(clicked(3), [fullRows(3, 3)], COLUMNS)).toBeNull()
    const cellRange = { startRow: 1, endRow: 5, startColumn: 2, endColumn: 4 }
    expect(expandAutoHeightRanges(clicked(3), [cellRange], COLUMNS)).toBeNull()
  })

  it('leaves multi-range or multi-row commands alone (already expanded)', () => {
    const twoRanges = [...clicked(1), ...clicked(2)]
    expect(expandAutoHeightRanges(twoRanges, [fullRows(1, 5)], COLUMNS)).toBeNull()
    const multiRow = [{ startRow: 1, endRow: 5, startColumn: 0, endColumn: 5 }]
    expect(expandAutoHeightRanges(multiRow, [fullRows(1, 5)], COLUMNS)).toBeNull()
  })
})
