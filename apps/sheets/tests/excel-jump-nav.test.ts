/**
 * Ctrl/Cmd+Arrow data-edge jumps must use Excel's emptiness rule: a cell
 * with a formula (even one whose computed value is an empty string, or one
 * whose value hasn't been materialized in cache mode) is an entry the jump
 * stops at. Upstream Univer only looks at the computed display value.
 */
import { describe, expect, it } from 'vitest'
import { Direction, Styles, Worksheet } from '@univerjs/core'
import type { ICellData, IWorksheetData } from '@univerjs/core'
import { _internals } from '../src/renderer/excel-jump-nav'

const { excelCellHasValue, findNextGapRange } = _internals

function makeSheet(cellData: Record<number, Record<number, ICellData>>): Worksheet {
  const snapshot: Partial<IWorksheetData> = {
    id: 'test-sheet',
    name: 'Test',
    rowCount: 1000,
    columnCount: 26,
    cellData,
    mergeData: [],
    rowData: {},
    columnData: {},
    defaultRowHeight: 19,
    defaultColumnWidth: 73,
  }
  return new Worksheet('test-unit', snapshot as IWorksheetData, new Styles())
}

const cellAt = (row: number) => ({
  startRow: row,
  endRow: row,
  startColumn: 0,
  endColumn: 0,
})

describe('excelCellHasValue', () => {
  it('counts values, rich text, formulas, and shared formulas as entries', () => {
    expect(excelCellHasValue({ v: 42 })).toBe(true)
    expect(excelCellHasValue({ v: '' })).toBe(true)
    expect(excelCellHasValue({ v: 0 })).toBe(true)
    expect(excelCellHasValue({ f: '=IF(A1,"",B1)' })).toBe(true)
    expect(excelCellHasValue({ si: 'shared-1' } as ICellData)).toBe(true)
    expect(excelCellHasValue({ p: {} as ICellData['p'] })).toBe(true)
    expect(excelCellHasValue({})).toBe(false)
    expect(excelCellHasValue(null)).toBe(false)
    expect(excelCellHasValue({ s: 'style-only' } as ICellData)).toBe(false)
  })
})

describe('findNextGapRange (DOWN)', () => {
  it('walks block edge → next block start → isolated cell → sheet end', () => {
    const sheet = makeSheet({
      0: { 0: { v: 10 } },
      1: { 0: { v: 20 } },
      2: { 0: { v: 30 } },
      6: { 0: { v: 70 } },
      7: { 0: { v: 80 } },
      19: { 0: { v: 999 } },
    })
    let range = findNextGapRange(cellAt(0), Direction.DOWN, sheet)
    expect(range.startRow).toBe(2)
    range = findNextGapRange(cellAt(2), Direction.DOWN, sheet)
    expect(range.startRow).toBe(6)
    range = findNextGapRange(cellAt(7), Direction.DOWN, sheet)
    expect(range.startRow).toBe(19)
    range = findNextGapRange(cellAt(19), Direction.DOWN, sheet)
    expect(range.startRow).toBe(999)
  })

  it('stops at a formula block whose values are empty strings (r-genoffice-6#1)', () => {
    const sheet = makeSheet({
      0: { 0: { v: 10 } },
      6: { 0: { f: '=IF(B6,"","")', v: '' } },
      7: { 0: { f: '=IF(B7,"","")', v: '' } },
      12: { 0: { v: 99 } },
    })
    const range = findNextGapRange(cellAt(0), Direction.DOWN, sheet)
    expect(range.startRow).toBe(6)
  })

  it('stops at un-materialized formula cells (cache mode: f present, no v)', () => {
    const sheet = makeSheet({
      0: { 0: { v: 10 } },
      6: { 0: { f: '=SUM(B1:B9)' } },
      7: { 0: { si: 'shared-0' } as ICellData },
    })
    let range = findNextGapRange(cellAt(0), Direction.DOWN, sheet)
    expect(range.startRow).toBe(6)
    range = findNextGapRange(cellAt(6), Direction.DOWN, sheet)
    expect(range.startRow).toBe(7)
  })

  it('style-only cells are still gaps', () => {
    const sheet = makeSheet({
      0: { 0: { v: 10 } },
      6: { 0: { s: 'style-1' } as ICellData },
      12: { 0: { v: 99 } },
    })
    const range = findNextGapRange(cellAt(0), Direction.DOWN, sheet)
    expect(range.startRow).toBe(12)
  })
})

describe('findNextGapRange (UP/LEFT/RIGHT)', () => {
  it('mirrors the same semantics in the other directions', () => {
    const sheet = makeSheet({
      0: { 0: { v: 1 }, 3: { f: '=""' }, 8: { v: 9 } },
      5: { 0: { v: 5 } },
      9: { 0: { f: '=""' } },
    })
    // UP from row 9: formula cell at 9 is an entry; next stop is 5
    let range = findNextGapRange(cellAt(9), Direction.UP, sheet)
    expect(range.startRow).toBe(5)
    // RIGHT from A1: jump to the ="" formula cell at column 3
    range = findNextGapRange(
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      Direction.RIGHT,
      sheet,
    )
    expect(range.startColumn).toBe(3)
    // LEFT from I1 (col 8): back to the formula cell at column 3
    range = findNextGapRange(
      { startRow: 0, endRow: 0, startColumn: 8, endColumn: 8 },
      Direction.LEFT,
      sheet,
    )
    expect(range.startColumn).toBe(3)
  })
})
