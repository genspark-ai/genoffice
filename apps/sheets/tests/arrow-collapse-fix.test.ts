/**
 * Arrow keys on a multi-cell selection must move from the ACTIVE cell, not
 * from the selection range's edge: K97:L97 with active K97 + ArrowRight is
 * L97 in Excel, not M97. collapseRange decides what the selection collapses
 * to before the move replays.
 */
import { describe, expect, it } from 'vitest'
import { RANGE_TYPE } from '@univerjs/core'
import { collapseRange } from '../src/renderer/arrow-collapse-fix'

const cell = (row: number, col: number) => ({
  startRow: row,
  endRow: row,
  startColumn: col,
  endColumn: col,
})

describe('collapseRange', () => {
  it('collapses a multi-cell selection to the active cell', () => {
    // K97:L97 with active K97 (rows/cols 0-indexed: row 96, cols 10..11)
    const range = { startRow: 96, endRow: 96, startColumn: 10, endColumn: 11 }
    expect(collapseRange(range, cell(96, 10))).toEqual({
      ...cell(96, 10),
      rangeType: RANGE_TYPE.NORMAL,
    })
  })

  it('collapses to the primary even when it anchors the far end', () => {
    const range = { startRow: 96, endRow: 96, startColumn: 10, endColumn: 11 }
    expect(collapseRange(range, cell(96, 11))).toEqual({
      ...cell(96, 11),
      rangeType: RANGE_TYPE.NORMAL,
    })
  })

  it('collapses a whole-row selection to the active cell', () => {
    const range = {
      startRow: 5,
      endRow: 5,
      startColumn: 0,
      endColumn: 25,
      rangeType: RANGE_TYPE.ROW,
    }
    expect(collapseRange(range, cell(5, 3))).toEqual({
      ...cell(5, 3),
      rangeType: RANGE_TYPE.NORMAL,
    })
  })

  it('keeps a merged active block whole', () => {
    // selection B2:D5, active cell is the merge B2:C3
    const range = { startRow: 1, endRow: 4, startColumn: 1, endColumn: 3 }
    const merge = { startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 }
    expect(collapseRange(range, merge)).toEqual({ ...merge, rangeType: RANGE_TYPE.NORMAL })
  })

  it('returns null when the selection already is the active block', () => {
    expect(collapseRange(cell(96, 10), cell(96, 10))).toBeNull()
    const merge = { startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 }
    expect(collapseRange(merge, merge)).toBeNull()
  })

  it('returns null without a primary', () => {
    const range = { startRow: 96, endRow: 96, startColumn: 10, endColumn: 11 }
    expect(collapseRange(range, null)).toBeNull()
    expect(collapseRange(range, undefined)).toBeNull()
  })
})
