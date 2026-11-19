import { describe, expect, it } from 'vitest'
import { shiftMarksThroughOps, opsWouldMoveMarks } from '../src/renderer/redact-marks-shift'
import type { StructuralJournalOp } from '../src/renderer/edit-journal'
import type { CellMark } from '../src/renderer/ai/redact'

/**
 * A mark is an absolute rectangle, so a row inserted above it leaves it
 * describing whatever shifted down into its place. The app looks correct — the
 * tint is drawn from the mark, the reader sees a marked cell — and the leak
 * only shows up in the saved file, on the next open, where the withheld value
 * sits outside the rectangle that was supposed to cover it.
 *
 * So these are about the coordinates, not the tint.
 */

const mark = (over: Partial<CellMark> = {}): CellMark => ({
  startRow: 2,
  endRow: 2,
  startColumn: 1,
  endColumn: 1,
  label: 'client phone',
  ...over,
})

const insertRows = (index: number, count = 1): StructuralJournalOp => ({
  kind: 'insert-rows',
  index,
  count,
})
const removeRows = (index: number, count = 1): StructuralJournalOp => ({
  kind: 'remove-rows',
  index,
  count,
})
const insertCols = (index: number, count = 1): StructuralJournalOp => ({
  kind: 'insert-cols',
  index,
  count,
})

describe('a mark after a row is inserted above it', () => {
  it('moves down with the value it covers', () => {
    const out = shiftMarksThroughOps([mark()], [insertRows(0)])
    expect(out[0]).toMatchObject({ startRow: 3, endRow: 3 })
  })

  it('stays put when the insert is below it', () => {
    const out = shiftMarksThroughOps([mark()], [insertRows(5)])
    expect(out[0]).toMatchObject({ startRow: 2, endRow: 2 })
  })

  it('accounts for several inserts in order', () => {
    const out = shiftMarksThroughOps([mark()], [insertRows(0), insertRows(0, 2), insertRows(1)])
    expect(out[0]).toMatchObject({ startRow: 6, endRow: 6 })
  })

  it('leaves the column alone', () => {
    const out = shiftMarksThroughOps([mark()], [insertRows(0)])
    expect(out[0]).toMatchObject({ startColumn: 1, endColumn: 1 })
  })
})

describe('a mark after a row is removed', () => {
  it('moves up when the removal is above it', () => {
    const out = shiftMarksThroughOps([mark()], [removeRows(0)])
    expect(out[0]).toMatchObject({ startRow: 1, endRow: 1 })
  })

  it('is dropped when the removal takes the value with it', () => {
    // the withheld value is gone; a rectangle clamped onto a neighbour would
    // withhold something the reader never hid
    expect(shiftMarksThroughOps([mark()], [removeRows(2)])).toEqual([])
  })

  it('is dropped when the removal clips it', () => {
    expect(shiftMarksThroughOps([mark()], [removeRows(1, 3)])).toEqual([])
  })

  it('survives a removal of other rows entirely', () => {
    const out = shiftMarksThroughOps([mark()], [removeRows(8)])
    expect(out).toHaveLength(1)
  })
})

describe('columns', () => {
  it('shift on the column axis and leave rows alone', () => {
    const out = shiftMarksThroughOps([mark()], [insertCols(0)])
    expect(out[0]).toMatchObject({
      startColumn: 2,
      endColumn: 2,
      startRow: 2,
      endRow: 2,
    })
  })
})

describe('shapes it does not touch', () => {
  it('returns the marks unchanged when nothing structural happened', () => {
    const marks = [mark(), mark({ startRow: 9 })]
    const out = shiftMarksThroughOps(marks, [])
    expect(out).toEqual(marks)
    // and it is a copy, not the same array
    expect(out).not.toBe(marks)
  })

  it('ignores merges, which move nothing', () => {
    const ops: StructuralJournalOp[] = [
      { kind: 'merge-cells', range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 } },
    ]
    expect(shiftMarksThroughOps([mark()], ops)).toEqual([mark()])
  })

  it('carries the label through untouched', () => {
    const out = shiftMarksThroughOps([mark()], [insertRows(0)])
    expect(out[0]!.label).toBe('client phone')
  })
})

describe('the question a guard would ask', () => {
  it('says no when an insert would move a mark', () => {
    expect(opsWouldMoveMarks([mark()], [insertRows(0)])).toBe(true)
    expect(opsWouldMoveMarks([mark()], [insertRows(5)])).toBe(false)
  })

  it('says no when a removal would drop one', () => {
    expect(opsWouldMoveMarks([mark()], [removeRows(2)])).toBe(true)
  })

  it('says yes for a workbook with nothing withheld', () => {
    expect(opsWouldMoveMarks([], [insertRows(0)])).toBe(false)
  })
})
