import { describe, expect, it } from 'vitest'

import type { StructuralOp } from '../src/gateway/xlsx-structure'
import {
  fileRangeToScreenRange,
  fileToScreen,
  indexedThroughScreenRow,
  mapRangeResultToScreen,
  netAxisDelta,
  screenRangeToFileRange,
  screenToFile,
} from '../src/renderer/view-transform'

const insertRows = (index: number, count = 1): StructuralOp => ({
  kind: 'insert-rows',
  index,
  count,
})
const removeRows = (index: number, count = 1): StructuralOp => ({
  kind: 'remove-rows',
  index,
  count,
})
const insertCols = (index: number, count = 1): StructuralOp => ({
  kind: 'insert-cols',
  index,
  count,
})
const merge = (): StructuralOp => ({
  kind: 'merge-cells',
  range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
})

describe('fileToScreen / screenToFile', () => {
  it('shifts positions across inserts and removals and stays inverse-consistent', () => {
    const ops = [insertRows(5, 2), removeRows(10, 1), insertCols(2, 3)]
    // File row 4 is untouched; file row 5 moved down by the insert.
    expect(fileToScreen(ops, 'row', 4)).toBe(4)
    expect(fileToScreen(ops, 'row', 5)).toBe(7)
    // Screen rows 5-6 are the inserted lines — no file backing.
    expect(screenToFile(ops, 'row', 5)).toBeNull()
    expect(screenToFile(ops, 'row', 6)).toBeNull()
    expect(screenToFile(ops, 'row', 7)).toBe(5)
    // The removal happened at post-insert screen index 10 = file row 8.
    expect(fileToScreen(ops, 'row', 8)).toBeNull()
    expect(fileToScreen(ops, 'row', 9)).toBe(10)
    expect(screenToFile(ops, 'row', 10)).toBe(9)
    // Columns shift independently.
    expect(fileToScreen(ops, 'column', 1)).toBe(1)
    expect(fileToScreen(ops, 'column', 2)).toBe(5)
    expect(screenToFile(ops, 'column', 3)).toBeNull()
    // Round trip over surviving lines.
    for (let file = 0; file < 30; file += 1) {
      const screen = fileToScreen(ops, 'row', file)
      if (screen !== null) expect(screenToFile(ops, 'row', screen)).toBe(file)
    }
  })

  it('ignores merge operations and reports net axis deltas', () => {
    const ops = [merge(), insertRows(0, 4), removeRows(2, 1), merge()]
    expect(fileToScreen(ops, 'row', 0)).toBe(3)
    expect(netAxisDelta(ops, 'row')).toBe(3)
    expect(netAxisDelta(ops, 'column')).toBe(0)
  })
})

describe('screenRangeToFileRange', () => {
  const range = (startRow: number, endRow: number, startColumn = 0, endColumn = 5) => ({
    startRow,
    endRow,
    startColumn,
    endColumn,
  })

  it('translates a viewport past an insert back to file rows', () => {
    const ops = [insertRows(5, 2)]
    expect(screenRangeToFileRange(ops, range(10, 20))).toEqual(range(8, 18))
  })

  it('skips inserted screen lines at the edges of the request', () => {
    const ops = [insertRows(10, 3)]
    // Screen 10-12 are inserted; only screen 13-15 have file backing.
    expect(screenRangeToFileRange(ops, range(10, 15))).toEqual(range(10, 12))
  })

  it('returns null when the whole range is journal-owned', () => {
    const ops = [insertRows(0, 50)]
    expect(screenRangeToFileRange(ops, range(0, 49))).toBeNull()
  })

  it('spans deleted file rows so the request stays contiguous', () => {
    const ops = [removeRows(10, 5)]
    expect(screenRangeToFileRange(ops, range(8, 12))).toEqual(range(8, 17))
  })
})

describe('mapRangeResultToScreen', () => {
  it('shifts cells and drops content on deleted lines', () => {
    const ops = [removeRows(1, 1), insertCols(0, 1)]
    const mapped = mapRangeResultToScreen(ops, {
      cells: [
        { row: 0, column: 0, value: 'keep' },
        { row: 1, column: 0, value: 'deleted row' },
        { row: 2, column: 1, value: 'shifted' },
      ],
      rows: [
        { row: 1, hidden: true },
        { row: 2, height: 30, hidden: false },
      ],
      merges: [
        { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
        { startRow: 2, endRow: 3, startColumn: 0, endColumn: 0 },
      ],
      hyperlinks: [{ row: 2, column: 0, target: 'https://example.com' }],
      conditionalRules: [],
      autoFilter: null,
      dataValidations: [],
      sheetProtection: null,
      indexedThroughRow: null,
      indexingComplete: true,
    })
    expect(mapped.cells).toEqual([
      { row: 0, column: 1, value: 'keep' },
      { row: 1, column: 2, value: 'shifted' },
    ])
    expect(mapped.rows).toEqual([{ row: 1, height: 30, hidden: false }])
    // First merge touches the deleted row → skipped; second shifts.
    expect(mapped.merges).toEqual([{ startRow: 1, endRow: 2, startColumn: 1, endColumn: 1 }])
    expect(mapped.hyperlinks).toEqual([{ row: 1, column: 1, target: 'https://example.com' }])
  })
})

describe('indexedThroughScreenRow', () => {
  it('maps the cutoff through shifts and deleted rows', () => {
    expect(indexedThroughScreenRow([insertRows(0, 2)], 10)).toBe(12)
    // File rows 8-10 deleted: the cutoff falls back to file row 7 → screen 7.
    expect(indexedThroughScreenRow([removeRows(8, 3)], 10)).toBe(7)
    expect(indexedThroughScreenRow([], null)).toBeNull()
    // Everything up to the cutoff was deleted.
    expect(indexedThroughScreenRow([removeRows(0, 20)], 10)).toBe(-1)
  })
})

describe('move-rows mapping', () => {
  const move = (index: number, count: number, before: number) =>
    ({ kind: 'move-rows', index, count, before }) as const

  it('maps positions as a bijection in both directions', () => {
    const ops = [move(1, 2, 5)] // rows 1-2 relocate before row 5
    // Post-move screen order of file rows: 0, 3, 4, 1, 2, 5, …
    expect([0, 1, 2, 3, 4, 5].map((row) => fileToScreen(ops, 'row', row))).toEqual([
      0, 3, 4, 1, 2, 5,
    ])
    for (let row = 0; row < 8; row += 1) {
      const screen = fileToScreen(ops, 'row', row)
      expect(screen).not.toBeNull()
      expect(screenToFile(ops, 'row', screen ?? -1)).toBe(row)
    }
    expect(netAxisDelta(ops, 'row')).toBe(0)
  })

  it('composes with inserts and removals', () => {
    const ops = [insertRows(0, 1), move(2, 1, 5)]
    // File row 1 → screen 2 after insert, then moves before row 5 → screen 4.
    expect(fileToScreen(ops, 'row', 1)).toBe(4)
    expect(screenToFile(ops, 'row', 4)).toBe(1)
  })

  it('widens span mapping to an envelope across the swap', () => {
    const ops = [move(0, 1, 10)] // row 0 relocates to the end of 0-9
    // Screen rows 8-9 are file rows 9 and 0 — the file span must cover both.
    const range = screenRangeToFileRange(ops, {
      startRow: 8,
      endRow: 9,
      startColumn: 0,
      endColumn: 0,
    })
    expect(range).not.toBeNull()
    expect(range?.startRow).toBeLessThanOrEqual(0)
    expect(range?.endRow).toBeGreaterThanOrEqual(9)
  })

  it('covers the displaced-block start when widening file → screen', () => {
    const ops = [move(0, 1, 3)] // file→screen: 0→2, 1→0, 2→1
    // File row 1 (the displaced block start) lands on screen row 0 — the
    // envelope from the endpoints alone ([1,2]) would miss it.
    const range = fileRangeToScreenRange(ops, {
      startRow: 0,
      endRow: 2,
      startColumn: 0,
      endColumn: 0,
    })
    expect(range).not.toBeNull()
    expect(range?.startRow).toBeLessThanOrEqual(0)
    expect(range?.endRow).toBeGreaterThanOrEqual(2)
  })

  it('evaluates move blocks in post-insert coordinates (insert then move)', () => {
    const ops = [insertRows(0, 5), move(5, 2, 10)]
    // File rows 0-4 sit at screen 5-9 after the insert; the move then swaps
    // screen blocks 5-6 and 7-9, so their screen images are 5-9 in scrambled
    // order: 0→8, 1→9, 2→5, 3→6, 4→7. Sampling the move's block edges in file
    // space (5-9, all outside the input span 0-4) would miss every extreme.
    const screens = [0, 1, 2, 3, 4].map((row) => fileToScreen(ops, 'row', row))
    expect(screens).toEqual([8, 9, 5, 6, 7])
    const range = fileRangeToScreenRange(ops, {
      startRow: 0,
      endRow: 4,
      startColumn: 0,
      endColumn: 0,
    })
    expect(range?.startRow).toBeLessThanOrEqual(5)
    expect(range?.endRow).toBeGreaterThanOrEqual(9)
    // And the inverse: screen rows 5-9 must map back to a span covering 0-4.
    const back = screenRangeToFileRange(ops, {
      startRow: 5,
      endRow: 9,
      startColumn: 0,
      endColumn: 0,
    })
    expect(back?.startRow).toBeLessThanOrEqual(0)
    expect(back?.endRow).toBeGreaterThanOrEqual(4)
  })

  it('covers extremes under chained moves', () => {
    const ops = [move(0, 2, 8), move(3, 3, 9)]
    // Brute-force the true envelope of every screen sub-span and check both
    // range translations against it.
    for (let start = 0; start < 10; start += 1) {
      for (let end = start; end < 10; end += 1) {
        const files = []
        for (let screen = start; screen <= end; screen += 1) {
          const file = screenToFile(ops, 'row', screen)
          if (file !== null) files.push(file)
        }
        const range = screenRangeToFileRange(ops, {
          startRow: start,
          endRow: end,
          startColumn: 0,
          endColumn: 0,
        })
        expect(range?.startRow).toBeLessThanOrEqual(Math.min(...files))
        expect(range?.endRow).toBeGreaterThanOrEqual(Math.max(...files))
      }
    }
  })

  it('covers the post-move block boundary when widening screen → file', () => {
    const ops = [move(0, 2, 6)] // screen→file: 0→2, 1→3, 2→4, 3→5, 4→0, 5→1
    // The inverse map splits at screen row 4: rows 3 and 4 map to file rows 5
    // and 0, both outside the endpoint envelope ([1,3]).
    const range = screenRangeToFileRange(ops, {
      startRow: 1,
      endRow: 5,
      startColumn: 0,
      endColumn: 0,
    })
    expect(range).not.toBeNull()
    expect(range?.startRow).toBeLessThanOrEqual(0)
    expect(range?.endRow).toBeGreaterThanOrEqual(5)
  })
})
