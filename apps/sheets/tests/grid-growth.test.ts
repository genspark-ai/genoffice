import { describe, expect, it } from 'vitest'

import { nextGridGrowth } from '../src/renderer/grid-growth'

/// A blank workbook's grid, and a viewport parked at A1.
const BLANK = { startColumn: 0, startRow: 0, columnCount: 26, rowCount: 1000 } as const

describe('nextGridGrowth', () => {
  it('leaves a blank sheet alone until the viewport nears an edge', () => {
    // column 0 + 20 columns of lookahead < 26, so opening a blank sheet must
    // not immediately grow anything
    expect(nextGridGrowth(BLANK)).toBeNull()
  })

  it('extends the columns when the viewport reaches the right edge', () => {
    // column 6 leaves 20 columns of lookahead, which is the edge
    expect(nextGridGrowth({ ...BLANK, startColumn: 6 })).toEqual({
      columnCount: 52,
      rowCount: 1000,
    })
  })

  it('extends the rows when the viewport reaches the bottom edge', () => {
    expect(nextGridGrowth({ ...BLANK, startRow: 800 })).toEqual({
      columnCount: 26,
      rowCount: 1200,
    })
  })

  it('grows only the axis that reached its edge', () => {
    const growth = nextGridGrowth({ ...BLANK, startColumn: 6 })
    expect(growth?.rowCount).toBe(1000)
    expect(growth?.columnCount).toBe(52)
  })

  it('is self-limiting: the size it returns needs no further growth', () => {
    // the scroll handler re-runs on every scroll state, including the one the
    // growth itself provokes. If the result still asked to grow, one scroll
    // would cascade into a runaway extension
    const first = nextGridGrowth({ ...BLANK, startColumn: 6 })!
    expect(
      nextGridGrowth({
        startColumn: 6,
        startRow: 0,
        columnCount: first.columnCount,
        rowCount: first.rowCount,
      }),
    ).toBeNull()
  })

  it('does not grow a grid that already fills the sheet', () => {
    expect(
      nextGridGrowth({
        startColumn: 16_300,
        startRow: 1_048_000,
        columnCount: 16_384,
        rowCount: 1_048_576,
      }),
    ).toBeNull()
  })

  it('stops at the .xlsx ceiling instead of growing past it', () => {
    // 100 columns short of the last column: one step overshoots, and the grid
    // must land exactly on 16384 rather than 16410
    const growth = nextGridGrowth({
      ...BLANK,
      startColumn: 16_364,
      columnCount: 16_384,
      rowCount: 1_048_576,
    })
    expect(growth).toBeNull()

    const near = nextGridGrowth({
      ...BLANK,
      startColumn: 16_350,
      columnCount: 16_370,
      rowCount: 1_048_576,
    })
    expect(near?.columnCount).toBe(16_384)
  })

  it('stops at the .xlsx row ceiling too', () => {
    // 76 rows short of the last row: one step of 200 overshoots, so the row
    // clamp has to land exactly on 1048576
    const near = nextGridGrowth({
      startColumn: 0,
      startRow: 1_048_400,
      columnCount: 26,
      rowCount: 1_048_500,
    })
    expect(near?.rowCount).toBe(1_048_576)
  })

  it('keeps scrolling right working past the original 26 columns', () => {
    // the reported symptom: a blank sheet stops at Z. Walk the viewport the
    // way a scrollbar drag does — the grid must extend ahead of the cursor
    let edge = { startColumn: 0, startRow: 0, columnCount: 26, rowCount: 1000 }
    for (let column = 0; column < 400; column += 5) {
      const growth = nextGridGrowth({ ...edge, startColumn: column })
      if (growth) {
        edge = { ...edge, columnCount: growth.columnCount, rowCount: growth.rowCount }
      }
    }
    expect(edge.columnCount).toBeGreaterThan(26)
  })
})
