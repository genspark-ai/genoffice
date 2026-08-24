import { describe, expect, it } from 'vitest'

import {
  boundsToA1,
  clampBoundsToExtent,
  columnScopeHeaders,
  resolveScopeChip,
} from '../src/renderer/ai/selection-scope'

const LIVE = { a1: 'D2:D80', sheetId: 'sh1' }
const RUN = { a1: 'B2:B50', sheetId: 'sh1' }

describe('resolveScopeChip', () => {
  it('offers the live selection while no run owns a scope', () => {
    expect(resolveScopeChip(undefined, LIVE, false)).toEqual({ range: 'D2:D80', locked: false })
  })

  it('shows nothing once the user dismisses the chip', () => {
    expect(resolveScopeChip(undefined, LIVE, true)).toEqual({ range: null, locked: false })
  })

  it('shows nothing when the resting selection is a single cell', () => {
    expect(resolveScopeChip(undefined, null, false)).toEqual({ range: null, locked: false })
  })

  it("names the run's frozen range, not the selection the user moved to mid-run", () => {
    expect(resolveScopeChip(RUN, LIVE, false)).toEqual({ range: 'B2:B50', locked: true })
  })

  it('keeps naming the frozen range even after a mid-run dismiss, which the run cannot see', () => {
    expect(resolveScopeChip(RUN, LIVE, true)).toEqual({ range: 'B2:B50', locked: true })
  })

  it('shows no chip for a run the user had already scoped to the whole sheet', () => {
    expect(resolveScopeChip(null, LIVE, false)).toEqual({ range: null, locked: true })
  })

  it('carries the column names through so the chip can label itself by them', () => {
    const scope = { a1: 'B1:B417', sheetId: 'sh1', columns: ['Amount'] }
    expect(resolveScopeChip(undefined, scope, false)).toEqual({
      range: 'B1:B417',
      locked: false,
      columns: ['Amount'],
    })
  })
})

/** A sheet with a header row and 416 rows of data in columns A..D. */
const EXTENT = { lastRow: 416, lastColumn: 3 }
/** What clicking the column-B header hands us: the sheet's full height. */
const WHOLE_COLUMN_B = { startRow: 0, startColumn: 1, endRow: 1048575, endColumn: 1 }

const HEADERS = ['Order', 'Amount', 'Qty', 'Region']
const headerAt = (column: number): string => HEADERS[column] ?? ''

describe('clampBoundsToExtent', () => {
  it('caps a whole-column click at the last row holding data', () => {
    expect(clampBoundsToExtent(WHOLE_COLUMN_B, EXTENT)).toEqual({
      startRow: 0,
      startColumn: 1,
      endRow: 416,
      endColumn: 1,
    })
  })

  it('leaves a block inside the data untouched', () => {
    const bounds = { startRow: 4, startColumn: 0, endRow: 9, endColumn: 2 }
    expect(clampBoundsToExtent(bounds, EXTENT)).toEqual(bounds)
  })

  it('caps a whole-row click at the last column holding data', () => {
    const bounds = { startRow: 4, startColumn: 0, endRow: 4, endColumn: 16383 }
    expect(clampBoundsToExtent(bounds, EXTENT)).toEqual({
      startRow: 4,
      startColumn: 0,
      endRow: 4,
      endColumn: 3,
    })
  })

  it('still caps the height of a whole-column click on a column past the data', () => {
    // the column axis cannot be capped here, which must not cost us the row one
    expect(
      clampBoundsToExtent({ ...WHOLE_COLUMN_B, startColumn: 5, endColumn: 5 }, EXTENT),
    ).toEqual({ startRow: 0, startColumn: 5, endRow: 416, endColumn: 5 })
  })

  it('passes through a block the user marked out entirely past the data', () => {
    // an empty spot picked on purpose stays the size it was drawn
    const bounds = { startRow: 900, startColumn: 8, endRow: 905, endColumn: 9 }
    expect(clampBoundsToExtent(bounds, EXTENT)).toEqual(bounds)
  })

  it('passes through on a sheet with no data at all', () => {
    const bounds = { startRow: 0, startColumn: 0, endRow: 10, endColumn: 2 }
    expect(clampBoundsToExtent(bounds, { lastRow: -1, lastColumn: -1 })).toEqual(bounds)
  })
})

describe('boundsToA1', () => {
  it('writes a range', () => {
    expect(boundsToA1({ startRow: 0, startColumn: 1, endRow: 416, endColumn: 1 })).toBe('B1:B417')
  })

  it('collapses a single cell', () => {
    expect(boundsToA1({ startRow: 4, startColumn: 27, endRow: 4, endColumn: 27 })).toBe('AB5')
  })
})

describe('columnScopeHeaders', () => {
  const clamped = (bounds: typeof WHOLE_COLUMN_B) => clampBoundsToExtent(bounds, EXTENT)

  it('names a whole-column click by its header', () => {
    expect(columnScopeHeaders(clamped(WHOLE_COLUMN_B), EXTENT, headerAt)).toEqual(['Amount'])
  })

  it('names every column of a multi-column click', () => {
    expect(
      columnScopeHeaders(clamped({ ...WHOLE_COLUMN_B, endColumn: 2 }), EXTENT, headerAt),
    ).toEqual(['Amount', 'Qty'])
  })

  it('names a drag that happens to cover the full data height', () => {
    const bounds = { startRow: 0, startColumn: 1, endRow: 416, endColumn: 1 }
    expect(columnScopeHeaders(bounds, EXTENT, headerAt)).toEqual(['Amount'])
  })

  it('declines a block that does not reach the last data row', () => {
    const bounds = { startRow: 0, startColumn: 1, endRow: 40, endColumn: 1 }
    expect(columnScopeHeaders(bounds, EXTENT, headerAt)).toBeNull()
  })

  it('declines a block that starts below the header row', () => {
    const bounds = { startRow: 1, startColumn: 1, endRow: 416, endColumn: 1 }
    expect(columnScopeHeaders(bounds, EXTENT, headerAt)).toBeNull()
  })

  it('declines when a covered column has no header text', () => {
    const bounds = clamped({ ...WHOLE_COLUMN_B, endColumn: 2 })
    expect(columnScopeHeaders(bounds, EXTENT, (column) => (column === 2 ? '   ' : 'Amount'))).toBe(
      null,
    )
  })

  it('declines once too many columns are covered to name them', () => {
    const bounds = clamped({ ...WHOLE_COLUMN_B, startColumn: 0, endColumn: 3 })
    expect(columnScopeHeaders(bounds, EXTENT, headerAt)).toBeNull()
  })

  it('declines on a sheet with no data', () => {
    const empty = { lastRow: -1, lastColumn: -1 }
    expect(columnScopeHeaders(WHOLE_COLUMN_B, empty, headerAt)).toBeNull()
  })

  it('ellipsizes a header too long for the chip', () => {
    const long = 'Total revenue recognised in the period'
    expect(columnScopeHeaders(clamped(WHOLE_COLUMN_B), EXTENT, () => long)).toEqual([
      'Total revenue recognised…',
    ])
  })
})
