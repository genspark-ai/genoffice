/**
 * A table with a live autoFilter re-ranks its row stripes by VISIBLE row
 * order (Excel ref: a filtered TableStyleLight19 alternates across the hidden
 * gaps; physical parity painted every visible row the same color). Tables
 * without live criteria — including manually hidden rows — keep physical
 * banding.
 */
import type { ICellData } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { applyTableBanding } from '../src/renderer/univer-sync'

const STRIPE = '#EDEDE3'

function stripedRows(options: {
  filterActive?: boolean
  hiddenRows?: ReadonlySet<number>
  coveredThrough?: number
}): (string | undefined)[] {
  const range = { startRow: 0, startColumn: 0, endRow: 9, endColumn: 1 }
  const matrix: ICellData[][] = Array.from({ length: 10 }, () =>
    Array.from({ length: 2 }, () => ({})),
  )
  applyTableBanding(
    matrix,
    range,
    [
      {
        range: { startRow: 0, startColumn: 0, endRow: 9, endColumn: 1 },
        headerRowCount: 1,
        showRowStripes: true,
        showColumnStripes: false,
        styleName: 'TableStyleLight19',
        stripeFill: STRIPE,
        ...(options.filterActive ? { filterActive: true } : {}),
      },
    ] as never,
    options.hiddenRows
      ? { rows: options.hiddenRows, coveredThrough: options.coveredThrough ?? 9 }
      : undefined,
  )
  return matrix.map((row) => (row[0]?.s as { bg?: { rgb?: string } } | undefined)?.bg?.rgb)
}

describe('applyTableBanding under a live filter', () => {
  it('ranks stripes by visible order across filter-hidden rows', () => {
    // Data rows 1-9; rows 2-4 filtered out → visible order 1, 5, 6, 7...
    const fills = stripedRows({ filterActive: true, hiddenRows: new Set([2, 3, 4]) })
    expect(fills[1]).toBe(STRIPE) // 1st visible
    expect(fills[5]).toBeUndefined() // 2nd visible
    expect(fills[6]).toBe(STRIPE) // 3rd visible
    expect(fills[7]).toBeUndefined()
  })

  it('keeps physical parity when rows are hidden without filter criteria', () => {
    const fills = stripedRows({ hiddenRows: new Set([2, 3, 4]) })
    expect(fills[1]).toBe(STRIPE)
    expect(fills[5]).toBe(STRIPE) // (5-1) even → physical firstRowStripe
    expect(fills[6]).toBeUndefined()
  })

  it('keeps physical parity when the filter shows every row', () => {
    const fills = stripedRows({ filterActive: true })
    expect(fills[1]).toBe(STRIPE)
    expect(fills[2]).toBeUndefined()
    expect(fills[3]).toBe(STRIPE)
  })

  it('keeps physical parity while row coverage stops short of the table end', () => {
    // Rows past coveredThrough are unknown, not visible — no re-rank yet.
    const fills = stripedRows({
      filterActive: true,
      hiddenRows: new Set([2, 3, 4]),
      coveredThrough: 5,
    })
    expect(fills[1]).toBe(STRIPE)
    expect(fills[5]).toBe(STRIPE) // physical parity, not visible ordinal
  })
})
