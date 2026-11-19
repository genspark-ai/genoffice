import { describe, expect, it } from 'vitest'

import { hiddenRowsInFilterRange } from '../src/renderer/filtered-copy'

import type { IRange } from '@univerjs/core'

const range = (startRow: number, endRow: number): IRange => ({
  startRow,
  endRow,
  startColumn: 0,
  endColumn: 3,
})

describe('hiddenRowsInFilterRange', () => {
  const hidden = new Set([2, 4, 6, 20])
  const isHidden = (row: number): boolean => hidden.has(row)

  it('reports hidden rows inside the filter range', () => {
    expect(hiddenRowsInFilterRange(range(0, 11), range(0, 11), isHidden)).toEqual([2, 4, 6])
  })

  it('returns nothing without an active filter', () => {
    expect(hiddenRowsInFilterRange(range(0, 11), null, isHidden)).toEqual([])
  })

  it('keeps hidden rows outside the filter range (Excel manual-hide behavior)', () => {
    // Rows 20+ are hidden but the filter only covers 0..11, so row 20 copies.
    expect(hiddenRowsInFilterRange(range(0, 30), range(0, 11), isHidden)).toEqual([2, 4, 6])
  })

  it('intersects the copy range with the filter range', () => {
    expect(hiddenRowsInFilterRange(range(3, 5), range(0, 11), isHidden)).toEqual([4])
  })

  it('copies everything when nothing is hidden', () => {
    expect(hiddenRowsInFilterRange(range(0, 11), range(0, 11), () => false)).toEqual([])
  })
})
