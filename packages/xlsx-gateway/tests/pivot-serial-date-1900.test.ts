import { describe, expect, it } from 'vitest'

import { parseDateParts } from '../src/domain/pivot-grouping'
import { groupValue } from '../src/domain/pivot-grouping'
import { monthKeyOf } from '../src/domain/pivot-timeline'

describe('Excel 1900 serial dates in pivot date grouping', () => {
  it('reads serials below the phantom leap day at their true month', () => {
    // 1900-01-01 .. 1900-01-31
    expect(parseDateParts(1)).toEqual({ year: 1900, month: 1 })
    expect(parseDateParts(31)).toEqual({ year: 1900, month: 1 })
    // 1900-02-01 .. 1900-02-28
    expect(parseDateParts(32)).toEqual({ year: 1900, month: 2 })
    expect(parseDateParts(59)).toEqual({ year: 1900, month: 2 })
    // serial 60 is the non-existent 1900-02-29
    expect(parseDateParts(60)).toEqual({ year: 1900, month: 2 })
    // serial 61 is 1900-03-01, where the epoch already lines up
    expect(parseDateParts(61)).toEqual({ year: 1900, month: 3 })
    expect(parseDateParts(45292)).toEqual({ year: 2024, month: 1 })
  })

  it('ignores the time fraction when picking the month bucket', () => {
    expect(parseDateParts(1.75)).toEqual({ year: 1900, month: 1 })
    expect(parseDateParts(59.999)).toEqual({ year: 1900, month: 2 })
  })

  it('groups a January-1900 serial into Jan rather than Dec 1899', () => {
    expect(groupValue({ kind: 'date', dateUnit: 'month' }, 1)).toEqual({
      label: 'Jan',
      sort: 1,
    })
    expect(groupValue({ kind: 'date', dateUnit: 'year' }, 1)).toEqual({
      label: '1900',
      sort: 1900,
    })
  })

  it('builds timeline month keys that stay inside the true month', () => {
    expect(monthKeyOf(1)).toBe(1900 * 12)
    expect(monthKeyOf(59)).toBe(1900 * 12 + 1)
    expect(monthKeyOf(61)).toBe(1900 * 12 + 2)
  })
})
