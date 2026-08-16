import { describe, expect, it } from 'vitest'

import { isCalendarDatePattern } from '../src/renderer/numfmt-fix'

describe('isCalendarDatePattern', () => {
  it('matches calendar date and datetime patterns', () => {
    expect(isCalendarDatePattern('m/d/yyyy')).toBe(true)
    expect(isCalendarDatePattern('yyyy-mm-dd h:mm')).toBe(true)
    expect(isCalendarDatePattern('d-mmm-yy')).toBe(true)
  })

  it('rejects time-only and elapsed patterns whose magnitude must not shift', () => {
    expect(isCalendarDatePattern('[h]:mm')).toBe(false)
    expect(isCalendarDatePattern('[mm]:ss')).toBe(false)
    expect(isCalendarDatePattern('h:mm:ss')).toBe(false)
    expect(isCalendarDatePattern('[h]:mm:ss;@')).toBe(false)
  })

  it('rejects plain number and text patterns', () => {
    expect(isCalendarDatePattern('#,##0.00')).toBe(false)
    expect(isCalendarDatePattern('@')).toBe(false)
    expect(isCalendarDatePattern('General')).toBe(false)
  })
})
