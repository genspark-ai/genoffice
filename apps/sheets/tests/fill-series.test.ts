import { describe, expect, it } from 'vitest'

import {
  type FillSeriesOptions,
  fillSeriesLine,
  parseSeriesNumber,
  planFillSeries,
  stepDate,
} from '../src/renderer/fill-series'

const base: FillSeriesOptions = {
  direction: 'columns',
  type: 'linear',
  dateUnit: 'day',
  trend: false,
  step: 1,
  stop: null,
}

// 2024-01-31 in the 1900 serial system
const JAN_31_2024 = 45322

describe('fillSeriesLine', () => {
  it('linear series adds the step from the first seed only', () => {
    expect(fillSeriesLine([5, 100], 4, { ...base, step: 2 })).toEqual([5, 7, 9, 11])
  })

  it('growth series multiplies by the step', () => {
    expect(fillSeriesLine([2], 4, { ...base, type: 'growth', step: 3 })).toEqual([2, 6, 18, 54])
  })

  it('stop value cuts the series before the first value past it', () => {
    expect(fillSeriesLine([1], 10, { ...base, step: 3, stop: 8 })).toEqual([1, 4, 7])
    expect(fillSeriesLine([10], 10, { ...base, step: -4, stop: 0 })).toEqual([10, 6, 2])
    expect(fillSeriesLine([1], 10, { ...base, type: 'growth', step: 2, stop: 10 })).toEqual([
      1, 2, 4, 8,
    ])
    expect(fillSeriesLine([-1], 10, { ...base, type: 'growth', step: 2, stop: -10 })).toEqual([
      -1, -2, -4, -8,
    ])
    expect(fillSeriesLine([-8], 10, { ...base, type: 'growth', step: 0.5, stop: -1 })).toEqual([
      -8, -4, -2, -1,
    ])
  })

  it('linear trend fits all seeds and rewrites them', () => {
    const values = fillSeriesLine([1, 3, 5.2], 5, { ...base, trend: true })
    expect(values.map((v) => Number(v.toFixed(4)))).toEqual([
      0.9667, 3.0667, 5.1667, 7.2667, 9.3667,
    ])
  })

  it('trend honours a stop value and never runs open-ended', () => {
    expect(fillSeriesLine([1, 2], 10, { ...base, trend: true, stop: 4 })).toEqual([1, 2, 3, 4])
    expect(fillSeriesLine([8, 4], 10, { ...base, trend: true, stop: 0 })).toEqual([8, 4, 0])
    expect(planFillSeries([[1]], { ...base, trend: true, stop: 100 }, 10_000)).toEqual([])
  })

  it('growth trend fits an exponential through the seeds', () => {
    const values = fillSeriesLine([2, 4, 8], 5, { ...base, type: 'growth', trend: true })
    expect(values.map((v) => Number(v.toFixed(9)))).toEqual([2, 4, 8, 16, 32])
    expect(fillSeriesLine([2, -4], 4, { ...base, type: 'growth', trend: true })).toEqual([2, -4])
  })

  it('autofill keeps the seeds and continues their difference', () => {
    expect(fillSeriesLine([10, 20], 5, { ...base, type: 'autofill' })).toEqual([10, 20, 30, 40, 50])
    expect(fillSeriesLine([7], 3, { ...base, type: 'autofill' })).toEqual([7, 8, 9])
    expect(fillSeriesLine([7], 9, { ...base, type: 'autofill', stop: 9 })).toEqual([7, 8, 9])
    expect(fillSeriesLine([9, 6], 9, { ...base, type: 'autofill', stop: 1 })).toEqual([9, 6, 3])
  })

  it('parses dialog numbers and, for dates, calendar text', () => {
    expect(parseSeriesNumber(' 2,5 ', false)).toBe(2.5)
    expect(parseSeriesNumber('', false)).toBeNull()
    expect(parseSeriesNumber('2024-01-31', false)).toBeNull()
    expect(parseSeriesNumber('2024-01-31T00:00', true)).toBe(JAN_31_2024)
    expect(parseSeriesNumber('2024-01-31', true)).toBe(JAN_31_2024)
    expect(parseSeriesNumber('2024-01-31', true, true)).toBe(JAN_31_2024 - 1462)
    expect(parseSeriesNumber('nonsense', true)).toBeNull()
  })

  it('date units step by day, weekday, month and year', () => {
    expect(stepDate(JAN_31_2024, 'day', 2)).toBe(JAN_31_2024 + 2)
    // 2024-01-31 is a Wednesday: +3 weekdays lands on Monday 02-05
    expect(stepDate(JAN_31_2024, 'weekday', 3)).toBe(JAN_31_2024 + 5)
    // month end clamps: Jan 31 -> Feb 29 (leap year)
    expect(stepDate(JAN_31_2024, 'month', 1)).toBe(JAN_31_2024 + 29)
    // Jan 31 2024 -> Jan 31 2025 = 366 days
    expect(stepDate(JAN_31_2024, 'year', 1)).toBe(JAN_31_2024 + 366)
    // time of day survives month stepping
    expect(stepDate(JAN_31_2024 + 0.5, 'month', 1)).toBeCloseTo(JAN_31_2024 + 29.5, 9)
  })

  it('1904-system serials step on the right calendar days', () => {
    const jan31In1904 = JAN_31_2024 - 1462
    expect(stepDate(jan31In1904, 'month', 1, true)).toBe(jan31In1904 + 29)
    expect(stepDate(jan31In1904, 'weekday', 3, true)).toBe(jan31In1904 + 5)
    expect(
      fillSeriesLine([jan31In1904], 3, { ...base, type: 'date', dateUnit: 'year', date1904: true }),
    ).toEqual([jan31In1904, jan31In1904 + 366, jan31In1904 + 731])
  })

  it('date series honours the stop serial', () => {
    expect(
      fillSeriesLine([JAN_31_2024], 10, {
        ...base,
        type: 'date',
        dateUnit: 'month',
        stop: JAN_31_2024 + 60,
      }),
    ).toEqual([JAN_31_2024, JAN_31_2024 + 29, JAN_31_2024 + 60])
  })
})

describe('planFillSeries', () => {
  it('fills down each column from its numeric head and skips text heads', () => {
    const writes = planFillSeries(
      [
        [1, 'x', 10],
        [null, null, null],
        [null, null, null],
      ],
      { ...base, step: 1 },
    )
    expect(writes).toEqual([
      { row: 1, column: 0, values: [[2], [3]] },
      { row: 1, column: 2, values: [[11], [12]] },
    ])
  })

  it('fills across rows and lets a trend rewrite the seeds', () => {
    expect(
      planFillSeries([[1, 2, null, null]], { ...base, direction: 'rows', trend: true }),
    ).toEqual([{ row: 0, column: 0, values: [[1, 2, 3, 4]] }])
    expect(planFillSeries([[1, 2, null, null]], { ...base, direction: 'rows' })).toEqual([
      { row: 0, column: 1, values: [[2, 3, 4]] },
    ])
  })

  it('a single cell with a stop value runs open-ended along the direction', () => {
    expect(planFillSeries([[1]], { ...base, step: 5, stop: 12 }, 100)).toEqual([
      { row: 1, column: 0, values: [[6], [11]] },
    ])
    expect(planFillSeries([[1]], { ...base, direction: 'rows', step: 5, stop: 12 }, 100)).toEqual([
      { row: 0, column: 1, values: [[6, 11]] },
    ])
    expect(planFillSeries([[1]], base)).toEqual([])
  })
})
