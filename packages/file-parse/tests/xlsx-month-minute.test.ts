import { expect, it } from 'vitest'
import { classifyFormatCode, formatSerial } from '../src/xlsx-dates'
it.each(['mmm h:mm', 'mm h:mm', 'h:mm mmmm'])(
  'keeps the calendar when %s has month and minute tokens',
  (code) => {
    const parts = classifyFormatCode(code)!
    expect(parts).toMatchObject({ date: true, time: true })
    expect(formatSerial(45292.5, parts, false)).toBe('2024-01-01 12:00')
  },
)
it('keeps a time-only format free of a calendar date', () => {
  expect(classifyFormatCode('h:mm')).toMatchObject({ date: false, time: true })
})
