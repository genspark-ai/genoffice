import { describe, expect, it } from 'vitest'
import { MAX_PRINT_TITLE_ROWS, clampTitleRows } from '../src/renderer/print-settings'

describe('clampTitleRows', () => {
  it('keeps spans within the cap untouched', () => {
    expect(clampTitleRows(1, 1)).toBe('1:1')
    expect(clampTitleRows(5, 25)).toBe('5:25')
    expect(MAX_PRINT_TITLE_ROWS).toBe(21)
  })

  it('clamps tall selections to 21 rows anchored at the start', () => {
    expect(clampTitleRows(1, 100)).toBe('1:21')
    expect(clampTitleRows(5, 104)).toBe('5:25')
  })

  it('repairs reversed and non-finite input', () => {
    expect(clampTitleRows(10, 3)).toBe('10:10')
    expect(clampTitleRows(Number.NaN, 50)).toBe('1:21')
    expect(clampTitleRows(4, Infinity)).toBe('4:4')
  })
})
