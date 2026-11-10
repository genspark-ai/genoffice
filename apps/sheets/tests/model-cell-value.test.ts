/**
 * Reading a cell's value for anything machine-facing.
 *
 * The view model's `value` is what the cell RENDERS. Two interceptor layers
 * rewrite it: number formats turn dates/percentages into display strings, and
 * General shrinks a number to the column's width (numfmt-fix.ts formatGeneral).
 * A reader that reports or re-saves that text turns a computed number into a
 * string — the bug this covers: `=1/3` in a narrow column reached `read_sheet`
 * and the saved file as "0.333333" with t="str".
 */
import { describe, expect, it } from 'vitest'
import { modelCellValue } from '../src/renderer/univer-sync'

describe('modelCellValue', () => {
  it('prefers the model number over the width-fitted display text', () => {
    // 1/3 in a default-width column renders as "0.333333"
    expect(modelCellValue({ value: '0.333333', rawValue: 0.3333333333333333 })).toBe(
      0.3333333333333333,
    )
    expect(modelCellValue({ value: '0.142857', rawValue: 0.14285714285714285 })).toBe(
      0.14285714285714285,
    )
  })

  it('prefers the model number over a number format (date, percent)', () => {
    expect(modelCellValue({ value: '9/18/2026', rawValue: 46282 })).toBe(46282)
    expect(modelCellValue({ value: '25%', rawValue: 0.25 })).toBe(0.25)
  })

  it('keeps a genuine text result as text', () => {
    expect(modelCellValue({ value: 'hello', rawValue: 'hello' })).toBe('hello')
  })

  it('falls back to the display when the model carries nothing', () => {
    expect(modelCellValue({ value: 12 })).toBe(12)
    expect(modelCellValue({ value: 'x', rawValue: null })).toBe('x')
  })

  it('keeps the booleans the engine produced', () => {
    expect(modelCellValue({ value: 'TRUE', rawValue: true })).toBe(true)
  })

  it('surfaces an error the engine produced, so consumers see it', () => {
    expect(modelCellValue({ value: '#DIV/0!', rawValue: '#DIV/0!' })).toBe('#DIV/0!')
  })

  it('keeps the cached display when the engine errored behind it', () => {
    // formula-cached-fallback.ts shows the file's cached value where the engine
    // returned #NAME? — the visible value is the better answer, not the error.
    expect(modelCellValue({ value: 120, rawValue: '#NAME?' })).toBe(120)
    expect(modelCellValue({ value: 'Widgets', rawValue: '#NAME?' })).toBe('Widgets')
  })
})
