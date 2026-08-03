import { CellValueType } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { clipboardField, plainTextFromCells } from '../src/renderer/clipboard-tsv'

describe('clipboard TSV serialization', () => {
  it('booleans copy as TRUE/FALSE, not 1/0', () => {
    expect(clipboardField({ v: 1, t: CellValueType.BOOLEAN })).toBe('TRUE')
    expect(clipboardField({ v: 0, t: CellValueType.BOOLEAN })).toBe('FALSE')
  })

  it('prefers the formatted display text', () => {
    expect(clipboardField({ v: 44614, displayV: '2/22/2022' })).toBe('2/22/2022')
  })

  it('quotes fields with embedded newlines and normalizes \\r to \\n', () => {
    expect(clipboardField({ v: 'greater \r\rthan' })).toBe('"greater \nthan"')
    expect(clipboardField({ v: 'a\tb' })).toBe('"a\tb"')
    expect(clipboardField({ v: 'say "hi"' })).toBe('"say ""hi"""')
    expect(clipboardField({ v: 'plain' })).toBe('plain')
  })

  it('assembles rows with empty fields for missing cells (no column drift)', () => {
    const cells: Record<string, { v: string }> = { '0:0': { v: 'a' }, '0:2': { v: 'c' } }
    const plain = plainTextFromCells([0, 1], [0, 1, 2], (row, column) => cells[`${row}:${column}`])
    expect(plain).toBe('a\t\tc\n\t\t')
  })
})
