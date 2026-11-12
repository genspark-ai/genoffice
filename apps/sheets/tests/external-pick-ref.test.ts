import { describe, expect, it } from 'vitest'

import {
  formatExternalPickRef,
  formatExternalRef,
  type ExternalPickTarget,
} from '../src/renderer/function-arguments-runtime'

const TARGET: ExternalPickTarget = {
  index: 2,
  name: 'prices.xlsx',
  source: '/tmp/data/prices.xlsx',
  open: true,
  wcId: 7,
  sheetNames: ['Sheet1', 'Price 2024'],
}

describe('formatExternalRef', () => {
  it('leaves a bare sheet name unquoted, the index prefix attached', () => {
    expect(formatExternalRef(2, 'Sheet1', 'A1:B2')).toBe('[2]Sheet1!A1:B2')
  })

  it('quotes a spaced name with the index inside the quotes', () => {
    expect(formatExternalRef(1, 'Price 2024', 'A1')).toBe(`'[1]Price 2024'!A1`)
  })

  it('doubles quotes inside the sheet name', () => {
    expect(formatExternalRef(3, "It's", 'A1')).toBe(`'[3]It''s'!A1`)
  })
})

describe('formatExternalPickRef', () => {
  it('uses the stored sheet spelling when the pick matches one', () => {
    expect(
      formatExternalPickRef(TARGET, {
        sheetName: 'price 2024',
        startRow: 0,
        endRow: 2,
        startColumn: 0,
        endColumn: 3,
      }),
    ).toBe(`'[2]Price 2024'!A1:D3`)
  })

  it('keeps whole-column and whole-row pick bodies', () => {
    expect(
      formatExternalPickRef(TARGET, {
        sheetName: 'Sheet1',
        startRow: 0,
        endRow: 0,
        startColumn: 0,
        endColumn: 2,
        rangeType: 2,
      }),
    ).toBe('[2]Sheet1!A:C')
    expect(
      formatExternalPickRef(TARGET, {
        sheetName: 'Sheet1',
        startRow: 1,
        endRow: 4,
        startColumn: 0,
        endColumn: 0,
        rangeType: 1,
      }),
    ).toBe('[2]Sheet1!2:5')
  })

  it('falls back to the picked live sheet name', () => {
    expect(
      formatExternalPickRef(TARGET, {
        sheetName: 'Renamed Later',
        startRow: 0,
        endRow: 0,
        startColumn: 0,
        endColumn: 0,
      }),
    ).toBe(`'[2]Renamed Later'!A1`)
  })
})
