import { describe, expect, it } from 'vitest'

import { fullColumnSpans, fullRowSpans } from '../src/renderer/autofit-selection'

describe('AutoFit spans', () => {
  const selections = [
    { startRow: 2, endRow: 4, startColumn: 1, endColumn: 1 },
    { startRow: 9, endRow: 9, startColumn: 3, endColumn: 5 },
  ]

  it('widens each selection to the full row so every cell in the row counts', () => {
    expect(fullRowSpans(selections, 26)).toEqual([
      { startRow: 2, endRow: 4, startColumn: 0, endColumn: 25 },
      { startRow: 9, endRow: 9, startColumn: 0, endColumn: 25 },
    ])
  })

  it('stretches each selection to the full column', () => {
    expect(fullColumnSpans(selections, 100)).toEqual([
      { startRow: 0, endRow: 99, startColumn: 1, endColumn: 1 },
      { startRow: 0, endRow: 99, startColumn: 3, endColumn: 5 },
    ])
  })
})
