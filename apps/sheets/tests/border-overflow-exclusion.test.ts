import { describe, expect, it } from 'vitest'

import { overflowCoversBorder } from '../src/renderer/border-overflow-exclusion'

const cache = {
  getRow: (row: number) =>
    row === 5
      ? { 2: { startColumn: 2, endColumn: 4 }, 9: { startColumn: 9, endColumn: 9 } }
      : undefined,
}

describe('border overflow exclusion lookup', () => {
  it('matches the stock semantics for left and right edges inside an overflow run', () => {
    // run B..E (2..4): left edges of 3 and 4 are covered, 2 is the run start
    expect(overflowCoversBorder(cache, 'l', 5, 3)).toBe(true)
    expect(overflowCoversBorder(cache, 'l', 5, 4)).toBe(true)
    expect(overflowCoversBorder(cache, 'l', 5, 2)).toBe(false)
    expect(overflowCoversBorder(cache, 'l', 5, 5)).toBe(false)
    // right edges of 2 and 3 are covered, 4 is the run end
    expect(overflowCoversBorder(cache, 'r', 5, 2)).toBe(true)
    expect(overflowCoversBorder(cache, 'r', 5, 3)).toBe(true)
    expect(overflowCoversBorder(cache, 'r', 5, 4)).toBe(false)
    // a single-cell entry covers no edge
    expect(overflowCoversBorder(cache, 'l', 5, 9)).toBe(false)
    expect(overflowCoversBorder(cache, 'r', 5, 9)).toBe(false)
  })

  it('never excludes top/bottom edges or rows without overflow', () => {
    expect(overflowCoversBorder(cache, 't', 5, 3)).toBe(false)
    expect(overflowCoversBorder(cache, 'b', 5, 3)).toBe(false)
    expect(overflowCoversBorder(cache, 'l', 6, 3)).toBe(false)
    expect(overflowCoversBorder(null, 'l', 5, 3)).toBe(false)
  })
})
