import { describe, expect, it } from 'vitest'

import { dedupeRows } from '../src/renderer/dedupe'

describe('dedupeRows', () => {
  it('keeps the first occurrence and reports the removed count', () => {
    const { rows, removed } = dedupeRows(
      [
        ['a', 1],
        ['b', 2],
        ['a', 1],
        ['a', 2],
        ['b', 2],
      ],
      false,
    )
    expect(rows).toEqual([
      ['a', 1],
      ['b', 2],
      ['a', 2],
    ])
    expect(removed).toBe(2)
  })

  it('compares text case-insensitively', () => {
    const { rows, removed } = dedupeRows([['Apple'], ['APPLE'], ['apple ']], false)
    expect(rows).toEqual([['Apple'], ['apple ']])
    expect(removed).toBe(1)
  })

  it('never removes the header row and never matches against it', () => {
    const { rows, removed } = dedupeRows([['Name'], ['Name'], ['name']], true)
    expect(rows).toEqual([['Name'], ['Name']])
    expect(removed).toBe(1)
  })

  it('does not confuse types or nulls with equal string forms', () => {
    const { rows, removed } = dedupeRows(
      [[1, null], ['1', null], [null, null], [null, null], [true], ['true']],
      false,
    )
    expect(removed).toBe(1)
    expect(rows).toHaveLength(5)
  })
})
