import { describe, expect, it } from 'vitest'
import { MERGE_RANGE_ALL, recordsInRange, type MergeRange } from '../src/renderer/mail-merge/range'

const list = ['a', 'b', 'c', 'd', 'e']
const custom = (from: number, to: number | null): MergeRange => ({ kind: 'custom', from, to })

describe('recordsInRange', () => {
  it('takes the whole list for All', () => {
    expect(recordsInRange(list, MERGE_RANGE_ALL, 3)).toEqual(list)
  })

  it('takes the previewed record for Current Record', () => {
    expect(recordsInRange(list, { kind: 'current', from: 2, to: 3 }, 3)).toEqual(['d'])
    expect(recordsInRange([], { kind: 'current', from: 1, to: null }, 0)).toEqual([])
  })

  it('takes From through To, counted from 1 and inclusive', () => {
    expect(recordsInRange(list, custom(2, 4), 0)).toEqual(['b', 'c', 'd'])
    expect(recordsInRange(list, custom(3, 3), 0)).toEqual(['c'])
  })

  it('runs to the end of the list with no To, and stops at the end with a To past it', () => {
    expect(recordsInRange(list, custom(4, null), 0)).toEqual(['d', 'e'])
    expect(recordsInRange(list, custom(4, 99), 0)).toEqual(['d', 'e'])
  })

  it('starts at the first record with a From below 1', () => {
    expect(recordsInRange(list, custom(0, 2), 0)).toEqual(['a', 'b'])
    expect(recordsInRange(list, custom(Number.NaN, 2), 0)).toEqual(['a', 'b'])
  })

  it('takes nothing when From is past To or past the list', () => {
    expect(recordsInRange(list, custom(4, 2), 0)).toEqual([])
    expect(recordsInRange(list, custom(9, null), 0)).toEqual([])
  })
})
