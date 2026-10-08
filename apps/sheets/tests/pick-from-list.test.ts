import { describe, expect, it } from 'vitest'

import {
  PICK_LIST_SCAN_LIMIT,
  collectPickListValues,
  filterCriteriaForText,
  type PickListEntry,
} from '../src/renderer/pick-from-list'

const column = (values: ReadonlyArray<string | null>) => (row: number) => {
  const value = values[row]
  return value === null || value === undefined ? null : { display: value, cell: { v: value } }
}

const displays = (entries: readonly PickListEntry[]) => entries.map((entry) => entry.display)

describe('collectPickListValues', () => {
  it('lists the contiguous block around the cell, unique and sorted', () => {
    const read = column(['Header', 'pear', 'Apple', 'apple', 'fig', null, 'other'])
    expect(displays(collectPickListValues(read, 2, 7))).toEqual(['Apple', 'fig', 'Header', 'pear'])
  })

  it('an empty cell takes the blocks directly above and below it', () => {
    const read = column(['a', 'b', null, 'c', null, 'z'])
    expect(displays(collectPickListValues(read, 2, 6))).toEqual(['a', 'b', 'c'])
    expect(displays(collectPickListValues(read, 4, 6))).toEqual(['c', 'z'])
  })

  it('whitespace-only cells end the block and never show up', () => {
    const read = column(['x', '  ', 'y'])
    expect(displays(collectPickListValues(read, 0, 3))).toEqual(['x'])
    expect(displays(collectPickListValues(read, 1, 3))).toEqual(['x', 'y'])
  })

  it('sorts numerically and keeps the first spelling of a duplicate', () => {
    const read = column(['10', '9', 'B', 'b', '100'])
    expect(displays(collectPickListValues(read, 0, 5))).toEqual(['9', '10', '100', 'B'])
  })

  it('returns nothing for an isolated empty cell and stops scanning at the limit', () => {
    expect(collectPickListValues(column([null, null, null]), 1, 3)).toEqual([])
    let reads = 0
    const endless = () => {
      reads += 1
      return { display: `v${reads}`, cell: null }
    }
    collectPickListValues(endless, 0, PICK_LIST_SCAN_LIMIT * 3)
    expect(reads).toBe(PICK_LIST_SCAN_LIMIT + 1)
  })
})

describe('filterCriteriaForText', () => {
  it('maps an empty cell to the blank filter and text to a one-value list', () => {
    expect(filterCriteriaForText('')).toEqual({ blank: true })
    expect(filterCriteriaForText('42')).toEqual({ filters: ['42'] })
  })
})
