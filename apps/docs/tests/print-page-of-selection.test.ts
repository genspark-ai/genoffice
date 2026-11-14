import { describe, expect, it } from 'vitest'
import { initialPrintPageIndex, pageOfBlockId } from '../src/renderer/print-page-of-selection'

describe('pageOfBlockId', () => {
  it('finds the sheet holding the caret block', () => {
    expect(pageOfBlockId([['0', '1'], ['2', '3'], ['4']], '3')).toBe(1)
    expect(pageOfBlockId([['0', '1'], ['2', '3'], ['4']], '0')).toBe(0)
  })

  it('uses the first sheet for a block split across pages', () => {
    expect(pageOfBlockId([['7'], ['7'], ['8']], '7')).toBe(0)
  })

  it('returns null when the block is not in the preview', () => {
    expect(pageOfBlockId([['0'], ['1']], '99')).toBeNull()
    expect(pageOfBlockId([], '0')).toBeNull()
    expect(pageOfBlockId([['0']], null)).toBeNull()
    expect(pageOfBlockId([['0']], '')).toBeNull()
  })
})

describe('initialPrintPageIndex', () => {
  it('opens on the caret page when it can be resolved', () => {
    expect(initialPrintPageIndex([['0'], ['1'], ['2']], '2')).toBe(2)
  })

  it('falls back to the first page when it cannot', () => {
    expect(initialPrintPageIndex([['0'], ['1']], null)).toBe(0)
    expect(initialPrintPageIndex([['0'], ['1']], '42')).toBe(0)
    expect(initialPrintPageIndex([], null)).toBe(0)
  })
})
