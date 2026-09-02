import { describe, expect, it } from 'vitest'

import { evenPageRanges, stitchPlan, variantForPage } from '../src/main/pdf-page-variants'

describe('variantForPage', () => {
  it('uses the odd templates everywhere when no variant is active', () => {
    const flags = { hasFirst: false, hasEven: false }
    expect([1, 2, 3, 4].map((page) => variantForPage(page, flags))).toEqual([
      'odd',
      'odd',
      'odd',
      'odd',
    ])
  })

  it('gives page 1 the first-page templates under differentFirst', () => {
    const flags = { hasFirst: true, hasEven: false }
    expect([1, 2, 3].map((page) => variantForPage(page, flags))).toEqual(['first', 'odd', 'odd'])
  })

  it('alternates odd/even under differentOddEven', () => {
    const flags = { hasFirst: false, hasEven: true }
    expect([1, 2, 3, 4, 5].map((page) => variantForPage(page, flags))).toEqual([
      'odd',
      'even',
      'odd',
      'even',
      'odd',
    ])
  })

  it('lets the first page win over parity when both flags are set', () => {
    const flags = { hasFirst: true, hasEven: true }
    expect([1, 2, 3, 4].map((page) => variantForPage(page, flags))).toEqual([
      'first',
      'even',
      'odd',
      'even',
    ])
  })
})

describe('evenPageRanges', () => {
  it('lists the even pages for Chromium pageRanges', () => {
    expect(evenPageRanges(1)).toBe('')
    expect(evenPageRanges(2)).toBe('2')
    expect(evenPageRanges(7)).toBe('2,4,6')
    expect(evenPageRanges(8)).toBe('2,4,6,8')
  })
})

describe('stitchPlan', () => {
  it('indexes each page into the pass that printed it', () => {
    expect(stitchPlan(5, { hasFirst: true, hasEven: true })).toEqual([
      { page: 1, source: 'first', index: 0 },
      { page: 2, source: 'even', index: 0 },
      { page: 3, source: 'odd', index: 2 },
      { page: 4, source: 'even', index: 1 },
      { page: 5, source: 'odd', index: 4 },
    ])
  })

  it('is the identity over the odd pass without variants', () => {
    expect(stitchPlan(3, { hasFirst: false, hasEven: false })).toEqual([
      { page: 1, source: 'odd', index: 0 },
      { page: 2, source: 'odd', index: 1 },
      { page: 3, source: 'odd', index: 2 },
    ])
    expect(stitchPlan(0, { hasFirst: true, hasEven: true })).toEqual([])
  })
})
