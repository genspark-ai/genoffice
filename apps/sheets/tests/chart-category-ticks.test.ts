import { describe, expect, it } from 'vitest'

import { categoryTickLines, categoryTickStride } from '../src/renderer/WorkbookVisuals'

describe('categoryTickLines', () => {
  it('keeps labels that fit on one line', () => {
    expect(categoryTickLines('2015', 40)).toEqual(['2015'])
  })

  it('wraps at spaces when the slot is narrow (Excel "KW 01" → KW / 01)', () => {
    expect(categoryTickLines('KW 01', 12)).toEqual(['KW', '01'])
  })

  it('never wraps labels without a break point', () => {
    expect(categoryTickLines('Description', 12)).toEqual(['Description'])
  })
})

describe('categoryTickStride', () => {
  const labels = Array.from(
    { length: 48 },
    (_, index) => `KW ${String(index + 1).padStart(2, '0')}`,
  )

  it('thins overlapping labels to every n-th', () => {
    // 48 categories over 480 units: 10-unit slots, wrapped lines ~11 units.
    expect(categoryTickStride(labels, 48, 10)).toBeGreaterThan(1)
  })

  it('keeps every label when slots are wide', () => {
    expect(categoryTickStride(['2015', '2016', '2017'], 3, 160)).toBe(1)
  })
})
