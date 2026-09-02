import { describe, expect, it } from 'vitest'

import { formatPiePercent, lineSegments } from '../src/renderer/WorkbookVisuals'

describe('lineSegments', () => {
  it('plots one run when there are no blanks or the mode keeps zeros', () => {
    expect(lineSegments(3, undefined, 'gap')).toEqual([[0, 1, 2]])
    expect(lineSegments(3, [1], undefined)).toEqual([[0, 1, 2]])
    expect(lineSegments(3, [1], 'zero')).toEqual([[0, 1, 2]])
    expect(lineSegments(0, undefined, undefined)).toEqual([])
  })

  it('breaks the line at blank cells for dispBlanksAs=gap', () => {
    expect(lineSegments(6, [2, 3], 'gap')).toEqual([
      [0, 1],
      [4, 5],
    ])
    expect(lineSegments(4, [0, 3], 'gap')).toEqual([[1, 2]])
    expect(lineSegments(2, [0, 1], 'gap')).toEqual([])
  })

  it('bridges blank cells for dispBlanksAs=span', () => {
    expect(lineSegments(6, [2, 3], 'span')).toEqual([[0, 1, 4, 5]])
    expect(lineSegments(2, [0, 1], 'span')).toEqual([])
  })
})

describe('formatPiePercent', () => {
  it('rounds to whole percents by default, matching Excel showPercent (0%)', () => {
    expect(formatPiePercent(0.094, undefined)).toBe('9%')
    expect(formatPiePercent(0.095, undefined)).toBe('10%')
    expect(formatPiePercent(0.004, undefined)).toBe('0%')
    // Value formats on the source cells are not percent formats.
    expect(formatPiePercent(0.094, '#,##0')).toBe('9%')
  })

  it('honors an explicit percent numFmt on the dLbls', () => {
    expect(formatPiePercent(0.094, '0.0%')).toBe('9.4%')
    expect(formatPiePercent(0.094, '0%')).toBe('9%')
  })
})
