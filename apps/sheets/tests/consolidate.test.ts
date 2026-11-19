import { describe, expect, it } from 'vitest'

import {
  buildLabelMatrix,
  buildPositionMatrix,
  parseConsolidateReference,
  targetOverlapsSource,
  type ConsolidateArea,
} from '../src/renderer/consolidate'

const area = (
  sheetName: string,
  startRow: number,
  startColumn: number,
  rows: number,
  columns: number,
): ConsolidateArea => ({ sheetName, startRow, startColumn, rows, columns })

describe('parseConsolidateReference', () => {
  it('parses bare, qualified, and quoted references', () => {
    expect(parseConsolidateReference('B2:D9')).toEqual({
      sheetName: null,
      range: { startRow: 1, startColumn: 1, endRow: 8, endColumn: 3 },
    })
    expect(parseConsolidateReference('Sheet2!A1')?.sheetName).toBe('Sheet2')
    expect(parseConsolidateReference("'My Sheet'!A1:B2")?.sheetName).toBe('My Sheet')
    expect(parseConsolidateReference("'It''s'!A1")?.sheetName).toBe("It's")
  })

  it('normalizes case, $ markers, and inverted corners', () => {
    expect(parseConsolidateReference('$c$3:$a$1')?.range).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 2,
      endColumn: 2,
    })
  })

  it('rejects malformed input', () => {
    expect(parseConsolidateReference('')).toBeNull()
    expect(parseConsolidateReference('A1:B2 junk')).toBeNull()
    expect(parseConsolidateReference('Sheet2!')).toBeNull()
    expect(parseConsolidateReference('A:A')).toBeNull()
  })
})

describe('buildPositionMatrix', () => {
  it('aggregates the same offsets across areas', () => {
    const matrix = buildPositionMatrix('sum', [
      area('Q1', 1, 1, 2, 2),
      area('Q2', 0, 4, 2, 2),
    ])
    expect(matrix).toHaveLength(2)
    expect(matrix[0]?.[0]).toEqual({ f: '=SUM(Q1!B2,Q2!E1)' })
    expect(matrix[1]?.[1]).toEqual({ f: '=SUM(Q1!C3,Q2!F2)' })
  })

  it('pads ragged areas and quotes sheet names with spaces', () => {
    const matrix = buildPositionMatrix('max', [
      area('My Sheet', 0, 0, 1, 1),
      area('Plain', 0, 0, 2, 1),
    ])
    expect(matrix[0]?.[0]).toEqual({ f: "=MAX('My Sheet'!A1,Plain!A1)" })
    expect(matrix[1]?.[0]).toEqual({ f: '=MAX(Plain!A2)' })
  })
})

describe('buildLabelMatrix', () => {
  const areas = [area('Q1', 1, 0, 3, 2), area('Q2', 1, 3, 2, 2)]
  const labels = [['East', 'West', 'East'], ['West', 'North']]

  it('unions labels in first-appearance order with per-area SUMIF parts', () => {
    const matrix = buildLabelMatrix('sum', areas, labels, { row: 0, column: 6 })
    expect(Array.isArray(matrix)).toBe(true)
    if (typeof matrix === 'string') return
    expect(matrix.map((line) => line[0]?.v)).toEqual(['East', 'West', 'North'])
    expect(matrix[0]?.[1]).toEqual({ f: '=SUMIF(Q1!$A$2:$A$4,$G1,Q1!B2:B4)' })
    expect(matrix[1]?.[1]).toEqual({
      f: '=SUMIF(Q1!$A$2:$A$4,$G2,Q1!B2:B4)+SUMIF(Q2!$D$2:$D$3,$G2,Q2!E2:E3)',
    })
    expect(matrix[2]?.[1]).toEqual({ f: '=SUMIF(Q2!$D$2:$D$3,$G3,Q2!E2:E3)' })
  })

  it('average divides summed parts by counted parts', () => {
    const matrix = buildLabelMatrix('average', areas, labels, { row: 0, column: 6 })
    if (typeof matrix === 'string') throw new Error(matrix)
    expect(matrix[2]?.[1]?.f).toBe(
      '=(SUMIF(Q2!$D$2:$D$3,$G3,Q2!E2:E3))/(COUNTIF(Q2!$D$2:$D$3,$G3))',
    )
  })

  it('rejects max/min, single-column areas, and label-less sources', () => {
    expect(buildLabelMatrix('max', areas, labels, { row: 0, column: 0 }))
      .toMatch(/不支持/)
    expect(buildLabelMatrix('sum', [area('Q1', 0, 0, 2, 1)], [['A']], { row: 0, column: 0 }))
      .toMatch(/两列/)
    expect(buildLabelMatrix('sum', areas, [[''], []], { row: 0, column: 0 }))
      .toMatch(/没有找到标签/)
  })
})

describe('targetOverlapsSource', () => {
  const sources = [area('Sheet1', 0, 0, 3, 2)]

  it('detects same-sheet overlap and clears disjoint or cross-sheet targets', () => {
    expect(targetOverlapsSource(sources, 'Sheet1', { row: 2, column: 1, rows: 2, columns: 2 }))
      .toBe(true)
    expect(targetOverlapsSource(sources, 'Sheet1', { row: 0, column: 2, rows: 3, columns: 2 }))
      .toBe(false)
    expect(targetOverlapsSource(sources, 'Sheet2', { row: 0, column: 0, rows: 3, columns: 2 }))
      .toBe(false)
  })
})
