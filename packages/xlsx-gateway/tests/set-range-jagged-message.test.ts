import { describe, expect, it } from 'vitest'

import { expandToPrimitiveOps, workbookOperationSchema } from '../src/domain/workbook-dsl'

function setRange(values: (string | number | null)[][]): unknown {
  return { op: 'set_range', sheetId: '1', start: 'A1', values }
}

function jaggedMessage(values: (string | number | null)[][]): string {
  const parsed = workbookOperationSchema.parse(setRange(values))
  try {
    expandToPrimitiveOps([parsed])
  } catch (error) {
    return (error as Error).message
  }
  throw new Error('expected a jagged-values error')
}

describe('set_range jagged values message', () => {
  it('names the 0-based array position of the offending row', () => {
    // values[2] is the short row. Reporting it as "row 3" reads like a
    // spreadsheet row number and points one line below the real offender.
    const message = jaggedMessage([[1, 2], [3, 4], [5]])
    expect(message).toContain('values[0] has 2 cell(s)')
    expect(message).toContain('values[2] has 1')
    expect(message).not.toMatch(/\brow \d/)
  })

  it('still describes the rectangle requirement', () => {
    expect(jaggedMessage([[1, 2], [3]])).toContain('rectangular')
  })

  it('accepts a rectangular grid', () => {
    const parsed = workbookOperationSchema.parse(
      setRange([
        [1, null],
        [null, 4],
      ]),
    )
    expect(expandToPrimitiveOps([parsed])).toHaveLength(4)
  })
})
