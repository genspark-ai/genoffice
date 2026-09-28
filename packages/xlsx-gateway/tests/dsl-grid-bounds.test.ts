import { describe, expect, it } from 'vitest'

import { workbookOperationSchema } from '../src/domain/workbook-dsl'

function accepts(operation: unknown): boolean {
  return workbookOperationSchema.safeParse(operation).success
}

describe('workbook DSL grid bounds', () => {
  it('accepts the last grid cell and rejects past it', () => {
    expect(accepts({ op: 'set_cell', sheetId: '1', address: 'XFD1048576', value: 1 })).toBe(true)
    // The regex alone allowed any 1-3 letter column and any 7-digit row.
    expect(accepts({ op: 'set_cell', sheetId: '1', address: 'XFE1', value: 1 })).toBe(false)
    expect(accepts({ op: 'set_cell', sheetId: '1', address: 'ZZZ1', value: 1 })).toBe(false)
    expect(accepts({ op: 'set_cell', sheetId: '1', address: 'A1048577', value: 1 })).toBe(false)
  })

  it('rejects a past-the-edge column label', () => {
    const sort = (byColumn: string) => ({
      op: 'sort_range',
      sheetId: '1',
      range: 'A1:C10',
      byColumn,
      order: 'asc',
    })
    expect(accepts(sort('A'))).toBe(true)
    expect(accepts(sort('XFD'))).toBe(true)
    expect(accepts(sort('XFE'))).toBe(false)
  })

  it('names the grid limit in the error', () => {
    const result = workbookOperationSchema.safeParse({
      op: 'set_cell',
      sheetId: '1',
      address: 'XFE1',
      value: 1,
    })
    expect(result.success).toBe(false)
    expect(result.error?.issues[0]?.message).toMatch(/grid/i)
  })
})
