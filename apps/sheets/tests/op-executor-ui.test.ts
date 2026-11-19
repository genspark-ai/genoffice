import { describe, expect, it } from 'vitest'

import { nextSheetName, planFromOps } from '../src/renderer/op-executor'

const workbook = {
  getSheetBySheetId: (id: string) => (id === 'sh1' ? { getSheetName: () => 'Data' } : undefined),
} as unknown as Parameters<typeof planFromOps>[1]

describe('nextSheetName', () => {
  it('picks the lowest free SheetN, case-insensitively', () => {
    expect(nextSheetName([])).toBe('Sheet1')
    expect(nextSheetName(['Sheet1', 'sheet2', 'Data'])).toBe('Sheet3')
    expect(nextSheetName(['Sheet2'])).toBe('Sheet1')
  })
})

describe('planFromOps', () => {
  it('wraps ribbon-built ops into the plan shape the executor applies', () => {
    const plan = planFromOps(
      [
        { op: 'insert_rows', sheetId: 'sh1', row: 3, count: 1 },
        { op: 'set_freeze', sheetId: 'sh1', rows: 1, columns: 0 },
        { op: 'merge_cells', sheetId: 'sh1', range: 'A1:C1' },
        {
          op: 'format_range',
          sheetId: 'sh1',
          range: 'A1:C1',
          format: { horizontalAlign: 'center' },
        },
      ],
      workbook,
      'ribbon',
    )
    expect(plan.structuralChanges.map((change) => change.op.op)).toEqual([
      'insert_rows',
      'set_freeze',
      'merge_cells',
    ])
    expect(plan.formatChanges).toHaveLength(1)
    expect(plan.cellChanges).toHaveLength(0)
    expect(plan.transactionId).toMatch(/^ui-/)
  })
})
