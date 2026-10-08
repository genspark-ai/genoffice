import { describe, expect, it } from 'vitest'
import { planFileTableRegistrations, sanitizeTableName } from '../src/renderer/file-tables'
import type { WorkbookFile } from '../src/shared/desktop-api'

type FileSheet = WorkbookFile['sheets'][number]
type FileTable = FileSheet['tables'][number]

function table(partial: Partial<FileTable>, named = true): FileTable {
  return {
    range: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 1 },
    headerRowCount: 1,
    showRowStripes: true,
    showColumnStripes: false,
    columns: ['Item', 'Amount'],
    ...(named ? { name: 'T' } : {}),
    ...partial,
  } as FileTable
}

function sheet(id: string, name: string, tables: FileTable[]): FileSheet {
  return { id, name, tables } as unknown as FileSheet
}

describe('planFileTableRegistrations', () => {
  it('registers each table under its file displayName, same columns on two sheets', () => {
    const plan = planFileTableRegistrations([
      sheet('s1', 'Data', [table({ name: 'Sales' })]),
      sheet('s2', 'Other', [table({ name: 'Costs' })]),
    ])
    expect(plan.map((entry) => [entry.sheetId, entry.tableName, entry.tableId])).toEqual([
      ['s1', 'Sales', 'file-table-s1-0'],
      ['s2', 'Costs', 'file-table-s2-0'],
    ])
    expect(plan.map((entry) => entry.columns)).toEqual([
      ['Item', 'Amount'],
      ['Item', 'Amount'],
    ])
  })

  it('keeps names unique case-insensitively across sheets and session tables', () => {
    const plan = planFileTableRegistrations(
      [
        sheet('s1', 'Data', [table({ name: 'Sales' }), table({ name: 'Data' })]),
        sheet('s2', 'Other', [table({ name: 'sales' }), table({ name: 'Report' })]),
      ],
      ['report'],
    )
    expect(plan.map((entry) => entry.tableName)).toEqual(['Sales', 'Data_2', 'sales_2', 'Report_2'])
  })

  it('sanitises Excel-invalid names and falls back when the part has none', () => {
    const plan = planFileTableRegistrations([
      sheet('s1', 'Data', [
        table({ name: 'Q1 Sales/EU' }),
        table({ name: '2024' }),
        table({ name: 'A1' }),
        table({}, false),
        table({ name: '\u8868\u683c.1' }),
      ]),
    ])
    expect(plan.map((entry) => entry.tableName)).toEqual([
      'Q1_Sales_EU',
      '_2024',
      '_A1',
      'Table1',
      '\u8868\u683c.1',
    ])
  })

  it('skips headerless tables and trims the totals band from the range', () => {
    const plan = planFileTableRegistrations([
      sheet('s1', 'Data', [
        table({ name: 'NoHeader', headerRowCount: 0 }),
        table({ name: 'WithTotals', totalsRowCount: 1 }),
      ]),
    ])
    expect(plan).toHaveLength(1)
    expect(plan[0]!.tableName).toBe('WithTotals')
    expect(plan[0]!.range).toEqual({ startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 })
  })
})

describe('sanitizeTableName', () => {
  it('caps the length at 255', () => {
    expect(sanitizeTableName('x'.repeat(300))).toHaveLength(255)
  })
  it('returns an empty string for blank input', () => {
    expect(sanitizeTableName('   ')).toBe('')
  })
})
