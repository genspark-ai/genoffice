import { describe, expect, it } from 'vitest'

import {
  createEditJournal,
  recordTableAdd,
  recordTableEdit,
  tableEditFor,
  toSaveTableEdits,
  updateTableAdd,
} from '../src/renderer/edit-journal'
import {
  isValidTableName,
  optionsOfTableAdd,
  resizeProblem,
  resolveCreateTableRange,
  totalsFunctionFor,
  uniqueColumnNames,
} from '../src/renderer/table-design'
import { t } from '../src/renderer/i18n/locale'
import { applyTint, builtinTablePalette, tableThemeJson } from '../src/renderer/table-styles'

const _ = null
type Cell = string | number | null

function hasValueIn(grid: Cell[][]): (row: number, column: number) => boolean {
  return (row, column) => {
    const value = grid[row]?.[column]
    return value !== null && value !== undefined && value !== ''
  }
}

describe('resolveCreateTableRange', () => {
  const grid: Cell[][] = [
    [_, _, _, _, _],
    [_, 'Name', 'Qty', 'Amount', _],
    [_, 'a', 1, 10, _],
    [_, 'b', _, 20, _],
    [_, _, _, _, _],
    [_, 'x', 'y', _, _],
  ]
  const last = { row: 5, column: 4 }

  it('widens a single cell to the contiguous data block', () => {
    const cell = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 }
    expect(resolveCreateTableRange(cell, hasValueIn(grid), last)).toEqual({
      startRow: 1,
      startColumn: 1,
      endRow: 3,
      endColumn: 3,
    })
  })

  it('keeps a multi-cell selection as drawn', () => {
    const selection = { startRow: 1, startColumn: 1, endRow: 2, endColumn: 2 }
    expect(resolveCreateTableRange(selection, hasValueIn(grid), last)).toEqual(selection)
  })

  it('falls back to the cell itself when it is empty', () => {
    const cell = { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }
    expect(resolveCreateTableRange(cell, hasValueIn(grid), last)).toEqual(cell)
  })
})

describe('handleTableDesignCommand parsing', () => {
  it('keeps the colon inside the Create Table range', async () => {
    const { handleTableDesignCommand } = await import('../src/renderer/table-design-actions')
    const messages: string[] = []
    const ctx = {
      univerRef: { current: null },
      lazyWorkbookRef: { current: null },
      setMessage: (message: string) => messages.push(message),
      setPendingEdits: () => undefined,
      refreshSelection: () => undefined,
    }
    expect(handleTableDesignCommand(ctx, 'table-create:A1:D10:0')).toBe(true)
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(messages).toEqual([t('appTablesNeedFile')])
  })
})

describe('table names and columns', () => {
  it('accepts Excel table names and rejects cell-like ones', () => {
    expect(isValidTableName('Sales_2024')).toBe(true)
    expect(isValidTableName('_tbl.v2')).toBe(true)
    expect(isValidTableName('A1')).toBe(false)
    expect(isValidTableName('R1C1')).toBe(false)
    expect(isValidTableName('my table')).toBe(false)
    expect(isValidTableName('1st')).toBe(false)
    expect(isValidTableName('')).toBe(false)
  })

  it('fills blank headers and de-duplicates repeats like Excel', () => {
    expect(uniqueColumnNames(['Name', '', 'name', 'Qty', 'Qty'])).toEqual([
      'Name',
      'Column2',
      'name2',
      'Qty',
      'Qty2',
    ])
  })

  it('picks SUM for numeric columns and COUNTA otherwise', () => {
    expect(totalsFunctionFor([1, 2, null])).toBe(109)
    expect(totalsFunctionFor(['a', 2])).toBe(103)
    expect(totalsFunctionFor([])).toBe(103)
  })
})

describe('resizeProblem', () => {
  const current = { startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 }
  const options = { headerRow: true, totalsRow: false }

  it('allows growing and shrinking while the header row stays', () => {
    expect(resizeProblem(current, { ...current, endRow: 9, endColumn: 4 }, options)).toBeNull()
    expect(resizeProblem(current, { ...current, endRow: 2 }, options)).toBeNull()
  })

  it('rejects moving the header row, losing overlap or losing data rows', () => {
    expect(resizeProblem(current, { ...current, startRow: 1 }, options)).toBe('header-row-moved')
    expect(
      resizeProblem(current, { startRow: 0, startColumn: 5, endRow: 4, endColumn: 6 }, options),
    ).toBe('no-overlap')
    expect(resizeProblem(current, { ...current, endRow: 0 }, options)).toBe('no-data-rows')
    expect(
      resizeProblem(current, { ...current, endRow: 1 }, { headerRow: true, totalsRow: true }),
    ).toBe('no-data-rows')
  })
})

describe('built-in table styles', () => {
  it('tints like Excel', () => {
    expect(applyTint('#4472C4', 0.8)).toBe('#DAE3F3')
    expect(applyTint('#4472C4', -0.25)).toBe('#2F5597')
  })

  it('resolves Medium 2 against the default theme', () => {
    const palette = builtinTablePalette('TableStyleMedium2', undefined)
    expect(palette.headerFill).toBe('#4472C4')
    expect(palette.headerFontColor).toBe('#FFFFFF')
    expect(palette.stripeFill).toBe('#DAE3F3')
    expect(palette.totalRowBorderStyle).toBe('double')
  })

  it('follows the workbook theme accents', () => {
    const theme = ['#FFFFFF', '#111111', '#EEEEEE', '#222222', '#AA0000', '#00AA00']
    const padded = [...theme, '#0000AA', '#AAAA00', '#AA00AA', '#00AAAA', '#000000', '#000000']
    expect(builtinTablePalette('TableStyleMedium3', padded).headerFill).toBe('#00AA00')
    expect(builtinTablePalette('TableStyleDark1', padded).headerFill).toBe('#111111')
  })

  it('maps palette and options onto a Univer range theme', () => {
    const theme = tableThemeJson('t', builtinTablePalette('TableStyleMedium2', undefined), {
      bandedRows: true,
      bandedColumns: false,
      firstColumn: true,
      lastColumn: false,
    })
    expect(theme.headerRowStyle?.bg?.rgb).toBe('#4472C4')
    expect(theme.firstRowStyle?.bg?.rgb).toBe('#DAE3F3')
    expect(theme.secondRowStyle).toEqual({})
    expect(theme.headerColumnStyle).toEqual({ bl: 1 })
    expect(theme.firstColumnStyle).toBeUndefined()
  })
})

describe('table journal', () => {
  it('merges file-table edits and collapses them on removal', () => {
    const journal = createEditJournal()
    recordTableEdit(journal, { sheetId: 's1', tableName: 'Sales', style: 'TableStyleLight9' })
    recordTableEdit(journal, { sheetId: 's1', tableName: 'sales', name: 'Revenue' })
    expect(tableEditFor(journal, 's1', 'Sales')).toEqual({
      sheetId: 's1',
      tableName: 'sales',
      style: 'TableStyleLight9',
      name: 'Revenue',
    })
    recordTableEdit(journal, { sheetId: 's1', tableName: 'Sales', remove: true })
    expect(toSaveTableEdits(journal)).toEqual([{ sheetId: 's1', tableName: 'Sales', remove: true }])
  })

  it('patches session tables with style options', () => {
    const journal = createEditJournal()
    recordTableAdd(journal, {
      sheetId: 's1',
      area: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 },
      name: 'Table1',
      columnNames: ['A', 'B'],
      bandedRows: true,
    })
    expect(updateTableAdd(journal, 's1', 'table1', { totalsRow: true, firstColumn: true })).toBe(
      true,
    )
    expect(optionsOfTableAdd(journal.tableAdds[0]!)).toEqual({
      headerRow: true,
      totalsRow: true,
      bandedRows: true,
      bandedColumns: false,
      firstColumn: true,
      lastColumn: false,
      filterButton: true,
    })
  })
})
