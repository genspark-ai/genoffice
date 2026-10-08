import { LocaleType, LogLevel, Univer, UniverInstanceType } from '@univerjs/core'
import { FUniver } from '@univerjs/core/facade'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import '@univerjs/sheets/facade'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import '@univerjs/sheets-formula/facade'
import { UniverSheetsTablePlugin } from '@univerjs/sheets-table'
import '@univerjs/sheets-table/facade'
import tableLocale from '@univerjs/sheets-table/locale/en-US'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { planFileTableRegistrations } from '../src/renderer/file-tables'
import type { WorkbookFile } from '../src/shared/desktop-api'

type FileSheet = WorkbookFile['sheets'][number]

function cellData(rows: (string | number)[][]) {
  return Object.fromEntries(
    rows.map((row, r) => [
      r,
      Object.fromEntries(
        row.map((v, c) => [c, typeof v === 'string' && v.startsWith('=') ? { f: v } : { v }]),
      ),
    ]),
  )
}

/// Two file tables on different sheets with identical column names; the
/// first carries a totals row.
const fileSheets = [
  {
    id: 's1',
    name: 'Data',
    tables: [
      {
        range: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 1 },
        headerRowCount: 1,
        totalsRowCount: 1,
        showRowStripes: true,
        showColumnStripes: false,
        name: 'Sales',
        columns: ['Item', 'Amount'],
      },
    ],
  },
  {
    id: 's2',
    name: 'Other',
    tables: [
      {
        range: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
        headerRowCount: 1,
        showRowStripes: true,
        showColumnStripes: false,
        name: 'Costs',
        columns: ['Item', 'Amount'],
      },
    ],
  },
] as unknown as FileSheet[]

const formulas: [sheetId: string, cell: string, formula: string, expected: unknown][] = [
  ['s1', 'E1', '=SUM(Sales[Amount])', 6],
  ['s1', 'E2', '=SUM(Costs[Amount])', 30],
  ['s1', 'E3', '=COUNTA(Sales[#All])', 8],
  ['s1', 'E4', '=SUM(Sales[[#Headers],[Amount]])', 0],
  ['s1', 'E5', '=SUM(Sales[[#Data],[Item]:[Amount]])', 6],
  ['s1', 'D2', '=Sales[[#This Row],[Amount]]*2', 2],
  ['s2', 'E1', '=SUM(Sales[Amount])+SUM(Costs[Amount])', 36],
  ['s1', 'E6', '=SUM(Nope[Amount])', '#NAME?'],
]

describe('structured references resolve against file table displayNames', () => {
  let univer: Univer
  let api: FUniver

  beforeAll(async () => {
    univer = new Univer({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: tableLocale },
      logLevel: LogLevel.ERROR,
    })
    univer.registerPlugin(UniverFormulaEnginePlugin)
    univer.registerPlugin(UniverSheetsPlugin)
    univer.registerPlugin(UniverSheetsFormulaPlugin)
    univer.registerPlugin(UniverSheetsTablePlugin)
    univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
      id: 'wb',
      sheetOrder: ['s1', 's2'],
      sheets: {
        s1: {
          id: 's1',
          name: 'Data',
          rowCount: 20,
          columnCount: 10,
          cellData: cellData([
            ['Item', 'Amount'],
            ['a', 1],
            ['b', 2],
            ['c', 3],
            ['Total', 6],
          ]),
        },
        s2: {
          id: 's2',
          name: 'Other',
          rowCount: 20,
          columnCount: 10,
          cellData: cellData([
            ['Item', 'Amount'],
            ['x', 10],
            ['y', 20],
          ]),
        },
      },
    })
    api = FUniver.newAPI(univer)
    const workbook = api.getActiveWorkbook()!
    for (const entry of planFileTableRegistrations(fileSheets)) {
      const ws = workbook.getSheetBySheetId(entry.sheetId)!
      const added = await ws.addTable(entry.tableName, entry.range, entry.tableId, {
        columns: entry.columns!.map((name, index) => ({
          id: `${entry.tableId}-col-${index}`,
          displayName: name,
        })),
      } as never)
      expect(added).toBe(true)
    }
    for (const [sheetId, cell, formula] of formulas) {
      workbook.getSheetBySheetId(sheetId)!.getRange(cell).setFormula(formula)
    }
    await new Promise((resolve) => setTimeout(resolve, 1500))
  }, 30_000)

  afterAll(() => {
    univer.dispose()
  })

  it.each(formulas)('%s!%s %s', (sheetId, cell, _formula, expected) => {
    const value = api
      .getActiveWorkbook()!
      .getSheetBySheetId(sheetId)!
      .getRange(cell)
      .getCellData()?.v
    expect(value).toBe(expected)
  })
})
