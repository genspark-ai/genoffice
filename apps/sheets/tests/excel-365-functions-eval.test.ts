import { CellValueType, LocaleType, LogLevel, Univer, UniverInstanceType } from '@univerjs/core'
import { FUniver } from '@univerjs/core/facade'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import '@univerjs/sheets/facade'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import '@univerjs/sheets-formula/facade'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

import { installExcel365Functions } from '../src/renderer/excel-365-functions'
import type { UniverRuntime } from '../src/renderer/univer-state'

const inputs: Record<string, string | number> = {
  A1: 'Order 1234',
  A2: 'Ana Silva',
  A3: 'Due 2026-11-13',
  A4: '1-2-3',
  A5: 10,
  A6: 30,
  A7: 60,
  A8: 'Apple',
}

const formulas: [cell: string, formula: string, expected: unknown][] = [
  ['B1', '=REGEXTEST(A1,"[0-9]+")', true],
  ['B2', '=REGEXEXTRACT(A2,"([A-Z])[a-z]+")', 'Ana'],
  ['B3', '=REGEXEXTRACT(A3,"(\\d{4})-(\\d{2})-(\\d{2})",2)', '2026'],
  ['B4', '=REGEXREPLACE(A4,"[0-9]","x",-1)', '1-2-x'],
  ['B5', '=PERCENTOF(A5:A6,A5:A7)', 0.4],
  ['B8', '=REGEXTEST(A8,"apple",1)', true],
  ['B9', '=IF(REGEXTEST(A1,"^Order"),"order","other")', 'order'],
]

describe('Microsoft 365 functions evaluate in the formula engine', () => {
  let univer: Univer
  let api: FUniver
  let installed: { dispose(): void }

  beforeAll(async () => {
    univer = new Univer({ locale: LocaleType.EN_US, locales: {}, logLevel: LogLevel.ERROR })
    univer.registerPlugin(UniverFormulaEnginePlugin)
    univer.registerPlugin(UniverSheetsPlugin)
    univer.registerPlugin(UniverSheetsFormulaPlugin)
    univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
      id: 'wb',
      sheetOrder: ['s1'],
      sheets: { s1: { id: 's1', name: 'Sheet1', rowCount: 20, columnCount: 10, cellData: {} } },
    })
    installed = installExcel365Functions({ univer } as unknown as UniverRuntime)
    await new Promise((resolve) => setTimeout(resolve, 0))
    api = FUniver.newAPI(univer)
    const sheet = api.getActiveWorkbook()!.getActiveSheet()
    for (const [cell, value] of Object.entries(inputs)) sheet.getRange(cell).setValue(value)
    for (const [cell, formula] of formulas) sheet.getRange(cell).setFormula(formula)
    await new Promise((resolve) => setTimeout(resolve, 1500))
  }, 30_000)

  afterAll(() => {
    installed.dispose()
    univer.dispose()
  })

  const valueAt = (cell: string) =>
    api.getActiveWorkbook()!.getActiveSheet().getRange(cell).getCellData()?.v

  it.each(formulas)('%s %s', (cell, _formula, expected) => {
    const value = valueAt(cell)
    if (typeof expected === 'number') expect(value).toBeCloseTo(expected)
    else if (typeof expected === 'boolean') {
      // booleans are stored as 1/0 with a boolean cell type
      const data = api.getActiveWorkbook()!.getActiveSheet().getRange(cell).getCellData()
      expect([data?.t, data?.v]).toEqual([CellValueType.BOOLEAN, expected ? 1 : 0])
    } else expect(value).toBe(expected)
  })

  it('spills the capture groups of return_mode 2 across the row', () => {
    expect([valueAt('B3'), valueAt('C3'), valueAt('D3')]).toEqual(['2026', '11', '13'])
  })
})
