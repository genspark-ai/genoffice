/**
 * End-to-end engine check for the prod_013 incident: `='Medição 06'!B3`
 * where B3 is a style-only cell (`<c r="B3" s="170"/>`). The patch pipeline
 * installs such cells as `{ v: null }` (never `{ v: '' }` — an empty STRING
 * makes the reference return '' where Excel returns 0), and the formula
 * engine then evaluates references to them as 0, so a serial-0 date shows
 * 1900/1/0 and a plain 0 shows 0 instead of a blank cell.
 */
import {
  ICommandService,
  IUniverInstanceService,
  LocaleType,
  LogLevel,
  Univer,
  UniverInstanceType,
} from '@univerjs/core'
import { IFormulaRuntimeService, UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import { describe, expect, it } from 'vitest'

import { coerceNullResult } from '../src/renderer/formula-null-result'

describe('references to style-only cells', () => {
  it('evaluate to 0 like Excel, not blank', async () => {
    const univer = new Univer({ logLevel: LogLevel.ERROR, locale: LocaleType.EN_US, locales: {} })
    univer.registerPlugin(UniverFormulaEnginePlugin)
    univer.registerPlugin(UniverSheetsPlugin)
    univer.registerPlugin(UniverSheetsFormulaPlugin)
    const injector = univer.__getInjector()
    const runtimeService = injector.get(IFormulaRuntimeService) as unknown as {
      setRuntimeData: (variant: unknown) => unknown
    }
    const original = runtimeService.setRuntimeData.bind(runtimeService)
    runtimeService.setRuntimeData = (variant: unknown) =>
      original(coerceNullResult(variant) as never)
    univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
      id: 'wb1',
      sheetOrder: ['s1', 's2'],
      name: 'wb',
      styles: {},
      sheets: {
        s1: { id: 's1', name: 'RESUMO', rowCount: 10, columnCount: 10, cellData: {} },
        s2: { id: 's2', name: 'Med 06', rowCount: 10, columnCount: 10, cellData: {} },
      },
    })
    const commandService = injector.get(ICommandService)
    // Style-only cells install with v: null (the patch pipeline's shape).
    await commandService.executeCommand('sheet.command.set-range-values', {
      unitId: 'wb1',
      subUnitId: 's2',
      range: { startRow: 2, endRow: 2, startColumn: 1, endColumn: 3 },
      value: {
        2: {
          1: { v: null, s: { bg: { rgb: '#ff0000' } } },
          3: { v: null, s: { bg: { rgb: '#ff0000' } } },
        },
      },
    })
    await commandService.executeCommand('sheet.command.set-range-values', {
      unitId: 'wb1',
      subUnitId: 's1',
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
      value: {
        0: {
          0: { f: "='Med 06'!B3", v: 0 },
          1: { f: "='Med 06'!D3", v: 0 },
        },
      },
    })
    await new Promise((resolve) => setTimeout(resolve, 2000))
    const instanceService = injector.get(IUniverInstanceService)
    const workbook = instanceService.getUnit('wb1') as unknown as {
      getSheetBySheetId(id: string): { getCellRaw(row: number, column: number): { v?: unknown } }
    }
    const resumo = workbook.getSheetBySheetId('s1')
    expect(resumo.getCellRaw(0, 0)?.v).toBe(0)
    expect(resumo.getCellRaw(0, 1)?.v).toBe(0)
    // The style-only install left the referenced cell value-less.
    const med = workbook.getSheetBySheetId('s2')
    expect(med.getCellRaw(2, 1)).not.toHaveProperty('v')
    univer.dispose()
  }, 20_000)
})
