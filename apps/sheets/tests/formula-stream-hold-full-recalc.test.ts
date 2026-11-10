/**
 * Streamed file loads can leave a formula's dependents stale: a cycle that
 * ran before a referenced sheet landed stores its precedent as 0, and a
 * later interrupted cycle never revisits the dependent (SUMIFS
 * keyed on `=Config!B6` showed 0.0 while the criteria cell showed its text).
 * The preload asks the stream hold for one forced, dependency-ordered pass
 * once every sheet is in; this drives the hold headlessly with a stale
 * dependent produced by an untracked mutation.
 */
import {
  ICommandService,
  IUniverInstanceService,
  LocaleType,
  LogLevel,
  Univer,
  UniverInstanceType,
} from '@univerjs/core'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import { describe, expect, it } from 'vitest'

import {
  FORMULA_STREAM_HOLD_MS,
  installFormulaStreamHold,
  noteFormulaStreamChunk,
  requestFullRecalcAfterStream,
} from '../src/renderer/formula-stream-hold'

const START_MUTATION = 'formula.mutation.set-formula-calculation-start'

describe('requestFullRecalcAfterStream', () => {
  it('recomputes dependents left stale by earlier cycles in one forced pass', async () => {
    const univer = new Univer({ logLevel: LogLevel.ERROR, locale: LocaleType.EN_US, locales: {} })
    univer.registerPlugin(UniverFormulaEnginePlugin)
    univer.registerPlugin(UniverSheetsPlugin)
    univer.registerPlugin(UniverSheetsFormulaPlugin)
    const injector = univer.__getInjector()
    const runtime = { univer, univerAPI: null as never }
    installFormulaStreamHold(runtime as never)
    univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
      id: 'wb1',
      sheetOrder: ['dash', 'data', 'config'],
      name: 'wb',
      styles: {},
      sheets: {
        dash: { id: 'dash', name: 'Dash', rowCount: 10, columnCount: 5, cellData: {} },
        data: { id: 'data', name: 'Data', rowCount: 20, columnCount: 15, cellData: {} },
        config: { id: 'config', name: 'Config', rowCount: 10, columnCount: 5, cellData: {} },
      },
    })
    const commandService = injector.get(ICommandService)
    const forcedStarts: boolean[] = []
    commandService.onCommandExecuted((command) => {
      if (command.id !== START_MUTATION) return
      forcedStarts.push(
        (command.params as { forceCalculation?: boolean }).forceCalculation === true,
      )
    })
    const setValues = (subUnitId: string, value: Record<number, Record<number, unknown>>) =>
      commandService.executeCommand('sheet.command.set-range-values', {
        unitId: 'wb1',
        subUnitId,
        range: { startRow: 0, endRow: 6, startColumn: 0, endColumn: 12 },
        value,
      })
    const settle = () => new Promise((resolve) => setTimeout(resolve, FORMULA_STREAM_HOLD_MS + 900))

    const names = ['kim', 'park', 'park', 'lee', 'park']
    const dataRows: Record<number, Record<number, unknown>> = {}
    names.forEach((name, index) => {
      dataRows[index + 1] = { 5: { v: index + 1 }, 11: { v: name, t: 1 } }
    })
    await setValues('data', dataRows)
    noteFormulaStreamChunk()
    await setValues('dash', {
      0: { 0: { f: '=Config!B1' }, 1: { f: '=SUMIFS(Data!F:F,Data!L:L,A1)' } },
    })
    noteFormulaStreamChunk()
    await settle()
    const workbook = injector.get(IUniverInstanceService).getUnit('wb1') as unknown as {
      getSheetBySheetId(id: string): { getCellRaw(row: number, column: number): { v?: unknown } }
    }
    const dash = workbook.getSheetBySheetId('dash')
    // Config is still empty: the lookup and the SUMIFS keyed on it are 0.
    expect(dash.getCellRaw(0, 0)?.v).toBe(0)
    expect(dash.getCellRaw(0, 1)?.v).toBe(0)

    // The precedent changes under the engine's feet (onlyLocal mutations
    // bypass dirty tracking, standing in for a cycle interrupted mid-way).
    await commandService.executeCommand(
      'sheet.mutation.set-range-values',
      {
        unitId: 'wb1',
        subUnitId: 'config',
        cellValue: { 0: { 1: { v: 'park', t: 1 } } },
      },
      { onlyLocal: true },
    )
    await settle()
    expect(dash.getCellRaw(0, 1)?.v).toBe(0)

    noteFormulaStreamChunk()
    requestFullRecalcAfterStream()
    await settle()
    expect(forcedStarts.at(-1)).toBe(true)
    expect(dash.getCellRaw(0, 0)?.v).toBe('park')
    expect(dash.getCellRaw(0, 1)?.v).toBe(10)
    univer.dispose()
  }, 20_000)
})
