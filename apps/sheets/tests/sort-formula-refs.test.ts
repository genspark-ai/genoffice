import {
  ICommandService,
  IUndoRedoService,
  IUniverInstanceService,
  LocaleType,
  LogLevel,
  RedoCommand,
  Univer,
  UndoCommand,
  UniverInstanceType,
} from '@univerjs/core'
import { FUniver } from '@univerjs/core/facade'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import '@univerjs/sheets/facade'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import { FormulaReorderController } from '@univerjs/sheets-formula-ui'
import { afterEach, describe, expect, it } from 'vitest'

import { createEditJournal, recordSetRangeValues } from '../src/renderer/edit-journal'
import { sharedFormulaResolverFor } from '../src/renderer/shared-formula-journal'
import { installSortFormulaRefs } from '../src/renderer/sort-formula-refs'
import { journalRangeSnapshot } from '../src/renderer/univer-sync'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'

/// Sorting A1:D4 by column B ascending: rows (3, 1, 2, 5) become (1, 2, 3, 5),
/// so row 1 takes old row 2, row 2 takes old row 3, row 3 takes old row 1.
const order = { 0: 1, 1: 2, 2: 0 }
const range = { startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 }

let univer: Univer | null = null
afterEach(() => {
  univer?.dispose()
  univer = null
})

function workbook(column: Record<number, { f?: string; si?: string }>) {
  univer = new Univer({ locale: LocaleType.EN_US, logLevel: LogLevel.ERROR })
  univer.registerPlugin(UniverFormulaEnginePlugin)
  univer.registerPlugin(UniverSheetsPlugin)
  univer.registerPlugin(UniverSheetsFormulaPlugin)
  const keys = [3, 1, 2, 5, 7, 8, 9, 10]
  univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
    id: 'wb',
    sheetOrder: ['s1'],
    sheets: {
      s1: {
        id: 's1',
        name: 'Data',
        rowCount: 20,
        columnCount: 10,
        cellData: Object.fromEntries(
          keys.map((key, row) => [
            row,
            { 0: { v: `r${row + 1}` }, 1: { v: key }, 2: { v: 10 }, 3: column[row] ?? {} },
          ]),
        ),
      },
    },
  })
  const api = FUniver.newAPI(univer)
  const injector = (
    univer as unknown as {
      __getInjector(): { add(dependencies: unknown[]): void; get<T>(token: unknown): T }
    }
  ).__getInjector()
  // registered as UniverSheetsFormulaUIPlugin registers it; the plugin creates it only at
  // lifecycle Steady, 3 s after the first render, so nothing here creates it before the install
  injector.add([FormulaReorderController])
  const commands = injector.get<ICommandService>(ICommandService)
  // resolving the undo service registers the undo/redo commands, which act
  // on the focused unit
  injector.get(IUndoRedoService)
  injector.get<IUniverInstanceService>(IUniverInstanceService).focusUnit('wb')
  const runtime = { univer, univerAPI: api } as unknown as UniverRuntime
  installSortFormulaRefs(runtime)
  // the app's journal: a reorder snapshots the sorted range, value mutations
  // record their cells (App.tsx CommandExecuted)
  const state = { editJournal: createEditJournal() } as LazyWorkbookState
  const writes: string[][] = []
  commands.onCommandExecuted((command) => {
    const params = command.params as {
      subUnitId: string
      range: never
      order: never
      cellValue: unknown
      trigger?: string
    }
    if (command.id === 'sheet.mutation.reorder-range') {
      journalRangeSnapshot(runtime, state, params.subUnitId, params.range, params.order)
    } else if (command.id === 'sheet.mutation.set-range-values') {
      const cells = params.cellValue as Record<string, Record<string, unknown>>
      if (params.trigger === 'sheet.command.reorder-range') {
        writes.push(
          Object.entries(cells).flatMap(([row, columns]) =>
            Object.keys(columns).map((column) => `${row}:${column}`),
          ),
        )
      }
      recordSetRangeValues(
        state.editJournal,
        params.subUnitId,
        params.cellValue,
        sharedFormulaResolverFor(runtime, params.subUnitId),
      )
    }
  })
  const sheet = api.getActiveWorkbook()!.getActiveSheet()
  const rows = (count: number) => Array.from({ length: count }, (_, row) => row)
  return {
    sort: () =>
      commands.executeCommand('sheet.command.reorder-range', {
        unitId: 'wb',
        subUnitId: 's1',
        range,
        order,
      }),
    undo: () => commands.executeCommand(UndoCommand.id),
    redo: () => commands.executeCommand(RedoCommand.id),
    formulas: (count = 4) => rows(count).map((row) => sheet.getRange(row, 3).getFormula()),
    labels: () => rows(4).map((row) => sheet.getRange(row, 0).getValue()),
    saved: (count = 4) =>
      rows(count).map((row) => state.editJournal.cells.get('s1')?.get(`${row}:3`)?.formula),
    raw: (row: number) => {
      const cell = sheet.getSheet().getCellRaw(row, 3)
      return { f: cell?.f ?? null, si: cell?.si ?? null }
    },
    writes: () => writes,
  }
}

const sortedRange = [0, 1, 2, 3].flatMap((row) => [0, 1, 2, 3].map((column) => `${row}:${column}`))

describe('sorting rows that carry formulas', () => {
  it('shifts each moved formula to the row it lands on', async () => {
    const wb = workbook({
      0: { f: '=B1*C1' },
      1: { f: '=B2*C2' },
      2: { f: '=B3*C3+$C$1' },
      3: { f: '=B4*C4' },
    })
    await wb.sort()
    expect(wb.labels()).toEqual(['r2', 'r3', 'r1', 'r4'])
    expect(wb.formulas()).toEqual(['=B1*C1', '=B2*C2+$C$1', '=B3*C3', '=B4*C4'])
  })

  it('keeps a shared-formula group pointing at each row after its master moves', async () => {
    const wb = workbook({
      0: { f: '=B1*C1', si: 'g' },
      1: { si: 'g' },
      2: { si: 'g' },
      3: { si: 'g' },
    })
    expect(wb.formulas()).toEqual(['=B1*C1', '=B2*C2', '=B3*C3', '=B4*C4'])
    await wb.sort()
    expect(wb.formulas()).toEqual(['=B1*C1', '=B2*C2', '=B3*C3', '=B4*C4'])
  })

  it('journals the shifted formulas, so the save writes what the sheet shows', async () => {
    const wb = workbook({
      0: { f: '=B1*C1' },
      1: { f: '=B2*C2', si: 'g' },
      2: { si: 'g' },
      3: { si: 'g' },
    })
    await wb.sort()
    expect(wb.saved()).toEqual(['=B1*C1', '=B2*C2', '=B3*C3', '=B4*C4'])
  })

  it('undoes and redoes the sort with its formulas', async () => {
    const wb = workbook({ 0: { f: '=B1*C1' }, 1: { f: '=B2*C2', si: 'g' }, 2: { si: 'g' } })
    const before = wb.formulas()
    await wb.sort()
    const sorted = wb.formulas()
    await wb.undo()
    expect(wb.labels()).toEqual(['r1', 'r2', 'r3', 'r4'])
    expect(wb.formulas()).toEqual(before)
    await wb.redo()
    expect(wb.formulas()).toEqual(sorted)
    expect(sorted).toEqual(['=B1*C1', '=B2*C2', '=B3*C3', ''])
  })

  it('keeps the followers below the sorted rows when the sort moves their master', async () => {
    const wb = workbook({
      0: { f: '=B1*C1', si: 'g' },
      ...Object.fromEntries([1, 2, 3, 4, 5, 6, 7].map((row) => [row, { si: 'g' }])),
    })
    const shown = ['=B1*C1', '=B2*C2', '=B3*C3', '=B4*C4', '=B5*C5', '=B6*C6', '=B7*C7', '=B8*C8']
    expect(wb.formulas(8)).toEqual(shown)
    await wb.sort()
    expect(wb.formulas(8)).toEqual(shown)
    expect([4, 5, 6, 7].map(wb.raw)).toEqual(shown.slice(4).map((f) => ({ f, si: null })))
    expect(wb.saved(8).slice(4)).toEqual(shown.slice(4))
    expect(wb.writes()).toEqual([sortedRange, ['4:3', '5:3', '6:3', '7:3']])
  })

  it('undoes and redoes the sort of a shared group cut by the sorted range', async () => {
    const wb = workbook({
      0: { f: '=B1*C1', si: 'g' },
      ...Object.fromEntries([1, 2, 3, 4, 5, 6, 7].map((row) => [row, { si: 'g' }])),
    })
    const before = wb.formulas(8)
    await wb.sort()
    const sorted = wb.formulas(8)
    await wb.undo()
    expect(wb.labels()).toEqual(['r1', 'r2', 'r3', 'r4'])
    expect(wb.formulas(8)).toEqual(before)
    expect(wb.raw(0)).toEqual({ f: '=B1*C1', si: 'g' })
    expect([4, 5, 6, 7].map(wb.raw)).toEqual([4, 5, 6, 7].map(() => ({ f: null, si: 'g' })))
    await wb.redo()
    expect(wb.formulas(8)).toEqual(sorted)
    expect(sorted).toEqual(before)
  })

  it('adds no write to a sort that cuts no shared group', async () => {
    const wb = workbook(
      Object.fromEntries(
        [0, 1, 2, 3, 4, 5, 6, 7].map((row) => [row, { f: `=B${row + 1}*C${row + 1}` }]),
      ),
    )
    await wb.sort()
    expect(wb.formulas(8)).toEqual(
      [0, 1, 2, 3, 4, 5, 6, 7].map((row) => `=B${row + 1}*C${row + 1}`),
    )
    expect(wb.writes()).toEqual([sortedRange])
  })
})
