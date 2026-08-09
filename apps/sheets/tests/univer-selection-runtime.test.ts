import { ICommandService, IUniverInstanceService, LocaleType } from '@univerjs/core'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import '@univerjs/sheets/lib/facade'
import {
  generateNullCellValue,
  SheetSkeletonService,
  SheetsSelectionsService,
  SetRangeValuesMutation,
  SetRangeValuesUndoMutationFactory,
  UniverSheetsPlugin,
} from '@univerjs/sheets'
import { describe, expect, it, vi } from 'vitest'

import { createUniver } from '../src/renderer/create-univer'

type Range = { startRow: number; endRow: number; startColumn: number; endColumn: number }

function createRuntime(ranges: readonly Range[]) {
  const runtime = createUniver({
    locale: LocaleType.EN_US,
    locales: { [LocaleType.EN_US]: UniverPresetSheetsCoreEnUS },
    presets: [],
    plugins: [UniverSheetsPlugin],
  })
  const workbook = runtime.univerAPI.createWorkbook({
    id: 'keyboard-test',
    name: 'Keyboard test',
    sheetOrder: ['sheet-1'],
    sheets: {
      'sheet-1': {
        id: 'sheet-1',
        name: 'Sheet1',
        cellData: {
          0: {
            0: { v: 'A', s: { bd: { t: { s: 1, cl: { rgb: '#005a9c' } } } } },
            1: { v: 'B', s: { bd: { t: { s: 1, cl: { rgb: '#005a9c' } } } } },
            2: { v: 'C', s: { bd: { t: { s: 1, cl: { rgb: '#005a9c' } } } } },
            3: { v: 'D', s: { bd: { t: { s: 1, cl: { rgb: '#005a9c' } } } } },
            4: { v: 'E', s: { bd: { t: { s: 1, cl: { rgb: '#005a9c' } } } } },
          },
          2: { 2: { v: 'F' }, 3: { v: 'G' } },
        },
      },
    },
  })
  const injector = runtime.univer.__getInjector()
  injector.get(IUniverInstanceService).setCurrentUnitForType('keyboard-test')
  const worksheet = injector
    .get(IUniverInstanceService)
    .getUniverSheetInstance('keyboard-test')!
    .getSheetBySheetId('sheet-1')!
  const selectionService = injector.get(SheetsSelectionsService)
  selectionService.setSelections(
    'keyboard-test',
    'sheet-1',
    ranges.map((range) => ({
      range,
      primary: {
        ...range,
        actualRow: range.startRow,
        actualColumn: range.startColumn,
        isMerged: false,
        isMergedMainCell: false,
      },
    })),
  )
  // The production UI supplies a real visible skeleton. This minimal viewport
  // keeps the production command's auto-height branch deterministic in Node
  // while all workbook, selection, clear, and undo/redo services are real.
  vi.spyOn(injector.get(SheetSkeletonService), 'getSkeleton').mockReturnValue({
    worksheet,
    scrollY: 0,
    getOffsetRelativeToRowCol: () => ({ row: 0 }),
  } as never)
  return {
    runtime,
    injector,
    worksheet,
    selectionService,
    commandService: injector.get(ICommandService),
  }
}

function valueAt(
  worksheet: ReturnType<typeof createRuntime>['worksheet'],
  row: number,
  col: number,
) {
  return worksheet.getCellRaw(row, col)?.v
}

describe('Univer clear-selection-content runtime integration', () => {
  it('clears A1:E1 through the real selection service and keeps borders through undo/redo', async () => {
    const runtime = createRuntime([{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 4 }])
    try {
      expect(runtime.selectionService.getCurrentSelections().map(({ range }) => range)).toEqual([
        { startRow: 0, endRow: 0, startColumn: 0, endColumn: 4 },
      ])
      const borderBefore = runtime.worksheet.getCellRaw(0, 0)?.s

      const clearParams = {
        unitId: 'keyboard-test',
        subUnitId: 'sheet-1',
        cellValue: generateNullCellValue(
          runtime.selectionService.getCurrentSelections().map(({ range }) => range),
        ),
      }
      const undoParams = SetRangeValuesUndoMutationFactory(runtime.injector, clearParams)
      await runtime.commandService.executeCommand(SetRangeValuesMutation.id, clearParams)
      expect([0, 1, 2, 3, 4].map((column) => valueAt(runtime.worksheet, 0, column))).toEqual([
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
      ])
      expect(runtime.worksheet.getCellRaw(0, 0)?.s).toEqual(borderBefore)

      await runtime.commandService.executeCommand(SetRangeValuesMutation.id, undoParams)
      expect([0, 1, 2, 3, 4].map((column) => valueAt(runtime.worksheet, 0, column))).toEqual([
        'A',
        'B',
        'C',
        'D',
        'E',
      ])
      await runtime.commandService.executeCommand(SetRangeValuesMutation.id, clearParams)
      expect([0, 1, 2, 3, 4].map((column) => valueAt(runtime.worksheet, 0, column))).toEqual([
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
      ])
    } finally {
      runtime.runtime.univer.dispose()
    }
  })

  it('clears all Ctrl-disjoint selections in one real command', async () => {
    const runtime = createRuntime([
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
      { startRow: 2, endRow: 2, startColumn: 2, endColumn: 3 },
    ])
    try {
      expect(runtime.selectionService.getCurrentSelections()).toHaveLength(2)
      await runtime.commandService.executeCommand('sheet.command.clear-selection-content')
      expect(valueAt(runtime.worksheet, 0, 0)).toBeUndefined()
      expect(valueAt(runtime.worksheet, 0, 1)).toBeUndefined()
      expect(valueAt(runtime.worksheet, 0, 2)).toBe('C')
      expect(valueAt(runtime.worksheet, 2, 2)).toBeUndefined()
      expect(valueAt(runtime.worksheet, 2, 3)).toBeUndefined()
    } finally {
      runtime.runtime.univer.dispose()
    }
  })
})
