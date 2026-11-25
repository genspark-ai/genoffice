import { describe, expect, it, vi } from 'vitest'
import { RANGE_TYPE, type IRange } from '@univerjs/core'
import { DeltaColumnWidthCommand, SetColWidthCommand } from '@univerjs/sheets'

import { installStreamedColumnWidth } from '../src/renderer/streamed-column-width'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'

type CommandEvent = {
  id: string
  cancel?: boolean
  params: { anchorCol?: number; deltaX?: number; ranges?: IRange[]; value?: number }
}

const wholeColumn = (column: number): IRange => ({
  startRow: 0,
  endRow: 1_000_000,
  startColumn: column,
  endColumn: column,
  rangeType: RANGE_TYPE.COLUMN,
})

function harness(preloadComplete = false) {
  let before = (_event: CommandEvent): void => {}
  const sheet = {
    getSheetId: () => 'sheet-3',
    getVisibleRange: () => ({ startRow: 350, endRow: 390 }),
    getMaxColumns: () => 18,
    getColumnWidth: () => 68,
  }
  const executeCommand = vi.fn().mockResolvedValue(true)
  const runtime = {
    univer: {
      __getInjector: () => ({
        get: () => ({
          getCurrentSelections: () => [wholeColumn(5), wholeColumn(8)].map((range) => ({ range })),
        }),
      }),
    },
    univerAPI: {
      Event: { BeforeCommandExecute: 'before' },
      addEvent: (_event: string, callback: typeof before) => {
        before = callback
        return { dispose() {} }
      },
      getActiveWorkbook: () => ({ getId: () => 'file', getActiveSheet: () => sheet }),
      executeCommand,
    },
  }
  const state = {
    flags: { preloadComplete },
    file: { sheets: [{ id: 'sheet-3' }] },
    loadedRanges: new Map([['sheet-3', { startRow: 300, endRow: 380 }]]),
  }
  installStreamedColumnWidth(runtime as unknown as UniverRuntime, {
    current: state as unknown as LazyWorkbookState,
  })
  return { before, executeCommand }
}

describe('streamed column width', () => {
  it('reissues a column drag with the selected columns and only resident visible rows', async () => {
    const { before, executeCommand } = harness()
    const event: CommandEvent = {
      id: DeltaColumnWidthCommand.id,
      params: { anchorCol: 5, deltaX: 40 },
    }
    before(event)
    await Promise.resolve()
    expect(event.cancel).toBe(true)
    expect(executeCommand).toHaveBeenCalledWith(SetColWidthCommand.id, {
      unitId: 'file',
      subUnitId: 'sheet-3',
      value: 108,
      ranges: [5, 8].map((column) => ({
        startRow: 350,
        endRow: 380,
        startColumn: column,
        endColumn: column,
      })),
    })
  })

  it('limits an explicit width command without treating its selection as a row span', () => {
    const event: CommandEvent = {
      id: SetColWidthCommand.id,
      params: { ranges: [wholeColumn(5)], value: 120 },
    }
    harness().before(event)
    expect(event.params.ranges).toEqual([
      { startRow: 350, endRow: 380, startColumn: 5, endColumn: 5 },
    ])
    expect(event.cancel).toBeUndefined()
  })

  it('lets fully loaded workbooks use Univer’s regular width command', () => {
    const event: CommandEvent = {
      id: SetColWidthCommand.id,
      params: { ranges: [wholeColumn(5)], value: 120 },
    }
    harness(true).before(event)
    expect(event.params.ranges).toEqual([wholeColumn(5)])
  })
})
