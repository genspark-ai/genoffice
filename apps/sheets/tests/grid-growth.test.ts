// @vitest-environment jsdom
import { CommandType, ICommandService, IUndoRedoService, LogLevel, Univer } from '@univerjs/core'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import '@univerjs/sheets/facade'
import { FUniver } from '@univerjs/core/lib/facade'
import { describe, expect, it } from 'vitest'

import { installGridGrowth, nextGridGrowth } from '../src/renderer/grid-growth'
import type { UniverRuntime } from '../src/renderer/univer-state'

/// A blank workbook's grid, and a viewport parked at A1.
const BLANK = { startColumn: 0, startRow: 0, columnCount: 26, rowCount: 1000 } as const

describe('nextGridGrowth', () => {
  it('leaves a blank sheet alone until the viewport nears an edge', () => {
    // column 0 + 20 columns of lookahead < 26, so opening a blank sheet must
    // not immediately grow anything
    expect(nextGridGrowth(BLANK)).toBeNull()
  })

  it('extends the columns when the viewport reaches the right edge', () => {
    // column 6 leaves 20 columns of lookahead, which is the edge
    expect(nextGridGrowth({ ...BLANK, startColumn: 6 })).toEqual({
      columnCount: 52,
      rowCount: 1000,
    })
  })

  it('extends the rows when the viewport reaches the bottom edge', () => {
    expect(nextGridGrowth({ ...BLANK, startRow: 800 })).toEqual({
      columnCount: 26,
      rowCount: 1200,
    })
  })

  it('grows only the axis that reached its edge', () => {
    const growth = nextGridGrowth({ ...BLANK, startColumn: 6 })
    expect(growth?.rowCount).toBe(1000)
    expect(growth?.columnCount).toBe(52)
  })

  it('is self-limiting: the size it returns needs no further growth', () => {
    // the scroll handler re-runs on every scroll state, including the one the
    // growth itself provokes. If the result still asked to grow, one scroll
    // would cascade into a runaway extension
    const first = nextGridGrowth({ ...BLANK, startColumn: 6 })!
    expect(
      nextGridGrowth({
        startColumn: 6,
        startRow: 0,
        columnCount: first.columnCount,
        rowCount: first.rowCount,
      }),
    ).toBeNull()
  })

  it('does not grow a grid that already fills the sheet', () => {
    expect(
      nextGridGrowth({
        startColumn: 16_300,
        startRow: 1_048_000,
        columnCount: 16_384,
        rowCount: 1_048_576,
      }),
    ).toBeNull()
  })

  it('stops at the .xlsx ceiling instead of growing past it', () => {
    // 100 columns short of the last column: one step overshoots, and the grid
    // must land exactly on 16384 rather than 16410
    const growth = nextGridGrowth({
      ...BLANK,
      startColumn: 16_364,
      columnCount: 16_384,
      rowCount: 1_048_576,
    })
    expect(growth).toBeNull()

    const near = nextGridGrowth({
      ...BLANK,
      startColumn: 16_350,
      columnCount: 16_370,
      rowCount: 1_048_576,
    })
    expect(near?.columnCount).toBe(16_384)
  })

  it('stops at the .xlsx row ceiling too', () => {
    // 76 rows short of the last row: one step of 200 overshoots, so the row
    // clamp has to land exactly on 1048576
    const near = nextGridGrowth({
      startColumn: 0,
      startRow: 1_048_400,
      columnCount: 26,
      rowCount: 1_048_500,
    })
    expect(near?.rowCount).toBe(1_048_576)
  })

  it('keeps scrolling right working past the original 26 columns', () => {
    // the reported symptom: a blank sheet stops at Z. Walk the viewport the
    // way a scrollbar drag does — the grid must extend ahead of the cursor
    let edge = { startColumn: 0, startRow: 0, columnCount: 26, rowCount: 1000 }
    for (let column = 0; column < 400; column += 5) {
      const growth = nextGridGrowth({ ...edge, startColumn: column })
      if (growth) {
        edge = { ...edge, columnCount: growth.columnCount, rowCount: growth.rowCount }
      }
    }
    expect(edge.columnCount).toBeGreaterThan(26)
  })
})

/// The review's defect: growth ran SetWorksheetColumnCountCommand, whose handler
/// calls undoRedoService.pushUndoRedo, so scrolling a pristine blank sheet lit
/// the Undo button and the first Ctrl+Z after typing at AA1 shrank the grid back
/// over the cell just edited instead of undoing the edit. Scrolling is a view
/// action, so the undo stack must not move.
///
/// These run against a real Univer: a real UniverSheetsPlugin, a real workbook,
/// and the real LocalUndoRedoService. Only the scroll command is stood in for,
/// because the real one lives in @univerjs/sheets-ui and needs a DOM to boot.
const UNIT_ID = 'grid-growth-undo-wb'
const SHEET_ID = 'grid-growth-undo-sheet'
const BLANK_SHEET_SIZE = { rowCount: 1000, columnCount: 26 }

function bootBlankWorkbook() {
  const univer = new Univer({ logLevel: LogLevel.SILENT })
  univer.registerPlugin(UniverSheetsPlugin)
  const univerAPI = FUniver.newAPI(univer)
  univerAPI.createWorkbook({
    id: UNIT_ID,
    name: 'blank',
    sheetOrder: [SHEET_ID],
    sheets: { [SHEET_ID]: { id: SHEET_ID, name: 'Sheet1', ...BLANK_SHEET_SIZE } },
  })
  const injector = univer.__getInjector()
  const commandService = injector.get(ICommandService)
  commandService.registerCommand({
    id: 'sheet.operation.set-scroll',
    type: CommandType.COMMAND,
    handler: () => true,
  })
  return {
    univer,
    univerAPI,
    commandService,
    undoRedoService: injector.get(IUndoRedoService),
    worksheet: univerAPI.getActiveWorkbook()!.getActiveSheet()!,
  }
}

/**
 * The undo stack's length for one workbook. Univer keeps it in LocalUndoRedoService
 * as a per-unit array; undoRedoStatus$ would also do, but it only recomputes on a
 * focus change, which a headless test never triggers.
 */
function undoStackLength(undoRedoService: IUndoRedoService): number {
  const stacks = (undoRedoService as unknown as { _undoStacks: Map<string, unknown[]> })._undoStacks
  return stacks.get(UNIT_ID)?.length ?? 0
}

describe('grid growth and the undo stack', () => {
  it('grows the grid on a scroll without touching the undo stack', () => {
    const { univer, univerAPI, commandService, undoRedoService, worksheet } = bootBlankWorkbook()
    const dispose = installGridGrowth({ univer, univerAPI } as unknown as UniverRuntime)
    try {
      // a pristine sheet: nothing to undo, so the Undo button starts disabled
      expect(undoStackLength(undoRedoService)).toBe(0)

      // scroll past column F and down to the row edge in one viewport move
      commandService.syncExecuteCommand('sheet.operation.set-scroll', {
        unitId: UNIT_ID,
        subUnitId: SHEET_ID,
        sheetViewStartColumn: 6,
        sheetViewStartRow: 800,
      })

      // the grid really did grow, so the dispatch was not a silent no-op
      expect(worksheet.getMaxColumns()).toBe(52)
      expect(worksheet.getMaxRows()).toBe(1200)
      expect(undoStackLength(undoRedoService)).toBe(0)
    } finally {
      dispose()
    }
  })

  it('keeps the stack flat across a long scroll that grows repeatedly', () => {
    const { univer, univerAPI, commandService, undoRedoService, worksheet } = bootBlankWorkbook()
    const dispose = installGridGrowth({ univer, univerAPI } as unknown as UniverRuntime)
    try {
      // one lookahead step per extension, the way a scrollbar drag arrives
      for (let column = 5; column <= 200; column += 5) {
        commandService.syncExecuteCommand('sheet.operation.set-scroll', {
          unitId: UNIT_ID,
          subUnitId: SHEET_ID,
          sheetViewStartColumn: column,
          sheetViewStartRow: 0,
        })
      }
      expect(worksheet.getMaxColumns()).toBeGreaterThan(200)
      // each step would have added one entry through the command variant
      expect(undoStackLength(undoRedoService)).toBe(0)
    } finally {
      dispose()
    }
  })

  it('control: the same growth through the facade does push an undo step', () => {
    // Guards the two tests above: if this ever stops growing the stack, the
    // assertions there are measuring nothing and have to be rewritten.
    const { univerAPI, undoRedoService, worksheet } = bootBlankWorkbook()
    expect(undoStackLength(undoRedoService)).toBe(0)

    worksheet.setColumnCount(52)
    expect(worksheet.getMaxColumns()).toBe(52)
    expect(undoStackLength(undoRedoService)).toBe(1)
  })
})
