import { describe, expect, it } from 'vitest'

import { ICommandService, IUndoRedoService } from '@univerjs/core'

import { createEditJournal, journalSize } from '../src/renderer/edit-journal'
import { handleRibbonCommand, type RibbonCommandContext } from '../src/renderer/ribbon-actions'

/// Enough of the facade + injector for the fill ribbon path: the active
/// range, the set-style command, and the undo stack the column-default
/// bookkeeping attaches to.
function makeHarness(selection: {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}) {
  const commands: { id: string; params: unknown }[] = []
  const messages: string[] = []
  const pending: number[] = []
  const undoStack: { unitID: string; undoMutations: unknown[]; redoMutations: unknown[] }[] = []
  const height = selection.endRow - selection.startRow + 1
  const width = selection.endColumn - selection.startColumn + 1
  const range = {
    getRange: () => selection,
    getRow: () => selection.startRow,
    getColumn: () => selection.startColumn,
    getHeight: () => height,
    getWidth: () => width,
    getCellStyleData: () => ({}),
    setBackground: (color: string) => {
      commands.push({ id: 'facade.setBackground', params: color })
      undoStack.push({ unitID: 'wb', undoMutations: [], redoMutations: [] })
    },
  }
  const worksheet = { getSheetId: () => 'sheet-1', getMaxRows: () => 1000 }
  const workbook = {
    getId: () => 'wb',
    getActiveSheet: () => worksheet,
    getActiveRange: () => range,
  }
  const univerAPI = {
    getActiveWorkbook: () => workbook,
    syncExecuteCommand: (id: string, params: unknown) => {
      commands.push({ id, params })
      undoStack.push({ unitID: 'wb', undoMutations: [], redoMutations: [] })
      return true
    },
    executeCommand: async (id: string, params: unknown) => {
      commands.push({ id, params })
      return true
    },
  }
  const services = new Map<unknown, unknown>([
    [ICommandService, { registerCommand: () => ({ dispose() {} }) }],
    [
      IUndoRedoService,
      {
        pitchTopUndoElement: () => undoStack[undoStack.length - 1] ?? null,
        pushUndoRedo: (item: (typeof undoStack)[number]) => undoStack.push(item),
      },
    ],
  ])
  const editJournal = createEditJournal()
  const ctx = {
    univerRef: {
      current: {
        univerAPI,
        univer: { __getInjector: () => ({ get: (token: unknown) => services.get(token) }) },
      },
    },
    lazyWorkbookRef: { current: { editJournal } },
    setMessage: (message: string) => messages.push(message),
    setPendingEdits: (count: number) => pending.push(count),
  } as unknown as RibbonCommandContext
  return { ctx, commands, messages, pending, editJournal, undoStack }
}

const WHOLE_COLUMN = { startRow: 0, endRow: 999, startColumn: 2, endColumn: 2 }
const SOME_CELLS = { startRow: 3, endRow: 8, startColumn: 2, endColumn: 2 }

describe('ribbon No Fill', () => {
  it('pins the clear with the empty-rgb sentinel instead of a strippable null', () => {
    // setBackground(null) becomes bg: { rgb: null } in the mutation, which
    // removeNull deletes — leaving a MISSING bg through which a <col style=>
    // fill composes right back onto the cells the user just cleared.
    const { ctx, commands, messages } = makeHarness(SOME_CELLS)
    handleRibbonCommand(ctx, 'fill:none')
    expect(commands).toEqual([
      {
        id: 'sheet.command.set-style',
        params: {
          unitId: 'wb',
          subUnitId: 'sheet-1',
          range: SOME_CELLS,
          style: { type: 'bg', value: { rgb: '' } },
        },
      },
    ])
    expect(messages.some((message) => /fail|error/i.test(message))).toBe(false)
  })

  it('journals a column-default fill clear for a full-column selection', () => {
    // Excel stores a whole-column fill as <col style=>; clearing only the
    // cells inside Univer's grid would leave rows below it green on reopen.
    const { ctx, editJournal, pending, undoStack } = makeHarness(WHOLE_COLUMN)
    handleRibbonCommand(ctx, 'fill:none')
    expect(editJournal.structuralOps.get('sheet-1')).toEqual([
      { kind: 'set-col-style', start: 2, end: 2, style: { fillColor: null } },
    ])
    expect(journalSize(editJournal)).toBe(1)
    expect(pending).toEqual([1])
    // The bookkeeping rides on the set-style undo entry: one ⌘Z reverts both.
    expect(undoStack).toHaveLength(1)
    expect(undoStack[0]?.undoMutations).toHaveLength(1)
  })

  it('journals the column-default fill when a full column is filled', () => {
    const { ctx, editJournal, commands } = makeHarness(WHOLE_COLUMN)
    handleRibbonCommand(ctx, 'fill:#00aa55')
    expect(commands).toEqual([{ id: 'facade.setBackground', params: '#00aa55' }])
    expect(editJournal.structuralOps.get('sheet-1')).toEqual([
      { kind: 'set-col-style', start: 2, end: 2, style: { fillColor: '#00AA55' } },
    ])
  })

  it('leaves the column default alone for a partial selection', () => {
    const { ctx, editJournal, pending } = makeHarness(SOME_CELLS)
    handleRibbonCommand(ctx, 'fill:none')
    handleRibbonCommand(ctx, 'fill:#00aa55')
    expect(editJournal.structuralOps.get('sheet-1')).toBeUndefined()
    // Cell-level deltas arrive through the set-range-values journal listener
    // (App.tsx), not through the ribbon dispatcher.
    expect(pending).toEqual([])
  })
})
