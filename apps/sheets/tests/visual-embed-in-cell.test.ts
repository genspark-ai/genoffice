import { describe, expect, it } from 'vitest'

import { ICommandService, IUndoRedoService } from '@univerjs/core'

import type { WorkbookVisualObject } from '../src/shared/desktop-api'
import { createEditJournal, journalSize, recordVisualAdd } from '../src/renderer/edit-journal'
import { mergeCellImages } from '../src/renderer/univer-sync'
import {
  handleVisualCommand,
  type VisualArrangeContext,
} from '../src/renderer/visual-arrange-actions'

const PNG = 'data:image/png;base64,iVBORw0KGgo='

type UndoHandler = (
  accessor: unknown,
  params?: { token: number; direction: 'undo' | 'redo' },
) => boolean

/// The facade surface the embed path touches: the host cell's clearContent
/// and the undo stack the journal step lands on.
function makeHarness() {
  const cleared: [number, number][] = []
  const pending: number[] = []
  const refreshes: string[] = []
  const undoStack: { undoMutations: { params: unknown }[] }[] = []
  let undoHandler: UndoHandler | null = null
  const worksheet = {
    getRange: (row: number, column: number) => ({
      clearContent: () => cleared.push([row, column]),
    }),
  }
  const univerAPI = {
    getActiveWorkbook: () => ({ getId: () => 'wb', getSheetBySheetId: () => worksheet }),
  }
  const services = new Map<unknown, unknown>([
    [
      ICommandService,
      {
        registerCommand: (command: { handler: UndoHandler }) => {
          undoHandler = command.handler
          return { dispose() {} }
        },
      },
    ],
    [
      IUndoRedoService,
      { pushUndoRedo: (item: (typeof undoStack)[number]) => undoStack.push(item) },
    ],
  ])
  const editJournal = createEditJournal()
  const state = {
    editJournal,
    file: { sessionId: 's', visuals: [] as WorkbookVisualObject[] },
  }
  const ctx = {
    univerRef: {
      current: {
        univerAPI,
        univer: { __getInjector: () => ({ get: (token: unknown) => services.get(token) }) },
      },
    },
    lazyWorkbookRef: { current: state },
    shapeEditRef: { current: () => {} },
    setMessage: () => {},
    refreshLazyVisuals: () => refreshes.push('visuals'),
    setPendingEdits: (count: number) => pending.push(count),
    openDialog: () => {},
    resyncChart: () => {},
    refreshCellImages: () => refreshes.push('cells'),
  } as unknown as VisualArrangeContext
  const runUndo = (direction: 'undo' | 'redo') => {
    const step = undoStack[undoStack.length - 1]?.undoMutations[0]?.params as { token: number }
    return undoHandler?.(null, { token: step.token, direction })
  }
  return { ctx, editJournal, cleared, pending, refreshes, undoStack, runUndo }
}

const picture: WorkbookVisualObject = {
  id: 'pic-1',
  sheetId: 'sheet-1',
  kind: 'image',
  anchor: { fromRow: 4, fromColumn: 2, toRow: 9, toColumn: 5 },
  mediaDataUrl: PNG,
} as unknown as WorkbookVisualObject

describe('embed picture in cell', () => {
  it('moves a session picture into the journal and counts it as a pending edit', async () => {
    const { ctx, editJournal, cleared, pending, undoStack } = makeHarness()
    recordVisualAdd(editJournal, picture)
    handleVisualCommand(ctx, 'visual:embed-in-cell', 'pic-1')
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(cleared).toEqual([[4, 2]])
    expect(editJournal.visualAdds).toHaveLength(0)
    expect(editJournal.cellImageAdds).toEqual([
      { sheetId: 'sheet-1', row: 4, column: 2, mediaType: 'image/png', dataUrl: PNG },
    ])
    expect(journalSize(editJournal)).toBe(1)
    expect(pending).toEqual([1])
    expect(undoStack).toHaveLength(1)
  })

  it('undo restores the floating picture and drops the cell copy; redo embeds again', async () => {
    const { ctx, editJournal, pending, runUndo } = makeHarness()
    recordVisualAdd(editJournal, picture)
    handleVisualCommand(ctx, 'visual:embed-in-cell', 'pic-1')
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(runUndo('undo')).toBe(true)
    expect(editJournal.cellImageAdds).toHaveLength(0)
    expect(editJournal.visualAdds.map((visual) => visual.id)).toEqual(['pic-1'])
    expect(pending.at(-1)).toBe(1)
    expect(runUndo('redo')).toBe(true)
    expect(editJournal.cellImageAdds).toHaveLength(1)
    expect(editJournal.visualAdds).toHaveLength(0)
  })

  it('a session embed replaces the file picture in the same cell', () => {
    const fileImages = [
      { id: 'file-a', row: 4, column: 2 },
      { id: 'file-b', row: 7, column: 7 },
    ]
    const adds = [
      { sheetId: 'sheet-1', row: 4, column: 2, mediaType: 'image/png' as const, dataUrl: PNG },
      { sheetId: 'sheet-2', row: 7, column: 7, mediaType: 'image/png' as const, dataUrl: PNG },
    ]
    expect(mergeCellImages(fileImages, adds, 'sheet-1')).toEqual([
      { id: 'file-b', row: 7, column: 7 },
      { id: 'session-4-2', row: 4, column: 2, dataUrl: PNG },
    ])
  })
})
