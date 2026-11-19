import { ICommandService, IUndoRedoService } from '@univerjs/core'
import { describe, expect, it, vi } from 'vitest'

import { BULK_FILL_UNDO_COMMAND_ID, pushBulkFillUndo } from '../src/renderer/bulk-fill-undo'
import { carriableSuffix, type CarriedUndoItem } from '../src/renderer/undo-carry'
import type { UniverRuntime } from '../src/renderer/univer-state'

describe('bulk constant-fill undo', () => {
  it('uses serializable mutations that survive cross-save undo carrying', () => {
    let registered:
      | {
          handler: (
            accessor: unknown,
            params?: {
              fill: {
                sheetId: string
                startRow: number
                endRow: number
                startColumn: number
                endColumn: number
                value: string
              }
              direction: 'undo' | 'redo'
            },
          ) => boolean
        }
      | undefined
    let pushed: CarriedUndoItem | undefined
    const commandService = {
      registerCommand: vi.fn((command) => {
        registered = command
      }),
    }
    const undoService = {
      pushUndoRedo: vi.fn((item) => {
        pushed = item
      }),
    }
    const runtime = {
      univerAPI: { getActiveWorkbook: () => ({ getId: () => 'file-before-save' }) },
      univer: {
        __getInjector: () => ({
          get: (token: unknown) =>
            token === ICommandService
              ? commandService
              : token === IUndoRedoService
                ? undoService
                : undefined,
        }),
      },
    } as unknown as UniverRuntime
    const directions: string[] = []
    const fill = {
      sheetId: 'sheet-1',
      startRow: 1,
      endRow: 88_587,
      startColumn: 97,
      endColumn: 97,
      value: 'merrick',
    }

    const purgedCells = [{ row: 5, column: 97, hasValue: true, value: null }]
    pushBulkFillUndo(runtime, fill, purgedCells, ({ direction }) => {
      directions.push(direction)
      return true
    })

    expect(undoService.pushUndoRedo).toHaveBeenCalledOnce()
    expect(pushed?.undoMutations).toEqual([
      { id: BULK_FILL_UNDO_COMMAND_ID, params: { fill, purgedCells, direction: 'undo' } },
    ])
    expect(pushed?.redoMutations).toEqual([
      { id: BULK_FILL_UNDO_COMMAND_ID, params: { fill, purgedCells, direction: 'redo' } },
    ])
    expect(carriableSuffix([pushed!])).toEqual([pushed])
    expect(registered?.handler(undefined, pushed?.undoMutations[0]?.params as never)).toBe(true)
    expect(registered?.handler(undefined, pushed?.redoMutations[0]?.params as never)).toBe(true)
    expect(directions).toEqual(['undo', 'redo'])
  })
})
