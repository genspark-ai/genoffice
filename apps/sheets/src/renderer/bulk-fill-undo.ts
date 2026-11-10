import { CommandType, ICommandService, IUndoRedoService } from '@univerjs/core'

import type { WorkbookBulkConstantFill } from '../shared/desktop-api'
import type { JournalEntry } from './edit-journal'
import type { UniverRuntime } from './univer-state'

export const BULK_FILL_UNDO_COMMAND_ID = 'sheets.mutation.bulk-constant-fill'

interface BulkFillUndoParams {
  readonly fill: WorkbookBulkConstantFill
  /// Value entries the fill purged at record time; undo reinstates them.
  readonly purgedCells: readonly JournalEntry[]
  readonly direction: 'undo' | 'redo'
}

type BulkFillUndoHandler = (params: BulkFillUndoParams) => boolean

const handlers = new WeakMap<object, BulkFillUndoHandler>()
const registeredRuntimes = new WeakSet<object>()

/**
 * Adds a range-level constant fill to Univer's undo stack without storing a
 * closure in the mutation params. The params remain plain data, so undo-carry
 * can move the whole AI batch to the workbook session created after Save.
 */
export function pushBulkFillUndo(
  runtime: UniverRuntime,
  fill: WorkbookBulkConstantFill,
  purgedCells: readonly JournalEntry[],
  handler: BulkFillUndoHandler,
): void {
  const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
  if (!unitId) return
  const injector = (
    runtime.univer as unknown as {
      __getInjector(): { get<T>(token: unknown): T }
    }
  ).__getInjector()
  handlers.set(runtime, handler)
  if (!registeredRuntimes.has(runtime)) {
    registeredRuntimes.add(runtime)
    injector
      .get<{
        registerCommand(command: {
          id: string
          type: unknown
          handler: (accessor: unknown, params?: BulkFillUndoParams) => boolean
        }): unknown
      }>(ICommandService)
      .registerCommand({
        id: BULK_FILL_UNDO_COMMAND_ID,
        type: CommandType.MUTATION,
        handler: (_accessor, params) => {
          const current = handlers.get(runtime)
          return params !== undefined && current !== undefined ? current(params) : false
        },
      })
  }
  injector
    .get<{
      pushUndoRedo(item: {
        unitID: string
        undoMutations: { id: string; params: BulkFillUndoParams }[]
        redoMutations: { id: string; params: BulkFillUndoParams }[]
      }): void
    }>(IUndoRedoService)
    .pushUndoRedo({
      unitID: unitId,
      undoMutations: [
        { id: BULK_FILL_UNDO_COMMAND_ID, params: { fill, purgedCells, direction: 'undo' } },
      ],
      redoMutations: [
        { id: BULK_FILL_UNDO_COMMAND_ID, params: { fill, purgedCells, direction: 'redo' } },
      ],
    })
}
