/**
 * Excel collapse-then-move for arrow navigation on a multi-cell selection.
 *
 * Univer's MoveSelectionCommand computes the arrow step from the selection
 * RANGE's edge in the pressed direction, so with K97:L97 selected (active
 * cell K97) ArrowRight lands on M97 — one past the range — instead of
 * Excel's L97, one step from the ACTIVE cell (alpha feedback, right after
 * cutting the two cells). Excel's rule: a plain or Ctrl arrow first
 * collapses the selection to the active cell, then moves from there.
 *
 * Cancel the move, collapse the selection to the primary's block, and
 * replay the same command: upstream's handler then computes the step from
 * the collapsed cell, keeping its hidden-row/column skipping and merge
 * handling (and the edge-wrap guard re-evaluates against the collapsed
 * range, so it no longer blocks moves that stay inside the old selection).
 * Enter/Tab (move-selection-enter-tab) keep upstream's move-within-selection
 * behavior and are not intercepted; Shift+arrows extend and never collapse.
 */
import { ICommandService, IUniverInstanceService, RANGE_TYPE } from '@univerjs/core'
import type { IRange } from '@univerjs/core'
import {
  SelectionMoveType,
  SetSelectionsOperation,
  SheetsSelectionsService,
  getSheetCommandTarget,
} from '@univerjs/sheets'

import { MOVE_SELECTION_COMMAND } from './selection-wrap-fix'
import type { UniverRuntime } from './univer-state'

/** genoffice's Ctrl+Arrow data-edge jump (excel-jump-nav.ts) — same rule */
const JUMP_MOVE_COMMAND = 'genoffice.command.move-selection-excel-jump'

export interface PrimaryBlock {
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}

/**
 * The single block an arrow should move from: the active cell (or its whole
 * merge). Null when the selection already is that block — nothing to do.
 */
export function collapseRange(
  range: IRange,
  primary: PrimaryBlock | null | undefined,
): IRange | null {
  if (!primary) return null
  if (
    range.startRow === primary.startRow &&
    range.endRow === primary.endRow &&
    range.startColumn === primary.startColumn &&
    range.endColumn === primary.endColumn
  ) {
    return null
  }
  return {
    startRow: primary.startRow,
    startColumn: primary.startColumn,
    endRow: primary.endRow,
    endColumn: primary.endColumn,
    rangeType: RANGE_TYPE.NORMAL,
  }
}

export function installArrowCollapse(runtime: UniverRuntime): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  let replaying = false
  return runtime.univerAPI.addEvent(runtime.univerAPI.Event.BeforeCommandExecute, (event) => {
    if (replaying) return
    if (event.id !== MOVE_SELECTION_COMMAND && event.id !== JUMP_MOVE_COMMAND) return
    const params = (event.params ?? {}) as {
      direction?: unknown
      extra?: string
      fromCurrentSelection?: boolean
    }
    if (params.direction === undefined) return
    // formula-editor ref moves and ref-range selections keep their semantics
    if (params.extra === 'formula-editor' || params.fromCurrentSelection) return
    const target = getSheetCommandTarget(injector.get(IUniverInstanceService))
    if (!target) return
    const selection = injector.get(SheetsSelectionsService).getCurrentLastSelection()
    if (!selection) return
    const collapsed = collapseRange(selection.range, selection.primary)
    if (!collapsed) return
    event.cancel = true
    // replay outside the before-execute hook so the collapse and the move
    // run as ordinary top-level commands
    const commandId = event.id
    queueMicrotask(() => {
      const commandService = injector.get(ICommandService)
      commandService.syncExecuteCommand(SetSelectionsOperation.id, {
        unitId: target.unitId,
        subUnitId: target.subUnitId,
        type: SelectionMoveType.MOVE_END,
        selections: [{ range: collapsed, primary: selection.primary }],
      })
      // both move handlers return a Promise, which syncExecuteCommand rejects;
      // before-execute listeners still run synchronously, so the guard only
      // needs to span the call
      replaying = true
      try {
        void commandService.executeCommand(commandId, params)
      } finally {
        replaying = false
      }
    })
  })
}
