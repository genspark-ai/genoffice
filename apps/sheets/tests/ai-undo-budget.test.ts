import { LocalUndoRedoService } from '@univerjs/core'
import { beforeEach, describe, expect, it } from 'vitest'

import {
  AI_UNDO_CELL_BUDGET,
  aiBulkUndoGate,
  installJournalSuppressionUndoFilter,
  undoPayloadCells,
} from '../src/renderer/univer-state'

function setRangeItem(cells: number) {
  const cellValue: Record<number, Record<number, object>> = {}
  for (let r = 0; r * 10 < cells; r++) {
    const row: Record<number, object> = {}
    for (let c = 0; c < 10 && r * 10 + c < cells; c++) row[c] = { v: 1 }
    cellValue[r] = row
  }
  return { unitID: 'wb', undoMutations: [], redoMutations: [{ params: { cellValue } }] }
}

/// The real Univer service internals the patched pushUndoRedo touches.
function fakeService(stack: unknown[], redoStack: unknown[] = []) {
  return {
    _getRedoStack: () => redoStack,
    _getUndoStack: () => stack,
    _pitchUndoElement: () => (stack.length ? stack[stack.length - 1] : null),
    _batchingStatus: new Map(),
    _updateStatus: () => {},
  }
}

describe('AI bulk-apply undo budget', () => {
  beforeEach(() => {
    installJournalSuppressionUndoFilter()
    aiBulkUndoGate.active = false
    aiBulkUndoGate.dropped = false
    aiBulkUndoGate.cells = 0
    aiBulkUndoGate.pushed = 0
  })

  it('counts payload cells with a bounded walk', () => {
    expect(undoPayloadCells(setRangeItem(250), 1000)).toBe(250)
    expect(undoPayloadCells({ unitID: 'wb' }, 1000)).toBe(0)
    expect(
      undoPayloadCells({ unitID: 'wb', redoMutations: [{ params: { style: {} } }] }, 1000),
    ).toBe(0)
  })

  it('keeps a batch under the budget on the stack', () => {
    const stack: unknown[] = []
    const push = (item: unknown) =>
      (LocalUndoRedoService.prototype.pushUndoRedo as (i: unknown) => void).call(
        fakeService(stack),
        item,
      )
    aiBulkUndoGate.active = true
    push(setRangeItem(500))
    push(setRangeItem(500))
    expect(stack.length).toBe(2)
    expect(aiBulkUndoGate.dropped).toBe(false)
  })

  it('discards the whole batch once the cumulative payload crosses the budget', () => {
    const stack: unknown[] = []
    const push = (item: unknown) =>
      (LocalUndoRedoService.prototype.pushUndoRedo as (i: unknown) => void).call(
        fakeService(stack),
        item,
      )
    aiBulkUndoGate.active = true
    push(setRangeItem(AI_UNDO_CELL_BUDGET))
    expect(stack.length).toBe(1)
    // The next chunk crosses the budget: the already-pushed part of the batch
    // must go too, or undo would revert only some chunks of the operation.
    push(setRangeItem(10))
    expect(stack.length).toBe(0)
    expect(aiBulkUndoGate.dropped).toBe(true)
    // Later chunks of the same batch stay dropped.
    push(setRangeItem(10))
    expect(stack.length).toBe(0)
  })

  it('a batch dropped on its first item still clears stale redo', () => {
    const stack: unknown[] = []
    const redoStack: unknown[] = [setRangeItem(5)]
    aiBulkUndoGate.active = true
    ;(LocalUndoRedoService.prototype.pushUndoRedo as (i: unknown) => void).call(
      fakeService(stack, redoStack),
      setRangeItem(AI_UNDO_CELL_BUDGET + 10),
    )
    expect(stack.length).toBe(0)
    expect(redoStack.length).toBe(0)
    expect(aiBulkUndoGate.dropped).toBe(true)
  })

  it('leaves non-AI pushes alone', () => {
    const stack: unknown[] = []
    ;(LocalUndoRedoService.prototype.pushUndoRedo as (i: unknown) => void).call(
      fakeService(stack),
      setRangeItem(AI_UNDO_CELL_BUDGET + 10),
    )
    expect(stack.length).toBe(1)
  })
})
