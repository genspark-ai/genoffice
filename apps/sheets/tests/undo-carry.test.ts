import { describe, expect, it } from 'vitest'

import {
  VISUAL_UNDO_COMMAND_ID,
  captureUndoCarry,
  carriableSuffix,
  consumePendingUndoCarry,
  hasPendingUndoCarry,
  isLoadArtifact,
  rewriteCarriedValue,
  stashUndoCarry,
  type CarriedUndoItem,
} from '../src/renderer/undo-carry'
import type { UniverRuntime } from '../src/renderer/univer-state'

const OLD_UNIT = 'file-oldsha'
const NEW_UNIT = 'file-newsha'

function cellItem(unitID: string, value: unknown): CarriedUndoItem {
  return {
    unitID,
    undoMutations: [
      { id: 'sheet.mutation.set-range-values', params: { unitId: unitID, cellValue: value } },
    ],
    redoMutations: [
      { id: 'sheet.mutation.set-range-values', params: { unitId: unitID, cellValue: value } },
    ],
  }
}

function visualItem(): CarriedUndoItem {
  return {
    unitID: OLD_UNIT,
    undoMutations: [{ id: VISUAL_UNDO_COMMAND_ID }],
    redoMutations: [{ id: VISUAL_UNDO_COMMAND_ID }],
  }
}

function tableArtifact(unitID: string): CarriedUndoItem {
  return {
    unitID,
    undoMutations: [{ id: 'sheet.mutation.delete-table', params: { unitId: unitID } }],
    redoMutations: [{ id: 'sheet.mutation.add-table', params: { unitId: unitID } }],
  }
}

describe('isLoadArtifact', () => {
  it('classifies decoration-only entries as artifacts', () => {
    expect(isLoadArtifact(tableArtifact(NEW_UNIT))).toBe(true)
  })

  it('keeps cell edits and mixed entries as user history', () => {
    expect(isLoadArtifact(cellItem(NEW_UNIT, 1))).toBe(false)
    const mixed: CarriedUndoItem = {
      unitID: NEW_UNIT,
      undoMutations: [
        { id: 'sheet.mutation.add-table' },
        { id: 'sheet.mutation.set-range-values' },
      ],
      redoMutations: [],
    }
    expect(isLoadArtifact(mixed)).toBe(false)
    expect(isLoadArtifact({ unitID: NEW_UNIT, undoMutations: [], redoMutations: [] })).toBe(false)
  })
})

describe('carriableSuffix', () => {
  it('keeps everything when no visual steps are present', () => {
    const items = [cellItem(OLD_UNIT, 1), cellItem(OLD_UNIT, 2)]
    expect(carriableSuffix(items)).toEqual(items)
  })

  it('drops the visual step and everything older', () => {
    const newer = cellItem(OLD_UNIT, 3)
    const items = [cellItem(OLD_UNIT, 1), visualItem(), newer]
    expect(carriableSuffix(items)).toEqual([newer])
  })

  it('returns empty when the newest step is visual', () => {
    expect(carriableSuffix([cellItem(OLD_UNIT, 1), visualItem()])).toEqual([])
  })

  it('truncates at the newest of several visual steps', () => {
    const newer = cellItem(OLD_UNIT, 9)
    const items = [visualItem(), cellItem(OLD_UNIT, 1), visualItem(), newer]
    expect(carriableSuffix(items)).toEqual([newer])
  })
})

describe('rewriteCarriedValue', () => {
  it('rewrites the old unitId everywhere, at any depth', () => {
    const item = {
      unitID: OLD_UNIT,
      undoMutations: [
        { id: 'm', params: { unitId: OLD_UNIT, nested: { list: [OLD_UNIT, 'other'] } } },
      ],
    }
    const out = rewriteCarriedValue(item, OLD_UNIT, NEW_UNIT, {}) as typeof item
    expect(out.unitID).toBe(NEW_UNIT)
    expect(out.undoMutations[0]!.params.unitId).toBe(NEW_UNIT)
    expect(out.undoMutations[0]!.params.nested.list).toEqual([NEW_UNIT, 'other'])
  })

  it('leaves unrelated strings and primitives untouched', () => {
    const value = { text: 'file-oldsha-suffix', v: 42, b: true, n: null }
    expect(rewriteCarriedValue(value, OLD_UNIT, NEW_UNIT, {})).toEqual(value)
  })

  it('inlines style ids from the old pool', () => {
    const styles = { st1: { bl: 1, cl: { rgb: '#ff0000' } } }
    const cell = { v: 'x', s: 'st1' }
    const out = rewriteCarriedValue(cell, OLD_UNIT, NEW_UNIT, styles) as typeof cell
    expect(out.s).toEqual(styles.st1)
    // The inlined style is a clone, not a shared reference.
    expect(out.s).not.toBe(styles.st1)
  })

  it('passes unknown style ids and inline style objects through', () => {
    const styles = { st1: { bl: 1 } }
    const unknownId = { s: 'missing' }
    expect(rewriteCarriedValue(unknownId, OLD_UNIT, NEW_UNIT, styles)).toEqual(unknownId)
    const inline = { s: { it: 1 } }
    expect(rewriteCarriedValue(inline, OLD_UNIT, NEW_UNIT, styles)).toEqual(inline)
  })

  it('only treats `s` keys as style references', () => {
    const styles = { st1: { bl: 1 } }
    const value = { note: 'st1', s: 'st1' }
    const out = rewriteCarriedValue(value, OLD_UNIT, NEW_UNIT, styles) as typeof value
    expect(out.note).toBe('st1')
    expect(out.s).toEqual(styles.st1)
  })

  it('does not mutate the input', () => {
    const item = { unitID: OLD_UNIT, params: { s: 'st1' } }
    rewriteCarriedValue(item, OLD_UNIT, NEW_UNIT, { st1: { bl: 1 } })
    expect(item).toEqual({ unitID: OLD_UNIT, params: { s: 'st1' } })
  })
})

interface FakeService {
  _undoStacks: Map<string, CarriedUndoItem[]>
  _redoStacks: Map<string, CarriedUndoItem[]>
  cleared: string[]
  pushed: unknown[]
  statusUpdates: number
  failPush?: boolean
  _updateStatus(): void
  clearUndoRedo(unitId: string): void
  pushUndoRedo(item: unknown): void
}

function fakeRuntime(service: FakeService, stylesPool: Record<string, unknown> = {}) {
  return {
    univer: { __getInjector: () => ({ get: () => service }) },
    univerAPI: {
      getActiveWorkbook: () => ({
        getId: () => OLD_UNIT,
        getSnapshot: () => ({ styles: stylesPool }),
      }),
    },
  } as unknown as UniverRuntime
}

function fakeService(stack: CarriedUndoItem[], redoStack: CarriedUndoItem[] = []): FakeService {
  const service: FakeService = {
    _undoStacks: new Map([[OLD_UNIT, stack]]),
    _redoStacks: new Map([[OLD_UNIT, redoStack]]),
    cleared: [],
    pushed: [],
    statusUpdates: 0,
    _updateStatus() {
      service.statusUpdates += 1
    },
    clearUndoRedo(unitId: string) {
      service.cleared.push(unitId)
      service._redoStacks.get(unitId)?.splice(0)
    },
    pushUndoRedo(item: unknown) {
      if (service.failPush) throw new Error('boom')
      service.pushed.push(item)
      // Mirror the real service: every push empties the unit's redo stack.
      service._redoStacks.get((item as CarriedUndoItem).unitID)?.splice(0)
    },
  }
  return service
}

describe('captureUndoCarry / consumePendingUndoCarry', () => {
  it('captures, rewrites, and installs across the swap', () => {
    const service = fakeService([cellItem(OLD_UNIT, 1), cellItem(OLD_UNIT, 2)])
    const runtime = fakeRuntime(service)
    const carry = captureUndoCarry(runtime, 'newsha')
    expect(carry).not.toBeNull()
    expect(carry!.unitId).toBe(NEW_UNIT)
    expect(carry!.items).toHaveLength(2)

    stashUndoCarry(carry)
    consumePendingUndoCarry(runtime, NEW_UNIT)
    // Load-artifact entries are cleared before the carried history lands.
    expect(service.cleared).toEqual([NEW_UNIT])
    expect(service.pushed).toHaveLength(2)
    const first = service.pushed[0] as CarriedUndoItem
    expect(first.unitID).toBe(NEW_UNIT)
    expect((first.undoMutations[0] as { params?: { unitId?: string } }).params?.unitId).toBe(
      NEW_UNIT,
    )
  })

  it('returns null when both stacks are empty or visual-only', () => {
    expect(captureUndoCarry(fakeRuntime(fakeService([])), 'newsha')).toBeNull()
    expect(
      captureUndoCarry(fakeRuntime(fakeService([visualItem()], [visualItem()])), 'newsha'),
    ).toBeNull()
  })

  it('carries a redo-only history (undo exhausted before an autosave)', () => {
    const service = fakeService([], [cellItem(OLD_UNIT, 1), cellItem(OLD_UNIT, 2)])
    const runtime = fakeRuntime(service)
    const carry = captureUndoCarry(runtime, 'newsha')
    expect(carry).not.toBeNull()
    expect(carry!.items).toHaveLength(0)
    expect(carry!.redoItems).toHaveLength(2)

    stashUndoCarry(carry)
    consumePendingUndoCarry(runtime, NEW_UNIT)
    const installed = service._redoStacks.get(NEW_UNIT)
    expect(installed).toHaveLength(2)
    expect((installed![0] as CarriedUndoItem).unitID).toBe(NEW_UNIT)
    expect(service.statusUpdates).toBe(1)
  })

  it('installs redo after the undo pushes so pushUndoRedo cannot wipe it', () => {
    const service = fakeService([cellItem(OLD_UNIT, 1)], [cellItem(OLD_UNIT, 2)])
    const runtime = fakeRuntime(service)
    stashUndoCarry(captureUndoCarry(runtime, 'newsha'))
    consumePendingUndoCarry(runtime, NEW_UNIT)
    expect(service.pushed).toHaveLength(1)
    expect(service._redoStacks.get(NEW_UNIT)).toHaveLength(1)
  })

  it('truncates redo entries at visual steps independently of undo', () => {
    const keptRedo = cellItem(OLD_UNIT, 9)
    const service = fakeService(
      [cellItem(OLD_UNIT, 1)],
      [cellItem(OLD_UNIT, 2), visualItem(), keptRedo],
    )
    const carry = captureUndoCarry(fakeRuntime(service), 'newsha')
    expect(carry!.items).toHaveLength(1)
    expect(carry!.redoItems).toHaveLength(1)
  })

  it('returns null without a runtime and never throws on broken internals', () => {
    expect(captureUndoCarry(null, 'newsha')).toBeNull()
    const broken = {
      univer: {
        __getInjector: () => {
          throw new Error('disposed')
        },
      },
      univerAPI: { getActiveWorkbook: () => ({ getId: () => OLD_UNIT }) },
    } as unknown as UniverRuntime
    expect(captureUndoCarry(broken, 'newsha')).toBeNull()
  })

  it('clears load artifacts even without a stash, installing nothing', () => {
    const service = fakeService([])
    const runtime = fakeRuntime(service)
    stashUndoCarry(null)
    consumePendingUndoCarry(runtime, NEW_UNIT)
    expect(service.cleared).toEqual([NEW_UNIT])
    expect(service.pushed).toEqual([])
  })

  it('drops a stale stash for another unit while still clearing artifacts', () => {
    const service = fakeService([cellItem(OLD_UNIT, 1)])
    const runtime = fakeRuntime(service)
    stashUndoCarry(captureUndoCarry(runtime, 'newsha'))

    consumePendingUndoCarry(runtime, 'file-unrelated')
    expect(service.cleared).toEqual(['file-unrelated'])
    expect(service.pushed).toEqual([])

    // The stale stash was dropped: a later reopen of its target unit must
    // not install it (the content baseline may have moved on).
    service.cleared = []
    consumePendingUndoCarry(runtime, NEW_UNIT)
    expect(service.cleared).toEqual([NEW_UNIT])
    expect(service.pushed).toEqual([])
  })

  it('consumes the stash once', () => {
    const service = fakeService([cellItem(OLD_UNIT, 1)])
    const runtime = fakeRuntime(service)
    stashUndoCarry(captureUndoCarry(runtime, 'newsha'))
    consumePendingUndoCarry(runtime, NEW_UNIT)
    expect(service.pushed).toHaveLength(1)

    // Already consumed: a second reopen of the same unit installs nothing
    // (but still resets the stack — a plain open of the same file).
    service.pushed = []
    service.cleared = []
    consumePendingUndoCarry(runtime, NEW_UNIT)
    expect(service.cleared).toEqual([NEW_UNIT])
    expect(service.pushed).toEqual([])
  })

  it('falls back to a cleared stack when installing fails midway', () => {
    const service = fakeService([cellItem(OLD_UNIT, 1)])
    const runtime = fakeRuntime(service)
    stashUndoCarry(captureUndoCarry(runtime, 'newsha'))
    service.failPush = true
    consumePendingUndoCarry(runtime, NEW_UNIT)
    // Cleared once before installing, once again as the failure fallback.
    expect(service.cleared).toEqual([NEW_UNIT, NEW_UNIT])
    expect(service.pushed).toEqual([])
  })

  it('a null stash overrides a previous carry: artifacts cleared, nothing installed', () => {
    const service = fakeService([cellItem(OLD_UNIT, 1)])
    const runtime = fakeRuntime(service)
    stashUndoCarry(captureUndoCarry(runtime, 'newsha'))
    stashUndoCarry(null)
    consumePendingUndoCarry(runtime, NEW_UNIT)
    expect(service.cleared).toEqual([NEW_UNIT])
    expect(service.pushed).toEqual([])
  })

  it('drops load artifacts but keeps edits raced during the reopen window', () => {
    const service = fakeService([cellItem(OLD_UNIT, 1)])
    const runtime = fakeRuntime(service)
    stashUndoCarry(captureUndoCarry(runtime, 'newsha'))
    const raced = cellItem(NEW_UNIT, 'typed-during-reopen')
    service._undoStacks.set(NEW_UNIT, [tableArtifact(NEW_UNIT), raced])

    consumePendingUndoCarry(runtime, NEW_UNIT)
    expect(service.pushed).toHaveLength(2)
    expect(service.pushed[1]).toBe(raced)
    expect(service.cleared).toEqual([NEW_UNIT])
  })

  it('keeps raced edits on a plain open too', () => {
    const service = fakeService([])
    const runtime = fakeRuntime(service)
    stashUndoCarry(null)
    const raced = cellItem(NEW_UNIT, 'typed-during-open')
    service._undoStacks.set(NEW_UNIT, [tableArtifact(NEW_UNIT), raced])

    consumePendingUndoCarry(runtime, NEW_UNIT)
    expect(service.cleared).toEqual([NEW_UNIT])
    expect(service.pushed).toEqual([raced])
  })

  it('a raced edit invalidates the carried redos', () => {
    const service = fakeService([cellItem(OLD_UNIT, 1)], [cellItem(OLD_UNIT, 2)])
    const runtime = fakeRuntime(service)
    stashUndoCarry(captureUndoCarry(runtime, 'newsha'))
    service._undoStacks.set(NEW_UNIT, [cellItem(NEW_UNIT, 'raced')])

    consumePendingUndoCarry(runtime, NEW_UNIT)
    expect(service.pushed).toHaveLength(2)
    expect(service._redoStacks.get(NEW_UNIT) ?? []).toHaveLength(0)
  })

  it('a raced edit that was undone again survives as the redo tail', () => {
    const service = fakeService([cellItem(OLD_UNIT, 1)], [cellItem(OLD_UNIT, 2)])
    const runtime = fakeRuntime(service)
    stashUndoCarry(captureUndoCarry(runtime, 'newsha'))
    const undone = cellItem(NEW_UNIT, 'typed-then-undone')
    service._undoStacks.set(NEW_UNIT, [tableArtifact(NEW_UNIT)])
    service._redoStacks.set(NEW_UNIT, [undone])

    consumePendingUndoCarry(runtime, NEW_UNIT)
    // Carried undo installed, artifacts dropped, and the undone raced edit —
    // not the carried redos it invalidated — is the redo stack.
    expect(service.pushed).toHaveLength(1)
    expect(service._redoStacks.get(NEW_UNIT)).toEqual([undone])
  })

  it('keeps both halves of a raced edit pair (one kept, one undone)', () => {
    const service = fakeService([cellItem(OLD_UNIT, 1)])
    const runtime = fakeRuntime(service)
    stashUndoCarry(captureUndoCarry(runtime, 'newsha'))
    const kept = cellItem(NEW_UNIT, 'kept')
    const undone = cellItem(NEW_UNIT, 'undone')
    service._undoStacks.set(NEW_UNIT, [tableArtifact(NEW_UNIT), kept])
    service._redoStacks.set(NEW_UNIT, [undone])

    consumePendingUndoCarry(runtime, NEW_UNIT)
    expect(service.pushed[1]).toBe(kept)
    expect(service._redoStacks.get(NEW_UNIT)).toEqual([undone])
  })

  it('restores an undone raced edit on a plain open too', () => {
    const service = fakeService([])
    const runtime = fakeRuntime(service)
    stashUndoCarry(null)
    const undone = cellItem(NEW_UNIT, 'undone-during-open')
    service._undoStacks.set(NEW_UNIT, [tableArtifact(NEW_UNIT)])
    service._redoStacks.set(NEW_UNIT, [tableArtifact(NEW_UNIT), undone])

    consumePendingUndoCarry(runtime, NEW_UNIT)
    expect(service.pushed).toEqual([])
    expect(service._redoStacks.get(NEW_UNIT)).toEqual([undone])
  })

  it('hasPendingUndoCarry tracks the stash lifecycle', () => {
    const service = fakeService([cellItem(OLD_UNIT, 1)])
    const runtime = fakeRuntime(service)
    stashUndoCarry(null)
    expect(hasPendingUndoCarry()).toBe(false)
    stashUndoCarry(captureUndoCarry(runtime, 'newsha'))
    expect(hasPendingUndoCarry()).toBe(true)
    consumePendingUndoCarry(runtime, NEW_UNIT)
    expect(hasPendingUndoCarry()).toBe(false)
  })

  it('inlines styles referenced by carried cell values', () => {
    const stylesPool = { st9: { bl: 1 } }
    const service = fakeService([cellItem(OLD_UNIT, { 0: { 0: { v: 'x', s: 'st9' } } })])
    const runtime = fakeRuntime(service, stylesPool)
    const carry = captureUndoCarry(runtime, 'newsha')
    const item = carry!.items[0] as {
      undoMutations: { params: { cellValue: { 0: { 0: { s: unknown } } } } }[]
    }
    expect(item.undoMutations[0]!.params.cellValue[0][0].s).toEqual(stylesPool.st9)
  })
})
