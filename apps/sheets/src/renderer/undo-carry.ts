/**
 * Cross-save undo: the ⌘S pipeline swaps the whole workbook session (new
 * sessionId, new `file-<sha>` Univer unit), which used to wipe the undo
 * history on every save (and every 30 s with AutoSave on). This module
 * carries Univer's own undo stack across that reopen.
 *
 * Why this works: the reopened unit renders exactly the pre-save screen
 * content, so the recorded undo mutations stay coordinate- and value-valid.
 * Only identifiers go stale, and both are fixed at capture time:
 *  - unitId: every `file-<oldsha>` string is rewritten to `file-<newsha>`,
 *    so replayed mutations resolve the new unit AND pass the journal
 *    listener's unitId filter (the replayed undo must re-enter the fresh
 *    edit journal to be included in the next save);
 *  - style ids: undo payloads reference the old unit's style pool (`s` as
 *    id string); the rebuilt unit re-pools styles from streaming with new
 *    ids, so referenced styles are inlined as objects while the old pool
 *    is still alive.
 *
 * Install happens via a stash/consume pair rather than right after the swap:
 * the reopen's load-time decoration (file-table registration, notes, defined
 * names) pushes its own undo entries, which are load artifacts, not user
 * edits. The consumer runs after that decoration, clears those artifacts,
 * and installs the carried items. Any failure anywhere degrades to the
 * pre-existing behavior: a cleared undo stack.
 *
 * Exclusions (the capture caller guards saves with sheet add/remove ops
 * and two-phase split saves): entries at or below an interactive visual
 * edit step are dropped — those undo via closures over the pre-save
 * LazyWorkbookState (see pushVisualUndo) and cannot survive the swap.
 *
 * The redo stack is carried the same way (undoing everything and then
 * letting AutoSave tick must not eat the redos). Install order matters:
 * pushUndoRedo empties the unit's redo stack on every call, so the redo
 * items land directly in the service's stack map after the undo pushes.
 */
import { IUndoRedoService } from '@univerjs/core'

import type { UniverRuntime } from './univer-state'

/// Custom mutation id used by pushVisualUndo (univer-sync.ts) for chart /
/// visual edit steps; its params are registry tokens resolving to closures.
export const VISUAL_UNDO_COMMAND_ID = 'sheets.mutation.visual-edit-step'

interface CarriedMutation {
  id: string
  params?: unknown
}

export interface CarriedUndoItem {
  unitID: string
  undoMutations: readonly CarriedMutation[]
  redoMutations: readonly CarriedMutation[]
}

export interface UndoCarry {
  unitId: string
  items: readonly unknown[]
  redoItems: readonly unknown[]
}

/// Entries at or below the newest visual-edit step are not replayable after
/// the session swap (closure-based); keep only the newer, pure suffix.
/// Applies to both stacks: they pop from the end, so entries below a visual
/// step are unreachable without replaying it first.
export function carriableSuffix<T extends CarriedUndoItem>(items: readonly T[]): T[] {
  let cut = 0
  for (let index = 0; index < items.length; index += 1) {
    const item = items[index]!
    const visual = [...item.undoMutations, ...item.redoMutations].some(
      (mutation) => mutation.id === VISUAL_UNDO_COMMAND_ID,
    )
    if (visual) cut = index + 1
  }
  return items.slice(cut)
}

/// Deep plain-data clone that (a) rewrites every string equal to the old
/// unitId, and (b) inlines cell style references: a property named `s`
/// holding an id found in the old style pool becomes the style object
/// itself. Unknown ids and already-inline objects pass through untouched.
export function rewriteCarriedValue(
  value: unknown,
  oldUnitId: string,
  newUnitId: string,
  styles: Readonly<Record<string, unknown>>,
): unknown {
  if (typeof value === 'string') return value === oldUnitId ? newUnitId : value
  if (Array.isArray(value)) {
    return value.map((entry) => rewriteCarriedValue(entry, oldUnitId, newUnitId, styles))
  }
  if (typeof value === 'object' && value !== null) {
    const out: Record<string, unknown> = {}
    for (const [key, entry] of Object.entries(value)) {
      if (key === 's' && typeof entry === 'string' && styles[entry] != null) {
        out[key] = rewriteCarriedValue(styles[entry], oldUnitId, newUnitId, styles)
      } else {
        out[key] = rewriteCarriedValue(entry, oldUnitId, newUnitId, styles)
      }
    }
    return out
  }
  return value
}

/// Mutations pushed by openLazyWorkbook's load-time decoration (file-table
/// registration, note installs, defined-name installs). An undo entry made of
/// nothing else is a load artifact, not a user edit. The distinction matters
/// because the decoration is partly async (addTable settles after the consume
/// is scheduled): the consume must drop artifacts while keeping any real edit
/// the user managed to make in that window.
const DECORATION_MUTATIONS = new Set([
  'sheet.mutation.add-table',
  'sheet.mutation.delete-table',
  'sheet.mutation.set-sheet-table',
  'sheet.mutation.set-table-filter',
  'sheet.mutation.update-note',
  'sheet.mutation.update-note-position',
  'sheet.mutation.remove-note',
  'sheet.mutation.toggle-note-popup',
  'formula.mutation.set-defined-name',
  'formula.mutation.remove-defined-name',
])

export function isLoadArtifact(item: CarriedUndoItem): boolean {
  const mutations = [...item.undoMutations, ...item.redoMutations]
  return (
    mutations.length > 0 && mutations.every((mutation) => DECORATION_MUTATIONS.has(mutation.id))
  )
}

interface UndoRedoServiceInternals {
  _undoStacks?: Map<string, CarriedUndoItem[]>
  _redoStacks?: Map<string, CarriedUndoItem[]>
  _updateStatus?: () => void
  clearUndoRedo(unitId: string): void
  pushUndoRedo(item: unknown): void
}

function undoRedoService(runtime: UniverRuntime): UndoRedoServiceInternals {
  return runtime.univer.__getInjector().get<UndoRedoServiceInternals>(IUndoRedoService)
}

export function undoStackDepth(runtime: UniverRuntime | null): number {
  if (!runtime) return 0
  try {
    const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
    if (!unitId) return 0
    return undoRedoService(runtime)._undoStacks?.get(unitId)?.length ?? 0
  } catch {
    return 0
  }
}

/// Snapshot the active unit's undo stack, rewritten against the saved file's
/// unitId. Must run while the pre-save unit is still alive (style pool).
/// Returns null when there is nothing carriable; never throws.
export function captureUndoCarry(runtime: UniverRuntime | null, newSha: string): UndoCarry | null {
  try {
    const workbook = runtime?.univerAPI.getActiveWorkbook()
    if (!runtime || !workbook) return null
    const oldUnitId = workbook.getId()
    const newUnitId = `file-${newSha}`
    const service = undoRedoService(runtime)
    const undoable = carriableSuffix(service._undoStacks?.get(oldUnitId) ?? [])
    const redoable = carriableSuffix(service._redoStacks?.get(oldUnitId) ?? [])
    if (undoable.length === 0 && redoable.length === 0) return null
    const styles = (workbook.getSnapshot().styles ?? {}) as Record<string, unknown>
    const rewrite = (item: CarriedUndoItem): unknown =>
      rewriteCarriedValue(item, oldUnitId, newUnitId, styles)
    return {
      unitId: newUnitId,
      items: undoable.map(rewrite),
      redoItems: redoable.map(rewrite),
    }
  } catch {
    // Fall back to today's behavior: the reopen clears the history.
    return null
  }
}

let pending: UndoCarry | null = null

/// Called by the save flow right before openLazyWorkbook; null clears any
/// stale stash (split saves and sheet-op saves opt out of carrying).
export function stashUndoCarry(carry: UndoCarry | null): void {
  pending = carry
}

/// True while a stashed carry has not been consumed yet — i.e. the previous
/// save's reopen has not finished its undo bookkeeping. A capture taken in
/// that state would snapshot a half-settled stack, so the save flow skips
/// carrying (stashing null also unblocks the next save if a broken reopen
/// never consumed).
export function hasPendingUndoCarry(): boolean {
  return pending !== null
}

/// Called by openLazyWorkbook once the rebuilt unit finished its load-time
/// decoration. The decoration's own undo entries (file-table, note and
/// defined-name installs) are load artifacts, not edits: they get dropped, so
/// every open starts clean — undoing a table registration would silently
/// strip its filter dropdowns. Anything else on the stack is an edit the
/// user made while the async part of the decoration settled; those are
/// re-pushed on top of the carried history (they are chronologically newest).
/// A stale stash for another unit is dropped either way.
export function consumePendingUndoCarry(runtime: UniverRuntime, unitId: string): void {
  const carry = pending !== null && pending.unitId === unitId ? pending : null
  pending = null
  let service: UndoRedoServiceInternals
  try {
    service = undoRedoService(runtime)
  } catch {
    return
  }
  try {
    const notArtifact = (item: CarriedUndoItem): boolean => !isLoadArtifact(item)
    const racedUndo = (service._undoStacks?.get(unitId) ?? []).filter(notArtifact)
    // A raced edit the user already undid sits on the redo stack; it must
    // survive the artifact cleanup just like its still-undoable siblings.
    const racedRedo = (service._redoStacks?.get(unitId) ?? []).filter(notArtifact)
    service.clearUndoRedo(unitId)
    for (const item of carry?.items ?? []) service.pushUndoRedo(item)
    for (const item of racedUndo) service.pushUndoRedo(item)
    // Any edit made in the window invalidates the carried redos, exactly as
    // a normal edit clears the redo stack — even one that was undone again
    // (its own push cleared them at the time). Its undone entries become the
    // redo tail instead.
    const redoTail =
      racedUndo.length > 0 || racedRedo.length > 0
        ? racedRedo
        : ((carry?.redoItems ?? []) as CarriedUndoItem[])
    if (redoTail.length > 0) {
      // pushUndoRedo empties the redo stack on every call, so the redo items
      // go in last, straight into the service's stack map.
      const stacks = service._redoStacks
      if (stacks) {
        const stack = stacks.get(unitId) ?? []
        stack.length = 0
        stack.push(...redoTail)
        stacks.set(unitId, stack)
        // Private, but the only way to light up the QAT redo button; the
        // stack itself is already consistent if this is ever removed.
        service._updateStatus?.()
      }
    }
  } catch {
    // A partially installed stack could replay out of order; reset to the
    // safe pre-feature state (empty history).
    try {
      service.clearUndoRedo(unitId)
    } catch {
      // The service itself is broken; leave whatever state it holds.
    }
  }
}
