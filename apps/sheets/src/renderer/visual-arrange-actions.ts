/**
 * Right-click menu and Arrange-tab commands for floating objects: z-order,
 * align/distribute, rotate/flip, cut/copy/paste, and the dialogs' applies.
 * Every change journals (file visuals as drawing edits, session visuals in
 * place) and lands on the Univer undo stack as one step.
 */
import { ADDABLE_SHAPE_TYPES } from '@genoffice/xlsx-gateway/shared/shape-types'
import { BehaviorSubject } from 'rxjs'
import type { WorkbookVisualObject } from '../shared/desktop-api'
import {
  isSheetRemoved,
  journalSize,
  recordVisualAdd,
  recordVisualEdit,
  removeVisualAdd,
  reorderVisualAdd,
  updateVisualAdd,
} from './edit-journal'
import { t } from './i18n/locale'
import {
  captureVisualJournal,
  projectVisualEdits,
  pushVisualUndo,
  restoreVisualJournal,
} from './univer-sync'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'
import {
  alignBoxes,
  changedPositions,
  distributeBoxes,
  normalizeDegrees,
  reorderZ,
  rotatedAabb,
  type AlignMode,
  type DistributeAxis,
  type ZOrderOp,
} from './visual-arrange'
import {
  anchorForBox,
  cellOrigin,
  frameBox,
  sheetMetrics,
  type SheetMetrics,
} from './visual-geometry'
import {
  clearVisualSelection,
  EMU_PER_PIXEL,
  installedVisualFrames,
  selectedVisualIds,
  type ShapeEditChanges,
} from './WorkbookVisuals'

export type VisualDialogKind = 'alt-text' | 'size' | 'hyperlink'

export interface VisualArrangeContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  shapeEditRef: { readonly current: (visualId: string, changes: ShapeEditChanges) => void }
  setMessage: (message: string) => void
  refreshLazyVisuals: (state: LazyWorkbookState) => void
  setPendingEdits: (count: number) => void
  openDialog: (kind: VisualDialogKind, visualId: string) => void
  /// Re-converges a restored chart onto its live cell refs (undo/redo).
  resyncChart: (state: LazyWorkbookState, editKey: string) => void
}

/// Pixel size of the frame the dialog edits (unzoomed sheet pixels).
export interface VisualFrameInfo {
  readonly width: number
  readonly height: number
}

let clipboard: WorkbookVisualObject | null = null
/// True while nothing is on the visual clipboard; the cell context menu's
/// Paste Picture/Shape entry disables on it.
export const visualPasteDisabled$ = new BehaviorSubject(true)

/// The visual as it currently renders: file object with pending edits, or
/// the session add.
export function liveVisual(
  state: LazyWorkbookState,
  visualId: string,
): WorkbookVisualObject | null {
  const added = state.editJournal.visualAdds.find((visual) => visual.id === visualId)
  if (added) return added
  const file = state.file.visuals.find((visual) => visual.id === visualId)
  if (!file) return null
  return projectVisualEdits([file], state.editJournal.visualEdits)[0] ?? null
}

/// Installed frame size of a visual (what the user sees), when known.
export function visualFrameInfo(
  state: LazyWorkbookState,
  visualId: string,
): VisualFrameInfo | null {
  const visual = liveVisual(state, visualId)
  if (!visual) return null
  const frame = installedVisualFrames(visual.sheetId).find((entry) => entry.visual.id === visualId)
  if (!frame) return null
  if (visual.rotation && visual.frameWidth && visual.frameHeight) {
    return { width: visual.frameWidth / EMU_PER_PIXEL, height: visual.frameHeight / EMU_PER_PIXEL }
  }
  return { width: frame.width, height: frame.height }
}

/// Dispatches a `visual:` command; false when the command is not one.
export function handleVisualCommand(
  ctx: VisualArrangeContext,
  command: string,
  targetId?: string,
): boolean {
  if (!command.startsWith('visual:')) return false
  const [, group = '', argument = ''] = command.split(':')
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  const visualId = targetId ?? selectedVisualIds().at(-1)
  if (group === 'paste') {
    pasteVisual(ctx)
    return true
  }
  if (!runtime || !visualId) return true
  if (!state) {
    // Demo workbooks have no journal: deletion still works object by
    // object through the shape edit path; arrange commands do not.
    if (group === 'delete') {
      const ids = selectedVisualIds().includes(visualId) ? selectedVisualIds() : [visualId]
      for (const id of ids) ctx.shapeEditRef.current(id, { remove: true })
      clearVisualSelection()
    } else {
      ctx.setMessage(t('appVisualNotEditable'))
    }
    return true
  }
  const visual = liveVisual(state, visualId)
  if (!visual) return true
  switch (group) {
    case 'cut':
    case 'copy':
      void copyVisual(ctx, visual).then((copied) => {
        if (!copied) return
        if (group === 'cut') ctx.shapeEditRef.current(visualId, { remove: true })
        else ctx.setMessage(t('appVisualCopied'))
      })
      return true
    case 'alt-text':
    case 'size':
    case 'hyperlink':
      ctx.openDialog(group, visualId)
      return true
    case 'delete': {
      const ids = selectedVisualIds()
      if (ids.length > 1 && ids.includes(visualId)) deleteSelection(ctx, runtime, state, ids)
      else ctx.shapeEditRef.current(visualId, { remove: true })
      return true
    }
    case 'z':
      reorder(ctx, runtime, state, visual, argument as ZOrderOp)
      return true
    case 'align':
      arrangeSelection(ctx, runtime, state, (boxes) => alignBoxes(boxes, argument as AlignMode), 2)
      return true
    case 'distribute':
      arrangeSelection(
        ctx,
        runtime,
        state,
        (boxes) => distributeBoxes(boxes, argument as DistributeAxis),
        3,
      )
      return true
    case 'rotate':
      rotateVisual(ctx, state, visual, argument === 'ccw' ? -90 : 90)
      return true
    case 'flip':
      ctx.shapeEditRef.current(
        visualId,
        argument === 'v' ? { flipV: !visual.flipV } : { flipH: !visual.flipH },
      )
      return true
    default:
      return true
  }
}

function activeMetrics(runtime: UniverRuntime, sheetId: string): SheetMetrics | null {
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(sheetId)
  return worksheet ? sheetMetrics(worksheet) : null
}

function reorder(
  ctx: VisualArrangeContext,
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  visual: WorkbookVisualObject,
  op: ZOrderOp,
): void {
  const journal = state.editJournal
  const isAdded = journal.visualAdds.some((candidate) => candidate.id === visual.id)
  if (isAdded) {
    // Session visuals sit above every file anchor; they reorder among
    // themselves on their sheet.
    const before = [...journal.visualAdds]
    const sheetIds = before.filter((v) => v.sheetId === visual.sheetId).map((v) => v.id)
    const after = reorderZ(sheetIds, visual.id, op)
    if (after.join() === sheetIds.join()) return
    // Position among the sheet's adds maps onto the global list through the
    // add now preceding it.
    const previousId = after[after.indexOf(visual.id) - 1]
    const others = before.filter((v) => v.id !== visual.id)
    const globalTo =
      previousId === undefined
        ? Math.max(
            0,
            others.findIndex((v) => v.sheetId === visual.sheetId),
          )
        : others.findIndex((v) => v.id === previousId) + 1
    reorderVisualAdd(journal, visual.id, globalTo)
    const afterList = [...journal.visualAdds]
    pushVisualUndo(runtime, {
      undo: () => {
        journal.visualAdds.splice(0, journal.visualAdds.length, ...before)
        ctx.refreshLazyVisuals(state)
      },
      redo: () => {
        journal.visualAdds.splice(0, journal.visualAdds.length, ...afterList)
        ctx.refreshLazyVisuals(state)
      },
    })
    ctx.setMessage(t('appVisualArranged'))
    ctx.refreshLazyVisuals(state)
    return
  }
  if (visual.drawingPath === undefined) {
    ctx.setMessage(t('appVisualNotEditable'))
    return
  }
  // Group children share their group's anchor and cannot move on their own.
  if (visual.drawingIndex === undefined) {
    ctx.setMessage(t('appVisualNotEditable'))
    return
  }
  const part = projectVisualEdits(
    state.file.visuals.filter(
      (candidate) =>
        candidate.drawingPath === visual.drawingPath && candidate.drawingIndex !== undefined,
    ),
    journal.visualEdits,
  )
  const order = part.map((candidate) => candidate.id)
  const after = reorderZ(order, visual.id, op)
  const changes = changedPositions(order, after)
  if (changes.size === 0) return
  const snapshots = new Map(
    [...changes.keys()].map((id) => [id, captureVisualJournal(state, id, undefined)] as const),
  )
  for (const [id, zIndex] of changes) {
    const target = state.file.visuals.find((candidate) => candidate.id === id)
    if (!target || !recordVisualEdit(journal, target, { zIndex })) {
      for (const [snapshotId, snapshot] of snapshots) {
        restoreVisualJournal(state, snapshotId, undefined, snapshot)
      }
      ctx.setMessage(t('appVisualNotEditable'))
      return
    }
  }
  const afterSnapshots = new Map(
    [...changes.keys()].map((id) => [id, captureVisualJournal(state, id, undefined)] as const),
  )
  pushVisualUndo(runtime, {
    undo: () => {
      for (const [id, snapshot] of snapshots) restoreVisualJournal(state, id, undefined, snapshot)
      ctx.refreshLazyVisuals(state)
    },
    redo: () => {
      for (const [id, snapshot] of afterSnapshots) {
        restoreVisualJournal(state, id, undefined, snapshot)
      }
      ctx.refreshLazyVisuals(state)
    },
  })
  ctx.setPendingEdits(journalSize(journal))
  ctx.setMessage(t('appVisualArranged'))
  ctx.refreshLazyVisuals(state)
}

/// Removes every selected object as one undo step (applyShapeEdit records
/// one step per call).
function deleteSelection(
  ctx: VisualArrangeContext,
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  ids: readonly string[],
): void {
  const journal = state.editJournal
  const chartPathOf = (id: string): string | undefined =>
    state.file.visuals.find((candidate) => candidate.id === id)?.chartPath
  const snapshots = new Map(
    ids.map((id) => [id, captureVisualJournal(state, id, chartPathOf(id))] as const),
  )
  // Session z-order is the add list order; restoreVisualJournal re-appends
  // a missing add, so the whole list is restored around it.
  const addsBefore = [...journal.visualAdds]
  const restoreAll = (
    entries: ReadonlyMap<string, ReturnType<typeof captureVisualJournal>>,
    adds: readonly WorkbookVisualObject[],
  ): void => {
    for (const [id, snapshot] of entries) restoreVisualJournal(state, id, chartPathOf(id), snapshot)
    journal.visualAdds.splice(0, journal.visualAdds.length, ...adds)
    for (const id of ids) ctx.resyncChart(state, chartPathOf(id) ?? id)
  }
  for (const id of ids) {
    if (removeVisualAdd(journal, id)) continue
    const visual = state.file.visuals.find((candidate) => candidate.id === id)
    if (!visual || !recordVisualEdit(journal, visual, { remove: true })) {
      for (const [snapshotId, snapshot] of snapshots) {
        restoreVisualJournal(state, snapshotId, chartPathOf(snapshotId), snapshot)
      }
      ctx.setMessage(t('appVisualNoDelete'))
      return
    }
    if (visual.chartPath !== undefined) journal.chartEdits.delete(visual.chartPath)
  }
  const afterSnapshots = new Map(
    ids.map((id) => [id, captureVisualJournal(state, id, chartPathOf(id))] as const),
  )
  const addsAfter = [...journal.visualAdds]
  clearVisualSelection()
  pushVisualUndo(runtime, {
    undo: () => {
      restoreAll(snapshots, addsBefore)
      ctx.refreshLazyVisuals(state)
    },
    redo: () => {
      restoreAll(afterSnapshots, addsAfter)
      clearVisualSelection()
      ctx.refreshLazyVisuals(state)
    },
  })
  ctx.setPendingEdits(journalSize(journal))
  ctx.setMessage(t('appVisualDeleted'))
  ctx.refreshLazyVisuals(state)
}

function journalAnchor(
  state: LazyWorkbookState,
  visualId: string,
  anchor: WorkbookVisualObject['anchor'],
): boolean {
  if (updateVisualAdd(state.editJournal, visualId, { anchor })) return true
  const visual = state.file.visuals.find((candidate) => candidate.id === visualId)
  return visual !== undefined && recordVisualEdit(state.editJournal, visual, { anchor })
}

function arrangeSelection(
  ctx: VisualArrangeContext,
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  plan: (boxes: ReturnType<typeof frameBox>[]) => readonly { id: string; x: number; y: number }[],
  minimum: number,
): void {
  const ids = selectedVisualIds()
  if (ids.length < minimum) {
    ctx.setMessage(t(minimum === 3 ? 'appVisualSelectThree' : 'appVisualSelectTwo'))
    return
  }
  const primary = liveVisual(state, ids[ids.length - 1]!)
  if (!primary) return
  const metrics = activeMetrics(runtime, primary.sheetId)
  if (!metrics) return
  const frames = installedVisualFrames(primary.sheetId).filter((frame) =>
    ids.includes(frame.visual.id),
  )
  const boxes = frames.map((frame) => frameBox(frame, metrics))
  const moves = plan(boxes)
  if (moves.length === 0) return
  const snapshots = new Map(
    moves.map((move) => [move.id, captureVisualJournal(state, move.id, undefined)] as const),
  )
  for (const move of moves) {
    const box = boxes.find((candidate) => candidate.id === move.id)!
    if (!journalAnchor(state, move.id, anchorForBox({ ...box, x: move.x, y: move.y }, metrics))) {
      for (const [id, snapshot] of snapshots) restoreVisualJournal(state, id, undefined, snapshot)
      ctx.setMessage(t('appVisualNoMove'))
      return
    }
  }
  const afterSnapshots = new Map(
    moves.map((move) => [move.id, captureVisualJournal(state, move.id, undefined)] as const),
  )
  pushVisualUndo(runtime, {
    undo: () => {
      for (const [id, snapshot] of snapshots) restoreVisualJournal(state, id, undefined, snapshot)
      ctx.refreshLazyVisuals(state)
    },
    redo: () => {
      for (const [id, snapshot] of afterSnapshots) {
        restoreVisualJournal(state, id, undefined, snapshot)
      }
      ctx.refreshLazyVisuals(state)
    },
  })
  ctx.setPendingEdits(journalSize(state.editJournal))
  ctx.setMessage(t('appVisualArranged'))
  ctx.refreshLazyVisuals(state)
}

/// Rotation / size changes keep the frame center and re-derive the anchor
/// as the rotated frame's AABB (the install's own invariant).
export function frameChangeEdit(
  state: LazyWorkbookState,
  runtime: UniverRuntime,
  visual: WorkbookVisualObject,
  next: { width: number; height: number; rotation: number },
): ShapeEditChanges | null {
  const frame = installedVisualFrames(visual.sheetId).find((entry) => entry.visual.id === visual.id)
  const metrics = activeMetrics(runtime, visual.sheetId)
  if (!frame || !metrics) return null
  const box = frameBox(frame, metrics)
  const rotation = normalizeDegrees(next.rotation)
  const aabb = rotatedAabb(next.width, next.height, rotation)
  const anchor = anchorForBox(
    {
      x: box.x + box.width / 2 - aabb.width / 2,
      y: box.y + box.height / 2 - aabb.height / 2,
      width: aabb.width,
      height: aabb.height,
    },
    metrics,
  )
  return {
    anchor,
    rotation,
    frameSize: {
      width: Math.max(1, Math.round(next.width * EMU_PER_PIXEL)),
      height: Math.max(1, Math.round(next.height * EMU_PER_PIXEL)),
    },
  }
}

function rotateVisual(
  ctx: VisualArrangeContext,
  state: LazyWorkbookState,
  visual: WorkbookVisualObject,
  delta: number,
): void {
  const runtime = ctx.univerRef.current
  const size = visualFrameInfo(state, visual.id)
  if (!runtime || !size) return
  if (visual.kind === 'chart') {
    ctx.setMessage(t('appVisualNotEditable'))
    return
  }
  const edit = frameChangeEdit(state, runtime, visual, {
    width: size.width,
    height: size.height,
    rotation: (visual.rotation ?? 0) + delta,
  })
  if (edit) ctx.shapeEditRef.current(visual.id, edit)
}

async function copyVisual(
  ctx: VisualArrangeContext,
  visual: WorkbookVisualObject,
): Promise<boolean> {
  const state = ctx.lazyWorkbookRef.current
  const copyable =
    (visual.kind === 'shape' && ADDABLE_SHAPE_TYPES.some((type) => type === visual.shapeType)) ||
    (visual.kind === 'image' &&
      (visual.mediaDataUrl !== undefined ||
        ['image/png', 'image/jpeg', 'image/gif'].includes(visual.mediaType ?? '')))
  if (!copyable || !state) {
    ctx.setMessage(t('appVisualNoCopy'))
    return false
  }
  const size = visualFrameInfo(state, visual.id)
  const frame = size
    ? { frameWidth: size.width * EMU_PER_PIXEL, frameHeight: size.height * EMU_PER_PIXEL }
    : {}
  let copy: WorkbookVisualObject = { ...visual, ...frame }
  if (visual.kind === 'image' && visual.mediaDataUrl === undefined) {
    // File pictures carry no bytes in the model; a cut must hold them
    // before the original goes, so the fetch completes first.
    try {
      const media = await window.desktopApi.readWorkbookMedia({
        sessionId: state.file.sessionId,
        visualId: visual.id,
      })
      copy = {
        ...copy,
        mediaType: media.mediaType,
        mediaDataUrl: `data:${media.mediaType};base64,${media.base64}`,
      }
    } catch {
      ctx.setMessage(t('appVisualNoCopy'))
      return false
    }
  }
  clipboard = copy
  visualPasteDisabled$.next(false)
  return true
}

function pasteVisual(ctx: VisualArrangeContext): void {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  const source = clipboard
  if (!runtime || !state || !source) return
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!worksheet || !range) return
  const sheetId = worksheet.getSheetId()
  if (isSheetRemoved(state.editJournal, sheetId)) return
  if (source.kind === 'image' && source.mediaDataUrl === undefined) {
    ctx.setMessage(t('appVisualNoCopy'))
    return
  }
  const metrics = sheetMetrics(worksheet)
  const origin = cellOrigin(range.getRow(), range.getColumn(), metrics)
  const width = (source.frameWidth ?? 96 * EMU_PER_PIXEL) / EMU_PER_PIXEL
  const height = (source.frameHeight ?? 96 * EMU_PER_PIXEL) / EMU_PER_PIXEL
  const rotation = normalizeDegrees(source.rotation ?? 0)
  const aabb = rotatedAabb(width, height, rotation)
  const stamp = `${Date.now().toString(36)}-${state.editJournal.visualAdds.length + 1}`
  const {
    id: _id,
    drawingPath: _path,
    drawingIndex: _index,
    mediaPath: _mediaPath,
    chart: _chart,
    chartPath: _chartPath,
    ...rest
  } = source
  const visual: WorkbookVisualObject = {
    ...rest,
    id: `added-${source.kind}-${stamp}`,
    sheetId,
    anchor: anchorForBox({ ...origin, width: aabb.width, height: aabb.height }, metrics),
    ...(rotation
      ? {
          rotation,
          frameWidth: Math.round(width * EMU_PER_PIXEL),
          frameHeight: Math.round(height * EMU_PER_PIXEL),
        }
      : {}),
  }
  recordVisualAdd(state.editJournal, visual)
  pushVisualUndo(runtime, {
    undo: () => {
      removeVisualAdd(state.editJournal, visual.id)
      clearVisualSelection(visual.id)
      ctx.refreshLazyVisuals(state)
    },
    redo: () => {
      recordVisualAdd(state.editJournal, visual)
      ctx.refreshLazyVisuals(state)
    },
  })
  ctx.setPendingEdits(journalSize(state.editJournal))
  ctx.setMessage(t('appVisualPasted'))
  ctx.refreshLazyVisuals(state)
}
