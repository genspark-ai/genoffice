import type { ShapeRenderNode } from '@genoffice/pptx-render'
import type { ActionCtx } from './action-context'
import { deleteVertex, shapeToEditablePath, vertices, type PathCmd } from './edit-points'

export function deleteSelectedVertex(ctx: ActionCtx): void {
  const target = ctx.editPointsTarget
  if (!target || target.vertex == null || target.dragging) return
  const node = ctx.slide?.nodes.find((n) => n.sourceId === target.sourceId)
  if (!node || (node.type !== 'shape' && node.type !== 'text')) return
  const shape = node as ShapeRenderNode
  const cmds = shapeToEditablePath(shape)
  const v = cmds && vertices(cmds)[target.vertex]
  const next = cmds && v ? deleteVertex(cmds, v) : null
  if (!next) return
  ctx.commitEditPoints(target.sourceId, { w: shape.box.w, h: shape.box.h, cmds: next }, false)
  ctx.setEditPointsTarget({ sourceId: target.sourceId, vertex: null })
}

export interface EditPointsCommit {
  /** Slide the drag started on: a page switch mid-drag must not redirect the flush */
  slideIndex: number
  sourceId: string
  path: { w: number; h: number; cmds: PathCmd[] }
}

/**
 * Tracks a drag whose previews have started but whose final commit has not
 * arrived. The main process merges every op into the open gesture until it sees
 * preview=false, so an exit that unmounts the handles mid-drag must flush one.
 */
export function createPreviewTracker() {
  let open: EditPointsCommit | null = null
  return {
    note(commit: EditPointsCommit, preview: boolean): void {
      open = preview ? commit : null
    },
    /** The last previewed path still awaiting its final commit, cleared on return */
    flush(): EditPointsCommit | null {
      const out = open
      open = null
      return out
    },
  }
}
