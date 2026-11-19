/**
 * Comment metadata undo-sync: the PM document stores only comment id anchors;
 * the metadata (author, text, date, replies) lives in App state. Undo of a
 * comment deletion therefore restored the mark but not the card — the pane
 * stayed empty and the comment looked lost (#Word-parity report: delete →
 * undo must restore the whole comment; redo must re-remove it).
 *
 * The sync follows the comment-id set across every document-changing
 * transaction, dragging metadata both ways, with SPLIT timing:
 * - removals are handled synchronously at dispatch — the live mirror still
 *   holds the victims (deleteComment writes its own setComments removal right
 *   after dispatch, which would make a deferred capture read an empty list);
 * - additions defer their revive decision to a microtask — by then a forward
 *   add's own metadata write has landed (skip: the store already holds the
 *   id), while an undo/redo write never does (revive from the graveyard).
 *
 * The graveyard clears on a document swap (resetCommentUndoStore via the
 * cross-doc reset), so ids of a previous document can never resurrect into
 * the next one.
 */
import type { Editor } from '@tiptap/core'
import type { CommentInfo } from '@genoffice/docx-engine'

const graveyard = new Map<string, CommentInfo>()

/** clear the graveyard on a document swap (cross-doc edit-state reset) */
export function resetCommentUndoStore(): void {
  graveyard.clear()
}

function commentIdsIn(doc: import('@tiptap/pm/model').Node): Set<string> {
  const ids = new Set<string>()
  doc.descendants((node) => {
    if (node.isText) {
      const mark = node.marks.find((m) => m.type.name === 'comment')
      if (!mark) return
      for (const id of String(mark.attrs.ids ?? '').split(' ')) if (id) ids.add(id)
      return
    }
    const starts = Array.isArray(node.attrs?.commentStarts)
      ? (node.attrs.commentStarts as string[])
      : []
    for (const id of starts) if (id) ids.add(id)
    const ends = Array.isArray(node.attrs?.commentEnds) ? (node.attrs.commentEnds as string[]) : []
    for (const id of ends) if (id) ids.add(id)
  })
  return ids
}

/**
 * Follow the comment-id set across transactions so undo/redo of comment
 * add/delete drags the metadata along. `getComments` reads the synchronous
 * live mirror; `applyMetadata` is the mirror+state setter the review actions
 * use.
 */
export function attachCommentUndoSync(
  editor: Editor,
  getComments: () => CommentInfo[],
  applyMetadata: (entries: CommentInfo[]) => void,
): () => void {
  let prevIds = commentIdsIn(editor.state.doc)
  let flushScheduled = false
  let pendingAdded: string[] = []

  // additions defer: the same tick may carry a forward add's own metadata
  // write, and the store must settle before deciding graveyard vs skip
  const flushAdded = (): void => {
    flushScheduled = false
    if (pendingAdded.length === 0) return
    const revived: CommentInfo[] = []
    for (const id of pendingAdded) {
      if (getComments().some((c) => c.id === id)) continue
      const entry = graveyard.get(id)
      if (!entry) continue
      graveyard.delete(id)
      revived.push(entry)
    }
    pendingAdded = []
    if (revived.length > 0) {
      applyMetadata([
        ...getComments(),
        ...revived.filter((r) => !getComments().some((c) => c.id === r.id)),
      ])
    }
  }

  const handler = ({ transaction }: { transaction?: { docChanged: boolean } } = {}): void => {
    if (!transaction?.docChanged) return
    const nowIds = commentIdsIn(editor.state.doc)
    const removed = [...prevIds].filter((id) => !nowIds.has(id))
    const added = [...nowIds].filter((id) => !prevIds.has(id))
    prevIds = nowIds
    if (removed.length === 0 && added.length === 0) return

    // removals synchronously: the mirror still holds the victims (a forward
    // delete writes its own setComments removal right after this returns —
    // and an undo-of-add's metadata must leave the store the same way)
    if (removed.length > 0) {
      const comments = getComments()
      const parked: string[] = []
      for (const id of removed) {
        const entry = comments.find((c) => c.id === id)
        if (!entry) continue
        graveyard.set(id, entry)
        parked.push(id)
      }
      if (parked.length > 0) {
        const gone = new Set(parked)
        applyMetadata(comments.filter((c) => !gone.has(c.id)))
      }
    }
    if (added.length > 0) {
      pendingAdded.push(...added)
      if (!flushScheduled) {
        flushScheduled = true
        queueMicrotask(flushAdded)
      }
    }
  }

  editor.on('transaction', handler)
  return () => editor.off('transaction', handler)
}
