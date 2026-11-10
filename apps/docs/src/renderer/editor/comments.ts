/**
 * Comment mark operations for Review → New/Delete Comment.
 *
 * The comment mark stores space-separated ids so overlapping comments share
 * one mark instance; adding/removing an id therefore rewrites the ids attr
 * per text node instead of blindly stacking marks.
 */
import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { EditorState } from '@tiptap/pm/state'
import type { CommentInfo } from '@genoffice/docx-engine'
import { TRACK_IGNORE } from './revisions'
import { wordSegmentAt } from './word-range'

/** smallest unused numeric comment id */
export function nextCommentId(comments: CommentInfo[]): string {
  const max = comments.reduce((acc, c) => Math.max(acc, parseInt(c.id, 10) || 0), 0)
  return String(max + 1)
}

/**
 * The word under a collapsed caret, as Word anchors a new comment when
 * nothing is selected. Null for a non-empty selection, a caret outside any
 * word, or a caret in a non-text block.
 */
export function wordRangeAtCaret(editor: Editor): { from: number; to: number } | null {
  const { $from, empty } = editor.state.selection
  if (!empty || !$from.parent.isTextblock) return null
  // NUL stands in for leaf nodes: never word-like, so offsets stay aligned
  const text = $from.parent.textBetween(0, $from.parent.content.size, '\0', '\0')
  const seg = wordSegmentAt(text, $from.parentOffset)
  if (!seg) return null
  const start = $from.start()
  return { from: start + seg.start, to: start + seg.end }
}

/** attach `id` to every text node in the current selection; false when selection is empty */
export function addCommentToSelection(editor: Editor, id: string): boolean {
  const { from, to } = editor.state.selection
  return addCommentToRange(editor, from, to, id)
}

/** attach `id` to every text node between the positions; false for an empty range */
export function addCommentToRange(editor: Editor, from: number, to: number, id: string): boolean {
  const { state } = editor
  if (from >= to) return false
  const markType = state.schema.marks.comment
  const tr = state.tr
  tr.setMeta(TRACK_IGNORE, true)
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText) return
    const start = Math.max(pos, from)
    const end = Math.min(pos + node.nodeSize, to)
    if (start >= end) return
    const existing = node.marks.find((m) => m.type === markType)
    const ids = new Set(
      String(existing?.attrs.ids ?? '')
        .split(' ')
        .filter(Boolean),
    )
    ids.add(id)
    tr.addMark(start, end, markType.create({ ids: [...ids].sort().join(' ') }))
  })
  editor.view.dispatch(tr)
  return true
}

/** strip `id` from every comment mark in the document (mark removed when it was the last id) */
export function removeCommentFromDoc(editor: Editor, id: string): void {
  removeCommentsFromDoc(editor, [id])
}

/** one pass for a whole batch (Delete All Comments): per-id passes would rescan the document each time */
export function removeCommentsFromDoc(editor: Editor, victims: Iterable<string>): void {
  const gone = new Set(victims)
  if (gone.size === 0) return
  const { state } = editor
  const markType = state.schema.marks.comment
  const tr = state.tr
  tr.setMeta(TRACK_IGNORE, true)
  state.doc.descendants((node, pos) => {
    if (node.isText) {
      const existing = node.marks.find((m) => m.type === markType)
      if (!existing) return
      const ids = String(existing.attrs.ids ?? '')
        .split(' ')
        .filter(Boolean)
      if (!ids.some((x) => gone.has(x))) return
      const remaining = ids.filter((x) => !gone.has(x))
      const from = pos
      const to = pos + node.nodeSize
      if (remaining.length === 0) tr.removeMark(from, to, markType)
      else tr.addMark(from, to, markType.create({ ids: remaining.join(' ') }))
      return
    }

    // Cross-paragraph ranges live on block attrs rather than text marks.
    // Clear both endpoints in the same transaction so a deleted comment can
    // never be serialized back as orphaned commentRangeStart/End markers.
    const starts = Array.isArray(node.attrs?.commentStarts)
      ? (node.attrs.commentStarts as string[])
      : []
    const ends = Array.isArray(node.attrs?.commentEnds) ? (node.attrs.commentEnds as string[]) : []
    if (!starts.some((x) => gone.has(x)) && !ends.some((x) => gone.has(x))) return
    const remainingStarts = starts.filter((value) => !gone.has(value))
    const remainingEnds = ends.filter((value) => !gone.has(value))
    tr.setNodeMarkup(pos, undefined, {
      ...node.attrs,
      commentStarts: remainingStarts.length > 0 ? remainingStarts : null,
      commentEnds: remainingEnds.length > 0 ? remainingEnds : null,
    })
  })
  if (tr.steps.length > 0) editor.view.dispatch(tr)
}

/** Append `newId` to every text range anchored by the `parentId` comment (Word: replies share the parent anchor) */
export function addReplyToCommentRange(editor: Editor, parentId: string, newId: string): boolean {
  const { state } = editor
  const markType = state.schema.marks.comment
  const tr = state.tr
  tr.setMeta(TRACK_IGNORE, true)
  let found = false
  state.doc.descendants((node, pos) => {
    if (!node.isText) return
    const existing = node.marks.find((m) => m.type === markType)
    if (!existing) return
    const ids = String(existing.attrs.ids ?? '')
      .split(' ')
      .filter(Boolean)
    if (!ids.includes(parentId)) return
    found = true
    const merged = [...new Set([...ids, newId])].sort()
    tr.addMark(pos, pos + node.nodeSize, markType.create({ ids: merged.join(' ') }))
  })
  if (found && tr.steps.length > 0) editor.view.dispatch(tr)
  return found
}

export interface CommentAnchor {
  id: string
  from: number
  to: number
}

/** first anchored range of every comment id, in document order (block-attr ranges count from the block start) */
export function commentAnchors(doc: PmNode): CommentAnchor[] {
  const markType = doc.type.schema.marks.comment
  const seen = new Map<string, CommentAnchor>()
  doc.descendants((node, pos) => {
    if (node.isText) {
      const mark = markType && node.marks.find((m) => m.type === markType)
      if (!mark) return
      for (const id of String(mark.attrs.ids ?? '').split(' ')) {
        if (!id) continue
        const hit = seen.get(id)
        if (!hit) seen.set(id, { id, from: pos, to: pos + node.nodeSize })
        else if (hit.to === pos) hit.to = pos + node.nodeSize
      }
      return
    }
    const starts = Array.isArray(node.attrs?.commentStarts)
      ? (node.attrs.commentStarts as string[])
      : []
    for (const id of starts) {
      if (!seen.has(id)) seen.set(id, { id, from: pos + 1, to: pos + 1 })
    }
  })
  return [...seen.values()].sort((a, b) => a.from - b.from)
}

/** comment ids whose mark covers `pos` (either side of a boundary) */
export function commentIdsAt(state: EditorState, pos: number): string[] {
  const markType = state.schema.marks.comment
  if (!markType || pos < 0 || pos > state.doc.content.size) return []
  const $pos = state.doc.resolve(pos)
  const ids = new Set<string>()
  for (const node of [$pos.nodeAfter, $pos.nodeBefore]) {
    const mark = node?.isText ? node.marks.find((m) => m.type === markType) : undefined
    if (!mark) continue
    for (const id of String(mark.attrs.ids ?? '').split(' ')) if (id) ids.add(id)
  }
  return [...ids]
}
