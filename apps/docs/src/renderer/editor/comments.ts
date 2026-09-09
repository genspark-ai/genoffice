/**
 * Comment mark operations for Review → New/Delete Comment.
 *
 * The comment mark stores space-separated ids so overlapping comments share
 * one mark instance; adding/removing an id therefore rewrites the ids attr
 * per text node instead of blindly stacking marks.
 */
import type { Editor } from '@tiptap/core'
import type { CommentInfo } from '@genoffice/docx-engine'
import { TRACK_IGNORE } from './revisions'

/** smallest unused numeric comment id */
export function nextCommentId(comments: CommentInfo[]): string {
  const max = comments.reduce((acc, c) => Math.max(acc, parseInt(c.id, 10) || 0), 0)
  return String(max + 1)
}

const WORD_SEGMENTER: Intl.Segmenter | null =
  typeof Intl !== 'undefined' && 'Segmenter' in Intl
    ? new Intl.Segmenter(undefined, { granularity: 'word' })
    : null

/** the word-like segment containing `off`, or the one ending exactly at it (caret right after a word) */
function wordSegmentAt(text: string, off: number): { start: number; end: number } | null {
  if (WORD_SEGMENTER) {
    let before: { start: number; end: number } | null = null
    for (const seg of WORD_SEGMENTER.segment(text)) {
      const start = seg.index
      const end = seg.index + seg.segment.length
      if (start > off) break
      if (!seg.isWordLike) continue
      if (off < end) return { start, end }
      if (end === off) before = { start, end }
    }
    return before
  }
  const isWordChar = (ch: string) => /[\p{L}\p{N}_]/u.test(ch)
  let at = off
  if ((at >= text.length || !isWordChar(text[at]!)) && at > 0 && isWordChar(text[at - 1]!)) at--
  if (at >= text.length || !isWordChar(text[at]!)) return null
  let start = at
  while (start > 0 && isWordChar(text[start - 1]!)) start--
  let end = at + 1
  while (end < text.length && isWordChar(text[end]!)) end++
  return { start, end }
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
  const { state } = editor
  const { from, to } = state.selection
  if (from === to) return false
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
      if (!ids.includes(id)) return
      const remaining = ids.filter((x) => x !== id)
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
    if (!starts.includes(id) && !ends.includes(id)) return
    const remainingStarts = starts.filter((value) => value !== id)
    const remainingEnds = ends.filter((value) => value !== id)
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
