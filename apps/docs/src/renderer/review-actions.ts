/**
 * Review-tab actions: footnotes/endnotes, comments, revisions, ink
 * annotations and compare (document protection lives in ProtectDialog +
 * App.applyProtectDialog). Extracted from App.tsx; the App component passes a
 * ReviewContext built fresh per call so state never goes stale.
 */
import type { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { nextNoteId, parseDocx, type CommentInfo, type NoteInfo } from '@genoffice/docx-engine'
import type { Dispatch, SetStateAction } from 'react'
import { fetchDocBytes } from './doc-bytes'
import type { DocState } from './doc-state'
import {
  addCommentToRange,
  addCommentToSelection,
  addReplyToCommentRange,
  commentAnchors,
  commentIdsAt,
  nextCommentId,
  removeCommentFromDoc,
  removeCommentsFromDoc,
  wordRangeAtCaret,
} from './editor/comments'
import {
  blockTexts,
  compareParagraphs,
  editorBlockTexts,
  type CompareEntry,
} from './editor/compare'
import { pendingCommentPluginKey } from './editor/extensions'
import type { InkAnnotation } from './editor/ink'
import {
  TRACK_IGNORE,
  acceptAllRevisions,
  acceptCurrentRevision,
  rejectAllRevisions,
  rejectCurrentRevision,
} from './editor/revisions'
import { t } from './i18n/locale'

/** open the note text dialog; id set = editing an existing note */
export interface NotePrompt {
  kind: 'footnote' | 'endnote'
  id?: string
}

/** The App state the review actions need; built fresh per call. */
export interface ReviewContext {
  editor: Editor | null
  doc: DocState | null
  dirtyRef: { current: boolean }
  setStatus: (status: string) => void
  notePrompt: NotePrompt | null
  setNotePrompt: (value: NotePrompt | null) => void
  footnotes: NoteInfo[]
  endnotes: NoteInfo[]
  setFootnotes: Dispatch<SetStateAction<NoteInfo[]>>
  setEndnotes: Dispatch<SetStateAction<NoteInfo[]>>
  setNotesDirty: (dirty: boolean) => void
  comments: CommentInfo[]
  setComments: Dispatch<SetStateAction<CommentInfo[]>>
  setCommentsDirty: (dirty: boolean) => void
  setCommentComposing: (composing: boolean) => void
  setShowComments: (show: boolean) => void
  /** highlight one thread in the comments pane (Previous / Next) */
  setCommentFocus: (focus: { id: string; nonce: number } | null) => void
  setInkAnnotations: Dispatch<SetStateAction<InkAnnotation[]>>
  setInksDirty: (dirty: boolean) => void
  setCompareResult: (value: { otherName: string; entries: CompareEntry[] } | null) => void
}

// ---- References: footnotes / endnotes ----

/** dialog submit: create a new note (+ caret marker) or update an existing one */
export function submitNote(ctx: ReviewContext, text: string): void {
  if (!ctx.notePrompt || !ctx.editor) return
  const { kind, id } = ctx.notePrompt
  const list = kind === 'footnote' ? ctx.footnotes : ctx.endnotes
  const setList = kind === 'footnote' ? ctx.setFootnotes : ctx.setEndnotes
  if (id !== undefined) {
    setList(list.map((n) => (n.id === id ? { ...n, text } : n)))
  } else {
    const newId = nextNoteId(list)
    setList([...list, { id: newId, text }])
    ctx.editor
      .chain()
      .focus()
      .insertContent({
        type: 'docNoteRef',
        attrs: { kind, id: newId, num: list.length + 1 },
      } as never)
      .run()
  }
  ctx.setNotesDirty(true)
}

/** delete a note, remove its in-text marker, renumber the remaining markers */
export function deleteNote(ctx: ReviewContext, kind: 'footnote' | 'endnote', id: string): void {
  const list = kind === 'footnote' ? ctx.footnotes : ctx.endnotes
  const next = list.filter((n) => n.id !== id)
  ;(kind === 'footnote' ? ctx.setFootnotes : ctx.setEndnotes)(next)
  ctx.setNotesDirty(true)
  if (!ctx.editor) return
  const editor = ctx.editor
  const tr = editor.state.tr
  const removals: Array<{ pos: number; size: number }> = []
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name !== 'docNoteRef' || node.attrs.kind !== kind) return
    if (String(node.attrs.id) === id) {
      removals.push({ pos, size: node.nodeSize })
    } else {
      const num = next.findIndex((n) => n.id === String(node.attrs.id)) + 1
      if (num > 0 && num !== node.attrs.num) {
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, num })
      }
    }
  })
  for (const { pos, size } of removals.reverse()) tr.delete(pos, pos + size)
  if (tr.docChanged) editor.view.dispatch(tr)
}

// ---- Review: comments / revisions / compare / protection ----

/** keeps the picked range highlighted while the composer holds focus */
export function setPendingCommentRange(
  ctx: ReviewContext,
  range: { from: number; to: number } | null,
): void {
  if (!ctx.editor) return
  ctx.editor.view.dispatch(ctx.editor.state.tr.setMeta(pendingCommentPluginKey, range))
}

export function cancelNewComment(ctx: ReviewContext): void {
  ctx.setCommentComposing(false)
  setPendingCommentRange(ctx, null)
}

/** New comment: open the pane with the composer; the mark is applied on submit */
export function startNewComment(ctx: ReviewContext): void {
  const editor = ctx.editor
  if (!editor) return
  if (editor.state.selection.empty) {
    // Word anchors on the word under a collapsed caret rather than refusing
    const word = wordRangeAtCaret(editor)
    if (!word) {
      ctx.setStatus(t('appSelectTextToComment'))
      return
    }
    editor.commands.setTextSelection(word)
  }
  const { from, to } = editor.state.selection
  setPendingCommentRange(ctx, { from, to })
  ctx.setShowComments(true)
  ctx.setCommentComposing(true)
}

export function submitNewComment(ctx: ReviewContext, text: string): void {
  if (!ctx.editor) return
  setPendingCommentRange(ctx, null)
  const id = nextCommentId(ctx.comments)
  if (!addCommentToSelection(ctx.editor, id)) {
    ctx.setStatus(t('appCommentSelectionLost'))
    ctx.setCommentComposing(false)
    return
  }
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
  ctx.setComments((prev) => [...prev, { id, author: 'User', date: now, text }])
  ctx.setCommentsDirty(true)
  ctx.setCommentComposing(false)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appCommentAdded'))
}

/** New thread on an explicit range (AI add_comment); the new id, null when the range holds no text */
export function addCommentAt(
  ctx: ReviewContext,
  range: { from: number; to: number },
  text: string,
  author: string,
  initials?: string,
): string | null {
  if (!ctx.editor) return null
  const id = nextCommentId(ctx.comments)
  if (!addCommentToRange(ctx.editor, range.from, range.to, id)) return null
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
  ctx.setComments((prev) => [
    ...prev,
    { id, author, date: now, text, ...(initials ? { initials } : {}) },
  ])
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appCommentAdded'))
  return id
}

/** Reply to a comment: the new entry carries parentId; the anchor shares the parent comment's range */
export function replyToComment(
  ctx: ReviewContext,
  parentId: string,
  text: string,
  author = 'User',
): boolean {
  if (!ctx.editor) return false
  const id = nextCommentId(ctx.comments)
  if (!addReplyToCommentRange(ctx.editor, parentId, id)) {
    ctx.setStatus(t('appCommentAnchorGone'))
    return false
  }
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
  ctx.setComments((prev) => [...prev, { id, author, date: now, text, parentId }])
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appCommentReplied'))
  return true
}

/** Word: comment text edits in place; the author, date and anchor stay */
export function editComment(ctx: ReviewContext, id: string, text: string): void {
  ctx.setComments((prev) => prev.map((c) => (c.id === id ? { ...c, text } : c)))
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appCommentEdited'))
}

/** Resolve/reopen: the whole thread (parent + replies) gets done set together */
export function resolveComment(ctx: ReviewContext, id: string, done: boolean): void {
  ctx.setComments((prev) =>
    prev.map((c) => (c.id === id || c.parentId === id ? { ...c, done } : c)),
  )
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(done ? t('appCommentResolvedMsg') : t('appCommentReopenedMsg'))
}

export function deleteComment(ctx: ReviewContext, id: string): void {
  if (!ctx.editor) return
  // cascade: deleting a parent comment also deletes its replies (including their anchor marks)
  const victims = [id, ...ctx.comments.filter((c) => c.parentId === id).map((c) => c.id)]
  for (const v of victims) removeCommentFromDoc(ctx.editor, v)
  ctx.setComments((prev) => prev.filter((c) => !victims.includes(c.id)))
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
}

/** the thread under the caret (a reply resolves to its parent) */
export function commentThreadAtCaret(ctx: ReviewContext): string | null {
  if (!ctx.editor) return null
  const ids = commentIdsAt(ctx.editor.state, ctx.editor.state.selection.from)
  for (const id of ids) {
    const info = ctx.comments.find((c) => c.id === id)
    if (info) return info.parentId ?? info.id
  }
  return null
}

export function deleteCommentAtCaret(ctx: ReviewContext): void {
  const id = commentThreadAtCaret(ctx)
  if (id) deleteComment(ctx, id)
}

/** Delete All Comments in Document / Delete All Resolved Comments */
export function deleteAllComments(ctx: ReviewContext, resolvedOnly: boolean): void {
  if (!ctx.editor) return
  // replies never carry done themselves: a resolved thread goes with all of its replies
  const doneThreads = new Set(ctx.comments.filter((c) => c.done && !c.parentId).map((c) => c.id))
  const victims = ctx.comments
    .filter((c) => !resolvedOnly || c.done || (c.parentId && doneThreads.has(c.parentId)))
    .map((c) => c.id)
  if (victims.length === 0) return
  removeCommentsFromDoc(ctx.editor, victims)
  const gone = new Set(victims)
  ctx.setComments((prev) => prev.filter((c) => !gone.has(c.id)))
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
}

/** Previous / Next Comment: select the thread's anchor and light it up in the pane; open threads only */
export function gotoComment(ctx: ReviewContext, dir: 1 | -1): boolean {
  const editor = ctx.editor
  if (!editor) return false
  const open = new Set(ctx.comments.filter((c) => !c.parentId && !c.done).map((c) => c.id))
  const anchors = commentAnchors(editor.state.doc).filter((a) => open.has(a.id))
  if (anchors.length === 0) return false
  // the thread under the caret is never a stop in either direction
  const { from } = editor.state.selection
  const target =
    dir === 1
      ? (anchors.find((a) => a.from > from) ?? anchors[0])
      : ([...anchors].reverse().find((a) => a.to <= from && a.from < from) ??
        anchors[anchors.length - 1])
  const tr = editor.state.tr
  tr.setSelection(TextSelection.between(tr.doc.resolve(target.from), tr.doc.resolve(target.to)))
  tr.scrollIntoView()
  tr.setMeta(TRACK_IGNORE, true)
  editor.view.dispatch(tr)
  editor.view.focus()
  ctx.setShowComments(true)
  ctx.setCommentFocus({ id: target.id, nonce: Date.now() })
  return true
}

export function handleRevision(
  ctx: ReviewContext,
  action: 'accept' | 'reject',
  all: boolean,
): void {
  if (!ctx.editor) return
  if (all) {
    if (action === 'accept') acceptAllRevisions(ctx.editor)
    else rejectAllRevisions(ctx.editor)
    ctx.setStatus(action === 'accept' ? t('appAllRevisionsAccepted') : t('appAllRevisionsRejected'))
  } else {
    const ok =
      action === 'accept' ? acceptCurrentRevision(ctx.editor) : rejectCurrentRevision(ctx.editor)
    if (!ok) ctx.setStatus(t('appNoRevisionsToHandle'))
  }
  ctx.dirtyRef.current = true
}

// ---- Draw: overlay annotations ----

export function addInk(ctx: ReviewContext, annotation: InkAnnotation): void {
  ctx.setInkAnnotations((prev) => [...prev, annotation])
  ctx.setInksDirty(true)
  ctx.dirtyRef.current = true
}

export function removeInks(ctx: ReviewContext, ids: string[]): void {
  const gone = new Set(ids)
  ctx.setInkAnnotations((prev) => prev.filter((a) => !gone.has(a.id)))
  ctx.setInksDirty(true)
  ctx.dirtyRef.current = true
}

export function clearInks(ctx: ReviewContext): void {
  ctx.setInkAnnotations([])
  ctx.setInksDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appInksCleared'))
}

/** Compare: pick a second .docx and diff it against the open document */
export async function compareWithFile(ctx: ReviewContext): Promise<void> {
  if (!ctx.doc) return
  const other = await window.desktop.openDocx()
  if (!other) return
  // password-protected comparison target: not wired through the decrypt prompt (yet)
  if ('needsPassword' in other) {
    ctx.setStatus(t('appCompareFailed', { error: t('appDocPwdTitle') }))
    return
  }
  try {
    const otherParsed = await parseDocx(await fetchDocBytes(other.dataUrl))
    const entries = compareParagraphs(
      ctx.editor
        ? editorBlockTexts(ctx.editor.getJSON(), ctx.doc.parsed.blocks)
        : blockTexts(ctx.doc.parsed.blocks),
      blockTexts(otherParsed.blocks),
    )
    ctx.setCompareResult({ otherName: other.name, entries })
  } catch (err) {
    ctx.setStatus(t('appCompareFailed', { error: String(err) }))
  }
}
