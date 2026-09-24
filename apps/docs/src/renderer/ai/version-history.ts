import type { ChatMessage } from '@genoffice/project-store'
import type { PmNode } from '../editor/convert'

/** Ring cap, matching the Markdown / HTML / Slides snapshot rings. */
export const MAX_VERSIONS = 20

/**
 * One AI turn that edited the document: the document as it was before the turn's
 * first edit, plus what is needed to take the turn back and to take it again.
 *
 * Versions are never destroyed by a roll back. A rolled-back version keeps the
 * state it replaced in `rolledBackFrom`, and the versions after it are only
 * *marked* as discarded futures, so undoing the roll back brings back both the
 * document and every rollback point it had voided.
 *
 * The document itself lives on disk between sessions (the project store keeps it
 * gzipped, addressed by `snapshotId`), so `doc` may be absent: it comes back when
 * the version is restored, or is found to be gone and the point reads as expired.
 */
export interface DocVersion {
  id: number
  /** the turn's instruction, trimmed — the "commit message" of this version */
  label: string
  /** wall-clock time of the turn, `HH:MM` */
  time: string
  /** document state before this turn's edits; absent until a stored one is read */
  doc?: PmNode
  /** key of the stored snapshot, when it was small enough to keep */
  snapshotId?: string
  /** the stored snapshot is gone: the point stays listed, but cannot be restored */
  expired?: boolean
  /** this version is currently rolled back to */
  rolledBack?: boolean
  /** the document as it was when the roll back happened; set only while `rolledBack` */
  rolledBackFrom?: PmNode
  /** a later version was rolled back, so this one describes a future that is no longer on the table */
  discarded?: boolean
}

/** Whether a roll back to this version can still be attempted. */
export function canRestore(version: DocVersion): boolean {
  if (version.discarded) return false
  // A point loaded from a reopened file has no document until its snapshot is
  // read; one whose snapshot is gone can never be restored.
  return version.doc !== undefined || !version.expired
}

/**
 * Append a version from a new AI run.
 *
 * A new run starts from whatever the document is now, so the past a roll back
 * moved away from can no longer be reached by undo — those entries are dropped
 * here, which is also the lazy prune that keeps the ring bounded (#543: the old
 * code cleared snapshots eagerly on every roll back, which was a memory hack
 * dressed up as a UI state change).
 */
export function pushVersion(list: DocVersion[], version: DocVersion): DocVersion[] {
  const live = list
    .filter((v) => !v.discarded)
    .map((v) => (v.rolledBack ? { ...v, rolledBack: undefined, rolledBackFrom: undefined } : v))
  return [...live, version].slice(-MAX_VERSIONS)
}

/**
 * Mark `id` as the version the document was rolled back to, keeping `before` —
 * the document as it is right now — so the roll back can be undone.
 *
 * Every later version is marked discarded rather than removed: they still hold
 * the snapshots undoing this roll back needs, and they stay listed so the panel
 * keeps showing the whole evolution of the document.
 */
export function markRolledBack(list: DocVersion[], id: number, before: PmNode): DocVersion[] {
  const idx = list.findIndex((v) => v.id === id)
  if (idx < 0) return list
  return list.map((v, i) => {
    if (i === idx) return { ...v, rolledBack: true, rolledBackFrom: before }
    if (i > idx) return { ...v, discarded: true, rolledBack: undefined, rolledBackFrom: undefined }
    return v
  })
}

/** Undo a roll back: the version is live again, and so is everything after it. */
export function clearRollback(list: DocVersion[], id: number): DocVersion[] {
  const idx = list.findIndex((v) => v.id === id)
  if (idx < 0 || !list[idx].rolledBack) return list
  return list.map((v, i) => {
    if (i === idx) return { ...v, rolledBack: undefined, rolledBackFrom: undefined }
    if (i > idx) return { ...v, discarded: undefined }
    return v
  })
}

/** The version a user may still act on: neither discarded nor already undone. */
export function rollbackableVersions(list: DocVersion[]): DocVersion[] {
  return list.filter((v) => !v.discarded)
}

/** Record where a version's document was stored, once the main process answered. */
export function attachSnapshot(list: DocVersion[], id: number, snapshotId: string): DocVersion[] {
  return list.map((v) => (v.id === id ? { ...v, snapshotId, expired: undefined } : v))
}

/** Cache a document read back from disk, so a second roll back needs no round trip. */
export function attachDoc(list: DocVersion[], id: number, doc: PmNode): DocVersion[] {
  return list.map((v) => (v.id === id ? { ...v, doc, expired: undefined } : v))
}

/** The snapshot is gone (pruned, never stored, or unreadable): say so, don't vanish. */
export function markExpired(list: DocVersion[], id: number): DocVersion[] {
  return list.map((v) => (v.id === id && !v.doc ? { ...v, expired: true } : v))
}

/**
 * Rebuild the version list from a reopened transcript.
 *
 * Only each message's `version` is read, so the parameter is the message shape
 * rather than the whole record: a caller may hand over anything that carries one
 * (the panel passes what `loadChat` returned).
 *
 * The transcript carries each turn's ref; the documents are separate files, so
 * anything the store no longer has is marked expired here rather than
 * discovered later as a dead button (#543).
 */
export function versionsFromChat(
  messages: ReadonlyArray<Partial<ChatMessage>>,
  availableSnapshotIds: string[],
): { versions: DocVersion[]; nextId: number } {
  const available = new Set(availableSnapshotIds)
  const refs = messages.flatMap((m) => (m.version ? [m.version] : []))
  const versions: DocVersion[] = refs.map((ref) => ({
    id: ref.id,
    label: ref.label,
    time: ref.time,
    ...(ref.snapshotId ? { snapshotId: ref.snapshotId } : {}),
    // no snapshot at all (too large when it was taken), or the store dropped it
    ...(ref.snapshotId && available.has(ref.snapshotId) ? {} : { expired: true }),
  }))
  // Ids are per chat and monotonic across sessions, so a new run must continue
  // past every ref the transcript holds, not just the ones still in the ring.
  const nextId = refs.reduce((max, ref) => Math.max(max, ref.id), 0) + 1
  return { versions: versions.slice(-MAX_VERSIONS), nextId }
}
