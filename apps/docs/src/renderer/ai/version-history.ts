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
 */
export interface DocVersion {
  id: number
  /** the turn's instruction, trimmed — the "commit message" of this version */
  label: string
  /** wall-clock time of the turn, `HH:MM` */
  time: string
  /** document state before this turn's edits */
  doc: PmNode
  /** this version is currently rolled back to */
  rolledBack?: boolean
  /** the document as it was when the roll back happened; set only while `rolledBack` */
  rolledBackFrom?: PmNode
  /** a later version was rolled back, so this one describes a future that is no longer on the table */
  discarded?: boolean
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
