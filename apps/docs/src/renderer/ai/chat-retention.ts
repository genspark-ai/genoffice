/**
 * Retention policy for the AI panel's live chat transcript.
 *
 * Every completed run pins a full-document snapshot (the whole ProseMirror
 * document as JSON) onto its chat entry, and the inline roll-back action
 * renders from that field. Nothing ever released them, so the transcript grew
 * without limit: a 122-page document over 20 turns of conversation held
 * hundreds of megabytes of document copies in React state for the whole
 * session.
 *
 * Policy: the transcript keeps its text for every retained turn, but only the
 * newest {@link MAX_ROLLBACK_SNAPSHOTS} runs keep their snapshot, so one
 * roll-back still undoes any recent run. Older snapshots are set to
 * `undefined` rather than removed from the entry — the panel renders its
 * roll-back button under `entry.snapshot &&`, so a released snapshot also
 * removes the action and no button can ever point at a snapshot that is gone.
 */

/** Runs whose roll-back point is kept. The sibling html/markdown panels cap at 20 string snapshots; a document JSON tree is far heavier, so this is tighter. */
export const MAX_ROLLBACK_SNAPSHOTS = 5

/** Entries retained in memory. Each turn adds a user entry plus one or more assistant entries. */
export const MAX_CHAT_ENTRIES = 60

/**
 * Bound a chat array: drop the oldest entries past {@link MAX_CHAT_ENTRIES},
 * then release every snapshot beyond the newest {@link MAX_ROLLBACK_SNAPSHOTS}.
 *
 * Returns the input array itself when nothing needed trimming, so the common
 * case (an append or a delta on a small transcript) does not re-render the log.
 */
export function boundChatHistory<T extends { snapshot?: unknown }>(entries: readonly T[]): T[] {
  const kept =
    entries.length > MAX_CHAT_ENTRIES ? entries.slice(entries.length - MAX_CHAT_ENTRIES) : entries
  // walk from the newest end: the first MAX_ROLLBACK_SNAPSHOTS snapshots survive
  let keptSnapshots = 0
  let released = false
  const out: T[] = new Array(kept.length)
  for (let i = kept.length - 1; i >= 0; i--) {
    const entry = kept[i]!
    if (entry.snapshot !== undefined) {
      keptSnapshots++
      if (keptSnapshots > MAX_ROLLBACK_SNAPSHOTS) {
        // clearing rather than deleting: the entry keeps its text and tools,
        // and the render gate that offers roll-back also hides the button
        out[i] = { ...entry, snapshot: undefined } as T
        released = true
        continue
      }
    }
    out[i] = entry
  }
  if (released || kept !== entries) return out
  return entries as T[]
}
