import type { PmNode } from './editor/convert'

/**
 * Word-style progressive open for large documents.
 *
 * Mounting a 60-page document in one setContent keeps the "Opening…" screen up
 * until the whole ProseMirror DOM is built and laid out — the dominant share of
 * cold-open time. Instead, mount enough blocks to overfill the first screens,
 * let the browser paint them, then append the tail in frame-sized chunks while
 * the user already reads page 1 (the page count in the status bar grows as the
 * tail streams in, like Word's background pagination).
 *
 * While the tail streams the document is incomplete, so:
 * - the editor is read-only (App gates editable on the loading flag),
 * - saves must wait on waitForFullContent() — a mid-stream save would
 *   serialize, and write to disk, a truncated document,
 * - appends bypass undo history and restore the dirty flag, so streaming is
 *   not an edit; history resets once the tail lands.
 */
export interface PhasedContentHost {
  /** full replace (Tiptap setContent) */
  setContent(doc: PmNode): void
  /** append top-level nodes at the document end, outside undo history */
  appendNodes(nodes: PmNode[]): void
  isDestroyed(): boolean
  resetHistory(): void
  /** read-only + save-gate flag while the tail streams */
  setLoading(loading: boolean): void
  getDirty(): boolean
  setDirty(dirty: boolean): void
}

/** blocks in the first synchronous mount: overfills the first screens at any zoom */
export const PHASE1_BLOCKS = 64
/** blocks appended per scheduled chunk while the tail streams in */
export const PHASE_CHUNK_BLOCKS = 128
/** documents at or below this many top-level blocks mount in one pass as before */
export const PHASED_MIN_BLOCKS = 192

/** bumping the token cancels the pending tail (its chunks become no-ops) */
let token = 0
let settled: Promise<void> = Promise.resolve()
/** clears the loading flag / resolves settled when a load is cancelled mid-stream */
let cancelPending: (() => void) | null = null

/** Resolves when the streaming tail has fully landed (immediately when none is pending). */
export function waitForFullContent(): Promise<void> {
  return settled
}

/** true while a phased open still has tail chunks to append */
export function isPhasedContentPending(): boolean {
  return cancelPending !== null
}

/** Any full content replacement outside the phased path must drop a pending tail. */
export function cancelPhasedContent(): void {
  token++
  cancelPending?.()
  cancelPending = null
}

/** double-rAF: the browser paints the previous mount between the two callbacks */
const nextPaintedFrame = (cb: () => void): void => {
  requestAnimationFrame(() => requestAnimationFrame(cb))
}

export function setContentPhased(
  host: PhasedContentHost,
  pmDoc: PmNode,
  schedule: (cb: () => void) => void = nextPaintedFrame,
): void {
  cancelPhasedContent()
  const content = pmDoc.content ?? []
  const my = ++token
  if (content.length <= PHASED_MIN_BLOCKS) {
    host.setContent(pmDoc)
    return
  }
  host.setContent({ ...pmDoc, content: content.slice(0, PHASE1_BLOCKS) })
  let settle!: () => void
  settled = new Promise<void>((resolve) => (settle = resolve))
  host.setLoading(true)
  cancelPending = () => {
    host.setLoading(false)
    settle()
  }
  let index = PHASE1_BLOCKS
  const finish = () => {
    if (my !== token) return
    cancelPending = null
    if (!host.isDestroyed()) host.resetHistory()
    host.setLoading(false)
    settle()
  }
  const appendChunk = () => {
    if (my !== token) return
    if (host.isDestroyed()) return finish()
    const chunk = content.slice(index, index + PHASE_CHUNK_BLOCKS)
    index += chunk.length
    const wasDirty = host.getDirty()
    try {
      host.appendNodes(chunk)
    } catch {
      // a refused chunk falls back to the one-pass mount instead of leaving the file truncated
      try {
        host.setContent(pmDoc)
      } catch {
        /* same failure as the one-pass path */
      }
      host.setDirty(wasDirty)
      return finish()
    }
    host.setDirty(wasDirty)
    if (index < content.length) schedule(appendChunk)
    else finish()
  }
  schedule(appendChunk)
}
