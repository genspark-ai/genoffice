/**
 * Draft splicing for voice input. An interim transcript owns a span of the draft and
 * is replaced in place as the transcriber revises it, so the user sees the sentence
 * settle while they speak. Pure string functions: the session hook is then only a
 * state machine over these edits.
 */

/** A half-open span `[start, end)` of the draft that the current session owns. */
export interface VoiceAnchor {
  readonly start: number
  readonly end: number
}

/** Where a session may write. `null` means the session is detached from its span. */
export type VoiceInsertion = VoiceAnchor | null

function clampIndex(value: string, index: number): number {
  if (!Number.isFinite(index)) return value.length
  return Math.min(Math.max(Math.trunc(index), 0), value.length)
}

/** Replace `[start, end)` with `text`, tolerating a stale or out-of-range span. */
export function replaceRange(value: string, start: number, end: number, text: string): string {
  const from = clampIndex(value, start)
  const to = Math.max(from, clampIndex(value, end))
  return value.slice(0, from) + text + value.slice(to)
}

/** Open a session over the current selection: an empty span inserts, a range replaces. */
export function anchorAt(value: string, start: number, end = start): VoiceAnchor {
  const from = clampIndex(value, start)
  return { start: from, end: Math.max(from, clampIndex(value, end)) }
}

/**
 * Write a chunk into an anchored span and return the span the next chunk replaces.
 * The caller decides whether the chunk is interim or final; the edit is the same.
 */
export function writeAnchored(
  value: string,
  anchor: VoiceAnchor,
  text: string,
): { readonly value: string; readonly anchor: VoiceAnchor } {
  const start = clampIndex(value, anchor.start)
  return {
    value: replaceRange(value, start, anchor.end, text),
    anchor: { start, end: start + text.length },
  }
}

/** Take an interim transcript back out, leaving the draft as it was before the session. */
export function discardAnchored(value: string, anchor: VoiceAnchor): string {
  return replaceRange(value, anchor.start, anchor.end, '')
}
