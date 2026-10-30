/**
 * The ask-intent handoff shared by the command palette's Ask AI fallback and
 * the `home:deep-link` bridge: both route a natural-language query into Home's
 * ask view, where the knowledge-base panel picks it up as a pending query.
 * Kept component-free so the routing decisions are testable without mounting
 * anything.
 */

/** A question handed to the Ask AI panel from outside. `seq` distinguishes two
 *  submissions of the same text: the panel consumes each seq at most once. */
export interface PendingAskQuery {
  readonly text: string
  readonly seq: number
  /** 'palette' is the user's own just-typed text; 'deep-link' is web-supplied */
  readonly origin: AskQueryOrigin
}

/**
 * Where a pending query came from. The palette fallback is text the user just
 * typed and confirmed, so it submits itself; a `genoffice://ask?q=…` query can
 * be injected by any web page through location.href, so it only parks in the
 * composer draft — the person sends it with an explicit Enter.
 */
export type AskQueryOrigin = 'palette' | 'deep-link'

export type PendingAskPlan =
  | { readonly consume: false }
  | {
      readonly consume: true
      readonly seq: number
      /** submit = send now; park = an answer is streaming, wait in the composer; drop = blank text */
      readonly action: 'submit' | 'park' | 'drop'
      readonly text: string
    }

/**
 * What the panel does with a pending query. An already-consumed seq and an
 * absent query are no-ops; a blank text is consumed and dropped; a
 * deep-link-origin query parks in the composer draft whatever the panel's
 * state (web-supplied text never sends itself); a palette query submits when
 * the panel is idle and parks while an answer streams, so it is never lost —
 * the user sends it when the current answer lands.
 */
export function planPendingAskQuery(
  pending: PendingAskQuery | null,
  lastConsumedSeq: number,
  busy: boolean,
): PendingAskPlan {
  if (!pending || pending.seq === lastConsumedSeq) return { consume: false }
  const text = pending.text.trim()
  if (!text) return { consume: true, seq: pending.seq, action: 'drop', text: '' }
  if (pending.origin === 'deep-link')
    return { consume: true, seq: pending.seq, action: 'park', text }
  return { consume: true, seq: pending.seq, action: busy ? 'park' : 'submit', text }
}

/** a deep link's query is web-supplied text: cap it before it reaches the composer */
const MAX_DEEP_LINK_QUERY = 2000

/**
 * The ask view's deep-link payload (`{ view: 'ask', query?: string }`),
 * normalized; null when the link addresses a different view. A missing or
 * non-string query reads as '' — open the panel without injecting a question —
 * and an over-long one is truncated to MAX_DEEP_LINK_QUERY characters.
 */
export function normalizeAskDeepLink(payload: unknown): { readonly query: string } | null {
  if (typeof payload !== 'object' || payload === null) return null
  if ((payload as { view?: unknown }).view !== 'ask') return null
  const query = (payload as { query?: unknown }).query
  return { query: typeof query === 'string' ? query.slice(0, MAX_DEEP_LINK_QUERY) : '' }
}
