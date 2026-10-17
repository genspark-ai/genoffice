/**
 * Account-scoping envelope for GenOffice Mail AI calls (P5 spec section 3).
 *
 * The picker-selected account set travels with EVERY AI retrieval call and
 * is re-checked where results render: main rejects unknown accountIds and
 * opens only per-account stores, while the renderer discards results whose
 * echoed scope no longer matches the live picker (stale-scope discard).
 */

/** Picker-selected accounts authorizing one AI call; non-empty, max 8. */
export interface AiScope {
  accountIds: string[]
}

/** Upper bound on accounts per AI call (spec section 3). */
export const AI_SCOPE_MAX_ACCOUNTS = 8

function badScope(message: string): Error {
  return new Error(`mail: bad AI scope (${message})`)
}

/**
 * Validate and normalize a raw scope payload. Returns a deduped copy;
 * throws a mail-prefixed error on any violation (empty set, over 8
 * accounts, non-string or over-long ids, unknown extra fields are left
 * to the zod schema at the IPC boundary).
 */
export function normalizeAiScope(scope: unknown): AiScope {
  if (!scope || typeof scope !== 'object') throw badScope('not an object')
  const ids = (scope as { accountIds?: unknown }).accountIds
  if (!Array.isArray(ids) || ids.length < 1) throw badScope('accountIds must be non-empty')
  if (ids.length > AI_SCOPE_MAX_ACCOUNTS) throw badScope('at most 8 accounts')
  const seen: string[] = []
  for (const id of ids) {
    if (typeof id !== 'string' || id.length < 1 || id.length > 256) {
      throw badScope('account ids must be 1-256 char strings')
    }
    if (!seen.includes(id)) seen.push(id)
  }
  return { accountIds: seen }
}

/**
 * True when a response scope no longer matches the live picker selection
 * (order-insensitive). The renderer discards stale results instead of
 * rendering mail from accounts the user has since deselected.
 */
export function isScopeStale(resultScope: AiScope, liveScope: AiScope): boolean {
  if (resultScope.accountIds.length !== liveScope.accountIds.length) return true
  return resultScope.accountIds.some((id) => !liveScope.accountIds.includes(id))
}
