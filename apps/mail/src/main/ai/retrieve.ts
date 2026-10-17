import type { MailEnvelope, MailMessageRef } from '@genoffice/mail-protocols'
import { redactForPrompt } from '../sync/redact'
import { normalizeAiScope } from './scope'
import type { AiScope } from './scope'
import { sharedAiRetrieveRateLimiter } from './rate-limit'
import type { AiRateLimiter } from './rate-limit'

/**
 * Retrieval core for the mail:ai-retrieve channel (P5 spec section 4).
 * One channel, three ops: triage runs on headers + FTS, bodies resolve
 * lazily for the explicitly cited set only. Envelopes are headers + snippet
 * by construction (toEnvelope never emits body text); body text crossing
 * into the response is redacted for prompts. Storage stays behind the
 * AiRetrieveStorePort so tests inject stub stores and production serves
 * the real per-account SQLite stores; cross-account joins are impossible
 * because every lookup opens one account at a time.
 */

export type AiRetrieveOp = 'triage-headers' | 'fetch-bodies' | 'cite-resolve'

export interface AiRetrieveRequest {
  scope: AiScope
  op: AiRetrieveOp
  query?: string
  folderId?: string
  limit?: number
  refs?: MailMessageRef[]
}

export interface AiBodyEntry {
  ref: MailMessageRef
  /** Scope account that held this ref (refs carry no account id). */
  accountId: string
  /** Cited body text, redacted for prompts. */
  text: string
}

export interface AiCitation {
  uid: string
  folderId: string
  msgIdHeader?: string
  /** Scope account that held this ref (refs carry no account id). */
  accountId: string
}

export interface AiRetrieveResponse {
  /** Echo of the request scope; the renderer discards on mismatch. */
  scope: AiScope
  /** Headers + snippet only; bodies never appear here. */
  envelopes: MailEnvelope[]
  /** fetch-bodies only, cited set only, redacted for prompts. */
  bodies?: AiBodyEntry[]
  citations: AiCitation[]
}

/** Per-account storage surface the retrieval core consumes. */
export interface AiRetrieveStorePort {
  search(
    accountId: string,
    query: string,
    opts: { limit: number; folderId?: string },
  ): MailEnvelope[]
  body(accountId: string, folderId: string, uid: string): string | undefined
  envelope(accountId: string, folderId: string, uid: string): MailEnvelope | undefined
}

export interface AiRetrieveDeps {
  /** Account ids the main process knows (demo seed, configured stores). */
  knownAccountIds(): string[] | Promise<string[]>
  /** Per-account store; called only for scope members. */
  storeFor(accountId: string): AiRetrieveStorePort | Promise<AiRetrieveStorePort>
  /**
   * Abuse + cost guard (H1 P-9). Defaults to the shared per-account
   * token bucket; tests inject a fake-clock limiter here.
   */
  rateLimiter?: AiRateLimiter
}

/** Default triage cap (FTS_DEFAULT_LIMIT, store.ts). */
export const AI_RETRIEVE_DEFAULT_LIMIT = 50
/** Hard cap per call (matches the IPC schema bound). */
export const AI_RETRIEVE_MAX_LIMIT = 200
/** Hard cap on explicitly cited refs per call. */
export const AI_RETRIEVE_MAX_REFS = 50

function badRequest(message: string): Error {
  return new Error(`mail: ${message}`)
}

function citationFor(accountId: string, env: MailEnvelope): AiCitation {
  const citation: AiCitation = { uid: env.ref.uid, folderId: env.ref.folderId, accountId }
  if (env.ref.msgIdHeader !== undefined) citation.msgIdHeader = env.ref.msgIdHeader
  return citation
}

/**
 * Run one scoped retrieval. Rejects unknown scope accounts before opening
 * any store; every op touches only scope members, so mail outside the
 * picker selection can never enter the response. Each call is stateless
 * (no server-side session): the renderer discards stale results via the
 * echoed scope, so a newer picker selection simply supersedes in-flight
 * calls — that echo check IS the query-cancellation contract (H6 seam).
 */
export async function aiRetrieve(
  request: AiRetrieveRequest,
  deps: AiRetrieveDeps,
): Promise<AiRetrieveResponse> {
  const scope = normalizeAiScope(request.scope)
  const known = new Set(await deps.knownAccountIds())
  for (const accountId of scope.accountIds) {
    if (!known.has(accountId)) throw badRequest('unknown account in AI scope')
  }
  ;(deps.rateLimiter ?? sharedAiRetrieveRateLimiter()).check(scope.accountIds)
  const limit = Math.min(
    Math.max(Math.trunc(request.limit ?? AI_RETRIEVE_DEFAULT_LIMIT), 1),
    AI_RETRIEVE_MAX_LIMIT,
  )

  if (request.op === 'triage-headers') {
    const query = request.query?.trim() ?? ''
    if (!query) throw badRequest('triage-headers requires a query')
    const envelopes: MailEnvelope[] = []
    const citations: AiCitation[] = []
    for (const accountId of scope.accountIds) {
      const store = await deps.storeFor(accountId)
      const hits = store.search(accountId, query, { limit, folderId: request.folderId })
      for (const env of hits) {
        envelopes.push(env)
        citations.push(citationFor(accountId, env))
      }
    }
    return { scope, envelopes: envelopes.slice(0, limit), citations: citations.slice(0, limit) }
  }

  if (request.op === 'fetch-bodies' || request.op === 'cite-resolve') {
    const refs = request.refs ?? []
    if (refs.length < 1) throw badRequest(`${request.op} requires refs`)
    if (refs.length > AI_RETRIEVE_MAX_REFS) throw badRequest('too many refs')
    const envelopes: MailEnvelope[] = []
    const citations: AiCitation[] = []
    const bodies: AiBodyEntry[] = []
    for (const ref of refs) {
      // D-H1-C3 exact-match preference: a ref may carry the accountId of
      // the citation that produced it. A hint naming a scope member that
      // holds the ref wins over scope order; unknown or stale hints fail
      // loud instead of silently resolving another account's mail.
      const hint = ref.accountId
      if (hint !== undefined) {
        if (!scope.accountIds.includes(hint)) {
          throw badRequest('stale account hint in message ref')
        }
        const store = await deps.storeFor(hint)
        const env = store.envelope(hint, ref.folderId, ref.uid)
        if (!env) throw badRequest('stale account hint in message ref')
        envelopes.push(env)
        citations.push(citationFor(hint, env))
        if (request.op === 'fetch-bodies') {
          const text = store.body(hint, ref.folderId, ref.uid) ?? ''
          bodies.push({ ref: env.ref, accountId: hint, text: redactForPrompt(text) })
        } else {
          // cite-resolve surfaces ambiguity: the hinted match first, then
          // every other in-scope holder (all in-scope, so no leak).
          for (const accountId of scope.accountIds) {
            if (accountId === hint) continue
            const other = await deps.storeFor(accountId)
            const otherEnv = other.envelope(accountId, ref.folderId, ref.uid)
            if (!otherEnv) continue
            envelopes.push(otherEnv)
            citations.push(citationFor(accountId, otherEnv))
          }
        }
        continue
      }
      // No hint: deterministic first-scope-match fallback (scope order =
      // picker order). fetch-bodies returns the single resolved body while
      // cite-resolve returns every in-scope match.
      for (const accountId of scope.accountIds) {
        const store = await deps.storeFor(accountId)
        const env = store.envelope(accountId, ref.folderId, ref.uid)
        if (!env) continue
        envelopes.push(env)
        citations.push(citationFor(accountId, env))
        if (request.op === 'fetch-bodies') {
          const text = store.body(accountId, ref.folderId, ref.uid) ?? ''
          bodies.push({ ref: env.ref, accountId, text: redactForPrompt(text) })
          break
        }
      }
    }
    const response: AiRetrieveResponse = { scope, envelopes, citations }
    if (request.op === 'fetch-bodies') response.bodies = bodies
    return response
  }

  throw badRequest('unknown AI retrieve op')
}
