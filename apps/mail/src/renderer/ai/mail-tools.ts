import type {
  AgentToolCall,
  AgentToolDef,
  ExecutedToolCall,
  ToolExecution,
} from '@genoffice/agent-core'
import type { MailEnvelope, MailMessageRef } from '@genoffice/mail-protocols'

/**
 * Shared mail AI tool definitions and executor (K lane owns this file).
 * Every tool calls ONLY the W-lane retrieval bridge
 * (`window.mailApi.aiRetrieve`, spec section 6: sole channel
 * `mail:ai-retrieve`). Bodies and subjects never enter logs: this module
 * emits no log output at all.
 */

// ── Spec section 3: account-scoping envelope ────────────────────────────────

/** Picker-selected account set; travels with EVERY AI call (1-8 ids). */
export interface AiScope {
  accountIds: string[]
}

// ── Spec section 4: retrieval API (Farah/Nadia joint, exact shape) ──────────

export type AiRetrieveOp = 'triage-headers' | 'fetch-bodies' | 'cite-resolve'

export interface AiRetrieveRequest {
  scope: AiScope
  op: AiRetrieveOp
  /** triage-headers: FTS query, headers+snippet ranked. */
  query?: string
  folderId?: string
  /** Default 50 (FTS_DEFAULT_LIMIT). */
  limit?: number
  /** fetch-bodies / cite-resolve: explicit user-cited set. */
  refs?: MailMessageRef[]
}

export interface AiRetrieveCitation {
  uid: string
  folderId: string
  msgIdHeader?: string
}

export interface AiRetrieveResponse {
  /** Echo — the caller discards the result on mismatch (stale scope). */
  scope: AiScope
  /** Headers+snippet ONLY (toEnvelope shape); bodies NEVER here. */
  envelopes: MailEnvelope[]
  /** fetch-bodies only, cited set only. */
  bodies?: { ref: MailMessageRef; text: string }[]
  citations: AiRetrieveCitation[]
}

/**
 * W-lane bridge method, resolved structurally: K-lane code consumes the
 * bridge but never edits the shared bridge types (W lane owns those).
 */
export interface MailAiBridge {
  aiRetrieve(request: AiRetrieveRequest): Promise<AiRetrieveResponse>
}

/** Live bridge when the W lane has bridged it, null otherwise (tests inject a fake). */
export function resolveAiBridge(): MailAiBridge | null {
  if (typeof window === 'undefined') return null
  const api = (window as unknown as { mailApi?: { aiRetrieve?: unknown } }).mailApi
  if (!api || typeof api.aiRetrieve !== 'function') return null
  // Bind to the holder: class-based bridges (fakes, real impls) read
  // instance state, and the extractor above would otherwise drop `this`.
  const aiRetrieve = (api.aiRetrieve as MailAiBridge['aiRetrieve']).bind(api)
  return { aiRetrieve: (request) => aiRetrieve(request) }
}

// ── Scope + request guards (mirror the spec refinements) ────────────────────

const MAX_SCOPE_ACCOUNTS = 8
const MAX_QUERY_CHARS = 500
const MAX_REFS = 50
const MAX_LIMIT = 200
const DEFAULT_LIMIT = 50

export function validateScope(scope: AiScope): string | null {
  if (!scope || !Array.isArray(scope.accountIds)) return 'scope must carry accountIds'
  if (scope.accountIds.length < 1) return 'scope must select at least one account'
  if (scope.accountIds.length > MAX_SCOPE_ACCOUNTS) {
    return `scope selects too many accounts (max ${MAX_SCOPE_ACCOUNTS})`
  }
  for (const id of scope.accountIds) {
    if (typeof id !== 'string' || id.length < 1 || id.length > 256)
      return 'scope holds a bad account id'
  }
  return null
}

/** Scope-echo comparison: same id set regardless of order. */
export function isSameScope(a: AiScope, b: AiScope): boolean {
  if (a.accountIds.length !== b.accountIds.length) return false
  const left = [...a.accountIds].sort()
  const right = [...b.accountIds].sort()
  return left.every((id, i) => id === right[i])
}

function scopeLabel(scope: AiScope): string {
  return scope.accountIds.join(',')
}

function clampLimit(raw: unknown): number {
  const n = typeof raw === 'number' ? Math.floor(raw) : DEFAULT_LIMIT
  if (!Number.isFinite(n)) return DEFAULT_LIMIT
  return Math.min(Math.max(n, 1), MAX_LIMIT)
}

function readString(input: Record<string, unknown>, key: string): string {
  const v = input[key]
  return typeof v === 'string' ? v : ''
}

function fail(output: string, summary: string): ToolExecution {
  return { output, isError: true, summary }
}

// ── Citation helpers ────────────────────────────────────────────────────────

/** Citation token emitted alongside every quoted mail fact: `[uid]`. */
export function citeToken(uid: string): string {
  return `[${uid}]`
}

function refKey(ref: Pick<MailMessageRef, 'uid' | 'folderId'>): string {
  return `${ref.folderId}/${ref.uid}`
}

/**
 * Bracket tokens that are NOT message citations: redaction placeholders the
 * W lane emits inside bodies (`[address]`, `[secret]`, ...) plus the
 * instructional `[uid]` placeholder from tool copy. Anything else shaped
 * like `[token]` alongside factual claims is treated as an unverified
 * citation (H3 prompt-injection fixture: citation forgery).
 */
const NON_CITATION_TOKENS = new Set([
  'uid',
  'address',
  'secret',
  'ip',
  'blob',
  'redacted',
  'redacted-key',
  'truncated',
  'message-id',
])

const UID_LIKE_TOKEN_RE = /\[([^\][\s()]+)\]/g
const UID_LIKE_SHAPE_RE = /^[A-Za-z0-9._:@-]{1,64}$/

/** Bracketed uid-like tokens in text that name no retrieved message. */
function unknownCitationTokens(stripped: string, cited: ReadonlySet<string>): string[] {
  const unknown: string[] = []
  UID_LIKE_TOKEN_RE.lastIndex = 0
  let match: RegExpExecArray | null
  while ((match = UID_LIKE_TOKEN_RE.exec(stripped)) !== null) {
    const token = match[1]
    if (cited.has(token)) continue
    if (NON_CITATION_TOKENS.has(token.toLowerCase())) continue
    if (!UID_LIKE_SHAPE_RE.test(token)) continue
    if (!unknown.includes(token)) unknown.push(token)
  }
  return unknown
}

function claimLinesOf(stripped: string): string[] {
  return stripped
    .split('\n')
    .map((line) => line.trim())
    .filter(
      (line) =>
        /^(?:[-*•>]|\d+[.)])\s+\S/.test(line) || (line.length > 200 && /[.!?。]$/.test(line)),
    )
}

export interface CitationGuard {
  /** Record uids the model may cite (populated on each successful retrieval). */
  noteCited(uid: string): void
  /**
   * verifyResponse for ask-mail/summarize: reject final text that states
   * retrieved facts without a recorded `[uid]` citation, forcing one
   * corrective turn. Also rejects uid-like bracket tokens that name no
   * retrieved message (citation forgery). Returns the correction or null
   * to accept.
   */
  verifyCited(finalText: string, executed: readonly ExecutedToolCall[]): string | null
}

/**
 * Guard factory shared by the summarize and search (ask-mail) skills.
 * `toolNames` are the retrieval tools whose success obliges citations.
 */
export function createCitationGuard(toolNames: readonly string[]): CitationGuard {
  const cited = new Set<string>()
  return {
    noteCited: (uid: string) => {
      cited.add(uid)
    },
    verifyCited: (finalText: string, executed: readonly ExecutedToolCall[]) => {
      const usedRetrieval = executed.some((call) => call.ok && toolNames.includes(call.name))
      if (!usedRetrieval || cited.size === 0) return null
      const stripped = finalText
        .replace(/```[\s\S]*?```/g, ' ')
        .replace(/`[^`]*`/g, ' ')
        .replace(/\[[^\]]*\]\([^)]*\)/g, ' ')
      const hasCitation = [...cited].some((uid) => stripped.includes(citeToken(uid)))
      const claimLines = claimLinesOf(stripped)
      if (hasCitation) {
        // Forgery check: claims accompanied by uid-like bracket tokens that
        // name no retrieved message (injected `[evil-999]`, cross-account
        // `[b-m1]`) are rejected even when a valid citation is also present.
        const unknown = unknownCitationTokens(stripped, cited)
        if (unknown.length === 0) return null
        if (claimLines.length === 0 && stripped.trim().length <= 200) return null
        return (
          'The reply cites messages outside the retrieved set ' +
          `(${unknown
            .slice(0, 8)
            .map((token) => `[${token}]`)
            .join(' ')}). ` +
          'Restate every factual claim with its per-claim message-id citation ' +
          `(${[...cited].slice(0, 8).map(citeToken).join(' ')}) and keep uncited text to chatter only.`
        )
      }
      if (claimLines.length === 0 && stripped.trim().length <= 200) return null
      return (
        'The reply states mail facts without a message citation. ' +
        'Restate every factual claim with its per-claim message-id citation ' +
        `(${[...cited].slice(0, 8).map(citeToken).join(' ')}) and keep uncited text to chatter only.`
      )
    },
  }
}

// ── Tool catalogue (names globally unique across the four skills) ────────────

export const MAIL_AGENT_TOOLS: AgentToolDef[] = [
  {
    name: 'rank_folder',
    description:
      'Rank one folder into digest buckets (priority / needs-reply / newsletter / noise) from headers + FTS only. Never returns bodies.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'FTS query ranking the folder headers' },
        folderId: { type: 'string', description: 'Folder to rank (default inbox)' },
        limit: { type: 'integer', description: 'Max headers to rank (default 50, max 200)' },
      },
      required: ['query'],
    },
  },
  {
    name: 'apply_label',
    description:
      'Propose a label for listed messages. Returns a proposal for the user to confirm — never moves, flags, or labels anything itself.',
    inputSchema: {
      type: 'object',
      properties: {
        uids: { type: 'array', description: 'Message uids to label', items: { type: 'string' } },
        folderId: { type: 'string', description: 'Folder holding the messages' },
        label: { type: 'string', description: 'Label to propose' },
      },
      required: ['uids', 'folderId', 'label'],
    },
  },
  {
    name: 'summarize_thread',
    description:
      'Fetch bodies for an explicitly cited message set and summarize them. Every summary bullet must carry its [uid] citation.',
    inputSchema: {
      type: 'object',
      properties: {
        uids: {
          type: 'array',
          description: 'Message uids to summarize',
          items: { type: 'string' },
        },
        folderId: { type: 'string', description: 'Folder holding the messages' },
      },
      required: ['uids', 'folderId'],
    },
  },
  {
    name: 'catch_up_folder',
    description:
      'Catch up on a folder: rank its headers for a topic, then read bodies of the top hits. Bodies are fetched only for the cited top set.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Topic ranking the folder headers' },
        folderId: { type: 'string', description: 'Folder to catch up on (default inbox)' },
        limit: { type: 'integer', description: 'Max headers to rank (default 20, max 200)' },
      },
      required: ['query'],
    },
  },
  {
    name: 'draft_from_bullets',
    description:
      'Assemble an editable draft from bullet points. Returns text for the compose box — never sends anything.',
    inputSchema: {
      type: 'object',
      properties: {
        bullets: {
          type: 'array',
          description: 'Points the draft must cover',
          items: { type: 'string' },
        },
        greeting: { type: 'string', description: 'Opening line (default Hello,)' },
        signoff: { type: 'string', description: 'Closing line (default Best regards,)' },
      },
      required: ['bullets'],
    },
  },
  {
    name: 'rewrite_tone',
    description:
      'Prepare a tone rewrite: returns the original text plus the tone target so the reply itself carries the rewrite. Never sends anything.',
    inputSchema: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'Draft text to restyle' },
        tone: {
          type: 'string',
          description: 'Target tone',
          enum: ['shorter', 'friendlier', 'formal', 'direct'],
        },
      },
      required: ['text', 'tone'],
    },
  },
  {
    name: 'reply_from_thread',
    description:
      'Quote cited thread context for a reply draft. Returns quoted context with [uid] citations; the reply text stays editable and is never sent.',
    inputSchema: {
      type: 'object',
      properties: {
        uids: {
          type: 'array',
          description: 'Thread message uids to quote',
          items: { type: 'string' },
        },
        folderId: { type: 'string', description: 'Folder holding the thread' },
        focus: { type: 'string', description: 'What the reply should address' },
      },
      required: ['uids', 'folderId'],
    },
  },
  {
    name: 'search_mail',
    description:
      'Search mail headers-first over FTS. Returns envelopes (headers + snippet) with [uid] refs; bodies stay unfetched until quoted.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'FTS query' },
        folderId: { type: 'string', description: 'Restrict to one folder' },
        limit: { type: 'integer', description: 'Max hits (default 50, max 200)' },
      },
      required: ['query'],
    },
  },
  {
    name: 'quote_message',
    description:
      'Resolve citable quote cards for explicitly cited messages (header-only, no bodies). Every card carries its [uid] citation payload.',
    inputSchema: {
      type: 'object',
      properties: {
        uids: { type: 'array', description: 'Message uids to quote', items: { type: 'string' } },
        folderId: { type: 'string', description: 'Folder holding the messages' },
      },
      required: ['uids', 'folderId'],
    },
  },
]

// ── Rendering (headers + snippet only; bodies only where the op allows) ─────

const DIGEST_MAX_CHARS = 12000
const BODY_EXCERPT_CHARS = 4000

function envelopeLine(env: MailEnvelope): string {
  const seen = env.flags.seen ? 'seen' : 'unseen'
  const flagged = env.flags.flagged ? ' flagged' : ''
  const labels = env.labels.length > 0 ? ` labels:${env.labels.join(',')}` : ''
  return (
    `- ${citeToken(env.ref.uid)} ${env.date} ${env.subject} | ` +
    `from: ${env.from.join(', ')} | folder: ${env.ref.folderId} | ${seen}${flagged}${labels}`
  )
}

const NEWSLETTER_HINT =
  /unsubscrib|newsletter|digest|promotion|promo|\bsale\b|\bdeal\b|\boffer\b|weekly|marketing/i

type Bucket = 'priority' | 'needs-reply' | 'newsletter' | 'noise'

function bucketOf(env: MailEnvelope): Bucket {
  if (env.flags.flagged || env.labels.includes('priority') || env.labels.includes('important')) {
    return 'priority'
  }
  if (NEWSLETTER_HINT.test(`${env.subject} ${env.snippet}`)) return 'newsletter'
  if (!env.flags.seen) return 'needs-reply'
  return 'noise'
}

function digestOutput(envelopes: MailEnvelope[], scope: AiScope, guard?: CitationGuard): string {
  const buckets: Record<Bucket, MailEnvelope[]> = {
    priority: [],
    'needs-reply': [],
    newsletter: [],
    noise: [],
  }
  for (const env of envelopes) buckets[bucketOf(env)].push(env)
  const lines = [`Triage digest (${envelopes.length} messages, scope: ${scopeLabel(scope)}):`]
  for (const name of ['priority', 'needs-reply', 'newsletter', 'noise'] as const) {
    lines.push(`## ${name} (${buckets[name].length})`)
    for (const env of buckets[name]) lines.push(envelopeLine(env))
  }
  if (guard) for (const env of envelopes) guard.noteCited(env.ref.uid)
  const joined = lines.join('\n')
  return joined.length > DIGEST_MAX_CHARS
    ? `${joined.slice(0, DIGEST_MAX_CHARS)}\n… (truncated — narrow the query or limit)`
    : joined
}

function hitsOutput(envelopes: MailEnvelope[], scope: AiScope, guard?: CitationGuard): string {
  if (guard) for (const env of envelopes) guard.noteCited(env.ref.uid)
  if (envelopes.length === 0) return `No mail hits (scope: ${scopeLabel(scope)}).`
  return [
    `Mail hits (${envelopes.length}, scope: ${scopeLabel(scope)}):`,
    ...envelopes.map(envelopeLine),
  ].join('\n')
}

function bodiesOutput(
  bodies: { ref: MailMessageRef; text: string }[],
  scope: AiScope,
  guard?: CitationGuard,
): string {
  if (guard) for (const item of bodies) guard.noteCited(item.ref.uid)
  return (
    bodies
      .map((item) => {
        const text =
          item.text.length > BODY_EXCERPT_CHARS
            ? `${item.text.slice(0, BODY_EXCERPT_CHARS)}\n… (excerpt truncated)`
            : item.text
        return `### ${citeToken(item.ref.uid)} folder:${item.ref.folderId}\n${text}`
      })
      .join('\n\n') || `(no bodies returned, scope: ${scopeLabel(scope)})`
  )
}

// ── Executor ────────────────────────────────────────────────────────────────

export interface MailToolContext {
  scope: AiScope
  bridge: MailAiBridge
  guard?: CitationGuard
  signal?: AbortSignal
}

function stopped(signal?: AbortSignal): boolean {
  return signal?.aborted === true
}

async function retrieve(
  ctx: MailToolContext,
  request: AiRetrieveRequest,
): Promise<AiRetrieveResponse | { error: string }> {
  if (stopped(ctx.signal)) return { error: 'stopped by the user' }
  const scopeError = validateScope(request.scope)
  if (scopeError) return { error: scopeError }
  let response: AiRetrieveResponse
  try {
    response = await ctx.bridge.aiRetrieve(request)
  } catch (err) {
    return {
      error: `retrieval failed (service error, not an empty result): ${err instanceof Error ? err.message : 'unknown error'}`,
    }
  }
  if (stopped(ctx.signal)) return { error: 'stopped by the user' }
  // Stale-scope results are discarded, never rendered.
  if (!response || !isSameScope(response.scope, request.scope)) {
    return { error: 'stale scope: the retrieval scope changed under this call; discard and retry' }
  }
  return response
}

/** Keep only bodies the caller explicitly cited; extras are discarded, never shown. */
function intersectBodies(
  bodies: { ref: MailMessageRef; text: string }[] | undefined,
  wanted: Set<string>,
): { ref: MailMessageRef; text: string }[] {
  return (bodies ?? []).filter((item) => wanted.has(refKey(item.ref)))
}

function uidsOf(input: Record<string, unknown>): string[] | null {
  const v = input.uids
  if (!Array.isArray(v)) return null
  const out: string[] = []
  for (const item of v) {
    if (typeof item !== 'string' || item.length < 1 || item.length > 256) return null
    out.push(item)
  }
  return out
}

/**
 * Route one tool call over the retrieval bridge (mirror of markdown
 * `executeTool`). Pure tools (draft assembly) never touch the bridge.
 */
export async function executeMailTool(
  call: AgentToolCall,
  ctx: MailToolContext,
): Promise<ToolExecution> {
  const input = call.input ?? {}
  switch (call.name) {
    case 'rank_folder': {
      const query = readString(input, 'query').trim()
      if (!query) return fail('query must not be empty', call.name)
      if (query.length > MAX_QUERY_CHARS) return fail('query is too long', call.name)
      const folderId = readString(input, 'folderId').trim() || undefined
      const response = await retrieve(ctx, {
        scope: ctx.scope,
        op: 'triage-headers',
        query,
        folderId,
        limit: clampLimit(input.limit),
      })
      if ('error' in response) return fail(response.error, call.name)
      // Triage is headers + snippet by construction: bodies can never arrive here.
      return {
        output: digestOutput(response.envelopes, ctx.scope, ctx.guard),
        mutated: false,
        summary: call.name,
      }
    }

    case 'apply_label': {
      const uids = uidsOf(input)
      const folderId = readString(input, 'folderId').trim()
      const label = readString(input, 'label').trim()
      if (!uids || uids.length < 1 || uids.length > MAX_REFS) {
        return fail('uids must list 1-50 message uids', call.name)
      }
      if (!folderId) return fail('folderId must not be empty', call.name)
      if (!label) return fail('label must not be empty', call.name)
      const scopeError = validateScope(ctx.scope)
      if (scopeError) return fail(scopeError, call.name)
      // User-confirmed ONLY: the skill proposes; the human applies the label
      // through the existing UI. Nothing is moved, flagged, or relabelled here.
      return {
        output: [
          `Label proposal (NOT applied — confirm in the UI first, scope: ${scopeLabel(ctx.scope)}):`,
          `label "${label}" on ${uids.length} message(s) in ${folderId}: ${uids.map(citeToken).join(' ')}`,
        ].join('\n'),
        mutated: false,
        summary: call.name,
      }
    }

    case 'summarize_thread': {
      const uids = uidsOf(input)
      const folderId = readString(input, 'folderId').trim()
      if (!uids || uids.length < 1 || uids.length > 20) {
        return fail('uids must list 1-20 message uids', call.name)
      }
      if (!folderId) return fail('folderId must not be empty', call.name)
      const refs: MailMessageRef[] = uids.map((uid) => ({ uid, folderId, threadId: '' }))
      const response = await retrieve(ctx, { scope: ctx.scope, op: 'fetch-bodies', refs })
      if ('error' in response) return fail(response.error, call.name)
      const wanted = new Set(refs.map(refKey))
      const bodies = intersectBodies(response.bodies, wanted)
      if (bodies.length === 0 && (response.bodies?.length ?? 0) > 0) {
        return fail('bridge returned bodies outside the cited set; discarded', call.name)
      }
      return {
        output:
          `Thread bodies (${bodies.length}/${refs.length} cited, scope: ${scopeLabel(ctx.scope)}). ` +
          `Summarize below: one bullet per message, every bullet ends with its ${'[uid]'} citation.\n\n` +
          bodiesOutput(bodies, ctx.scope, ctx.guard),
        mutated: false,
        summary: call.name,
      }
    }

    case 'catch_up_folder': {
      const query = readString(input, 'query').trim()
      if (!query) return fail('query must not be empty', call.name)
      if (query.length > MAX_QUERY_CHARS) return fail('query is too long', call.name)
      const folderId = readString(input, 'folderId').trim() || undefined
      const ranked = await retrieve(ctx, {
        scope: ctx.scope,
        op: 'triage-headers',
        query,
        folderId,
        limit: clampLimit(input.limit ?? 20),
      })
      if ('error' in ranked) return fail(ranked.error, call.name)
      const top = ranked.envelopes.slice(0, 10)
      if (top.length === 0) {
        return {
          output: `Nothing to catch up on (scope: ${scopeLabel(ctx.scope)}).`,
          mutated: false,
          summary: call.name,
        }
      }
      const refs: MailMessageRef[] = top.map((env) => ({ ...env.ref }))
      const fetched = await retrieve(ctx, { scope: ctx.scope, op: 'fetch-bodies', refs })
      if ('error' in fetched) return fail(fetched.error, call.name)
      const wanted = new Set(refs.map(refKey))
      const bodies = intersectBodies(fetched.bodies, wanted)
      return {
        output:
          `${digestOutput(ranked.envelopes, ctx.scope)}\n\n` +
          `Top-${bodies.length} bodies (explicit catch-up, scope: ${scopeLabel(ctx.scope)}). ` +
          `Summarize with one cited bullet per message.\n\n` +
          bodiesOutput(bodies, ctx.scope, ctx.guard),
        mutated: false,
        summary: call.name,
      }
    }

    case 'draft_from_bullets': {
      const v = input.bullets
      if (!Array.isArray(v) || v.length < 1 || v.length > 20) {
        return fail('bullets must list 1-20 points', call.name)
      }
      const bullets: string[] = []
      for (const item of v) {
        if (typeof item !== 'string' || !item.trim() || item.length > 500) {
          return fail('every bullet must be non-empty text (max 500 chars)', call.name)
        }
        bullets.push(item.trim())
      }
      const greeting = readString(input, 'greeting').trim() || 'Hello,'
      const signoff = readString(input, 'signoff').trim() || 'Best regards,'
      return {
        output: [greeting, '', ...bullets.map((b) => `- ${b}`), '', signoff].join('\n'),
        mutated: false,
        summary: call.name,
      }
    }

    case 'rewrite_tone': {
      const text = readString(input, 'text')
      const tone = readString(input, 'tone').trim()
      if (!text.trim()) return fail('text must not be empty', call.name)
      if (text.length > 8000) return fail('text is too long (max 8000 chars)', call.name)
      if (!['shorter', 'friendlier', 'formal', 'direct'].includes(tone)) {
        return fail('tone must be shorter, friendlier, formal, or direct', call.name)
      }
      // The rewrite itself happens in the reply prose; the tool scopes the
      // source text so the model restyles only what the user supplied.
      return {
        output: `Rewrite the draft below in a ${tone} tone. Reply with ONLY the rewritten text (editable, never sent).\n\n---\n${text}`,
        mutated: false,
        summary: call.name,
      }
    }

    case 'reply_from_thread': {
      const uids = uidsOf(input)
      const folderId = readString(input, 'folderId').trim()
      if (!uids || uids.length < 1 || uids.length > 10) {
        return fail('uids must list 1-10 thread message uids', call.name)
      }
      if (!folderId) return fail('folderId must not be empty', call.name)
      const focus = readString(input, 'focus').trim()
      const refs: MailMessageRef[] = uids.map((uid) => ({ uid, folderId, threadId: '' }))
      const response = await retrieve(ctx, { scope: ctx.scope, op: 'fetch-bodies', refs })
      if ('error' in response) return fail(response.error, call.name)
      const wanted = new Set(refs.map(refKey))
      const bodies = intersectBodies(response.bodies, wanted)
      return {
        output:
          `Thread context for an editable reply (scope: ${scopeLabel(ctx.scope)}${focus ? `, focus: ${focus}` : ''}). ` +
          `Draft the reply below as plain text with ${'[uid]'} citations where facts are reused — never send.\n\n` +
          bodiesOutput(bodies, ctx.scope, ctx.guard),
        mutated: false,
        summary: call.name,
      }
    }

    case 'search_mail': {
      const query = readString(input, 'query').trim()
      if (!query) return fail('query must not be empty', call.name)
      if (query.length > MAX_QUERY_CHARS) return fail('query is too long', call.name)
      const folderId = readString(input, 'folderId').trim() || undefined
      const response = await retrieve(ctx, {
        scope: ctx.scope,
        op: 'triage-headers',
        query,
        folderId,
        limit: clampLimit(input.limit),
      })
      if ('error' in response) return fail(response.error, call.name)
      return {
        output: hitsOutput(response.envelopes, ctx.scope, ctx.guard),
        mutated: false,
        summary: call.name,
      }
    }

    case 'quote_message': {
      const uids = uidsOf(input)
      const folderId = readString(input, 'folderId').trim()
      if (!uids || uids.length < 1 || uids.length > 10) {
        return fail('uids must list 1-10 message uids', call.name)
      }
      if (!folderId) return fail('folderId must not be empty', call.name)
      const refs: MailMessageRef[] = uids.map((uid) => ({ uid, folderId, threadId: '' }))
      // Header-only resolve: citable ids without bodies.
      const response = await retrieve(ctx, { scope: ctx.scope, op: 'cite-resolve', refs })
      if ('error' in response) return fail(response.error, call.name)
      const wanted = new Set(uids)
      const cards = response.citations.filter((cite) => wanted.has(cite.uid))
      const byUid = new Map(response.envelopes.map((env) => [env.ref.uid, env]))
      if (ctx.guard) for (const card of cards) ctx.guard.noteCited(card.uid)
      if (cards.length === 0) {
        return {
          output: `No citable messages resolved (scope: ${scopeLabel(ctx.scope)}).`,
          mutated: false,
          summary: call.name,
        }
      }
      return {
        output: cards
          .map((card) => {
            const env = byUid.get(card.uid)
            const header = env
              ? `${env.from.join(', ')} | ${env.subject} | ${env.date} | folder:${card.folderId}`
              : `folder:${card.folderId}`
            const msgId = card.msgIdHeader ? ` msg-id:${card.msgIdHeader}` : ''
            const snippet = env ? ` :: "${env.snippet}"` : ''
            return `${citeToken(card.uid)} ${header}${msgId}${snippet}`
          })
          .join('\n'),
        mutated: false,
        summary: call.name,
      }
    }

    default:
      return fail(`Unknown tool: ${call.name}`, call.name)
  }
}
