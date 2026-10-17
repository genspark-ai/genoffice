import { DEFAULT_AI_PANEL_PREFS } from '@genoffice/ui'
// Browser-safe subpath (package root drags node-only drivers into the
// renderer bundle; drivers stay main-only per the app-shell spec).
import { matchesQuery } from '@genoffice/mail-protocols/query'
import type {
  MailBody,
  MailEnvelope,
  MailEvent,
  MailFlags,
  MailFolder,
  MailMessageRef,
  MailQuery,
  SendRequest,
} from '@genoffice/mail-protocols'
import type {
  MailAccountSummary,
  MailApi,
  MessagePage,
  OfflineStatus,
  OutboxRecord,
  OutboxStatus,
} from '../shared/ipc'

/**
 * Renderer-side IPC fakes for Lane R UI dev and vitest (Lane R owns).
 * Implements the read/query surface of `MailApi` over seeded in-memory
 * envelopes; writes mutate local maps only. Fake addresses use
 * example.com / example.net only — never real mailboxes. The real sync
 * engine and main-process handlers are a later phase; until they land,
 * `resolveMailApi()` falls back to these fakes when no preload bridge
 * is present (plain vite dev).
 */

export interface FakeSeedMessage {
  uid: string
  folderId: string
  from: string[]
  to: string[]
  cc?: string[]
  subject: string
  date: string
  seen?: boolean
  flagged?: boolean
  snippet: string
  hasAttach?: boolean
  text: string
  html?: string
  replyTo?: string
  auth?: 'pass' | 'fail' | 'unknown'
}

export interface FakeSeed {
  accountId: string
  label: string
  folders: MailFolder[]
  messages: FakeSeedMessage[]
  /** Extra server-only hits returned by searchServer (local store lacks them). */
  serverExtra?: FakeSeedMessage[]
}

function toEnvelope(seed: FakeSeedMessage): MailEnvelope {
  return {
    ref: { uid: seed.uid, folderId: seed.folderId, threadId: `thread-${seed.uid}` },
    from: seed.from,
    to: seed.to,
    cc: seed.cc ?? [],
    subject: seed.subject,
    date: seed.date,
    flags: {
      seen: seed.seen ?? false,
      flagged: seed.flagged ?? false,
      answered: false,
      draft: false,
    },
    labels: [],
    snippet: seed.snippet,
    hasAttach: seed.hasAttach ?? false,
  }
}

function toBody(seed: FakeSeedMessage): FakeMailBody {
  return { text: seed.text, htmlRef: undefined, html: seed.html }
}

/** Fake-only body extension: P3 fakes inline HTML (never on the wire). */
export interface FakeMailBody extends MailBody {
  html?: string
}

export function defaultSeed(): FakeSeed {
  return {
    accountId: 'demo',
    label: 'Demo Account',
    folders: [
      { id: 'inbox', name: 'Inbox', role: 'inbox', unread: 3 },
      { id: 'sent', name: 'Sent', role: 'sent', unread: 0 },
      { id: 'drafts', name: 'Drafts', role: 'drafts', unread: 0 },
      { id: 'archive', name: 'Archive', role: 'archive', unread: 0 },
    ],
    messages: [
      {
        uid: 'm1',
        folderId: 'inbox',
        from: ['Acme Billing <billing@example.com>'],
        to: ['Reader <reader@example.net>'],
        subject: 'Your invoice is ready',
        date: '2026-10-10T09:00:00.000Z',
        snippet: 'Invoice INV-1042 for October is attached for review.',
        hasAttach: true,
        text: 'Hello,\n\nInvoice INV-1042 for October is attached for review.\n\nAcme Billing',
        html: '<p>Hello,</p><p>Invoice <b>INV-1042</b> for October is attached for review.</p><p style="color: #107c41;">Thank you for your business.</p>',
      },
      {
        uid: 'm2',
        folderId: 'inbox',
        from: ['Team Chat <chat@example.com>'],
        to: ['Reader <reader@example.net>'],
        subject: 'Weekly notes with a remote banner',
        date: '2026-10-11T12:30:00.000Z',
        seen: true,
        snippet: 'This week shipped the new reader with remote images blocked.',
        text: 'This week shipped the new reader with remote images blocked.',
        html: '<p>This week shipped the new reader.</p><img src="https://tracker.example.com/pixel.png" alt="Weekly banner"><p style="background-color: #fff4ce;">Remote images stay blocked until you allow them.</p>',
      },
      {
        uid: 'm3',
        folderId: 'inbox',
        from: ['Support <support@example.com>'],
        to: ['Reader <reader@example.net>'],
        subject: 'Reply-To differs on this one',
        date: '2026-10-12T08:15:00.000Z',
        snippet: 'A support note whose Reply-To points elsewhere.',
        text: 'A support note whose Reply-To points elsewhere.',
        html: '<p>A support note whose Reply-To points elsewhere.</p>',
        replyTo: 'outside@example.org',
      },
      {
        uid: 'm4',
        folderId: 'inbox',
        from: ['Acme <mail@xn--acme-pay-example.com>'],
        to: ['Reader <reader@example.net>'],
        subject: 'Lookalike sender sample',
        date: '2026-10-12T18:40:00.000Z',
        snippet: 'Sent from a localised domain that mimics a known sender.',
        text: 'Sent from a localised domain that mimics a known sender.',
        html: '<p>Sent from a localised domain that mimics a known sender.</p>',
      },
      {
        uid: 'm5',
        folderId: 'inbox',
        from: ['Flags Demo <flags@example.com>'],
        to: ['Reader <reader@example.net>'],
        subject: 'Flagged and unread sample',
        date: '2026-10-13T07:05:00.000Z',
        flagged: true,
        snippet: 'Starred so flag filters have something to match.',
        text: 'Starred so flag filters have something to match.',
      },
      {
        uid: 'm6',
        folderId: 'sent',
        from: ['Reader <reader@example.net>'],
        to: ['Acme Billing <billing@example.com>'],
        subject: 'Re: Your invoice is ready',
        date: '2026-10-13T09:20:00.000Z',
        seen: true,
        snippet: 'Thanks — paid this morning, receipt appreciated.',
        text: 'Thanks — paid this morning, receipt appreciated.',
      },
    ],
    serverExtra: [
      {
        uid: 'srv1',
        folderId: 'inbox',
        from: ['Archive Bot <archive@example.com>'],
        to: ['Reader <reader@example.net>'],
        subject: 'Server-only invoice copy',
        date: '2026-09-01T10:00:00.000Z',
        seen: true,
        snippet: 'An older invoice copy that lives on the server only.',
        text: 'An older invoice copy that lives on the server only.',
      },
    ],
  }
}

/** Local post-filter used by the search box: IPC fakes first, matchesQuery after. */
export function applyLocalSearch(envelopes: MailEnvelope[], text: string): MailEnvelope[] {
  const needle = text.trim()
  if (!needle) return envelopes
  return envelopes.filter((env) => matchesQuery(env, { text: needle }))
}

function noopUnsubscribe(): () => void {
  return () => undefined
}

// ── P5 AI retrieval shapes (spec sections 3-4; UI lane owns this fake) ──
// Mirrors the W-lane `mail:ai-retrieve` contract exactly so the K lane can
// consume this fake structurally (K `MailAiBridge`) until the real bridge
// lands; production responses carry the same fields.

/** Picker-selected account set; travels with EVERY AI call (1-8 ids). */
export interface AiScope {
  accountIds: string[]
}

export type AiRetrieveOp = 'triage-headers' | 'fetch-bodies' | 'cite-resolve'

export interface AiRetrieveRequest {
  scope: AiScope
  op: AiRetrieveOp
  /** triage-headers: header/snippet search text (required for that op). */
  query?: string
  folderId?: string
  /** Default 50 (FTS_DEFAULT_LIMIT). */
  limit?: number
  /** fetch-bodies / cite-resolve: explicit user-cited set (required). */
  refs?: MailMessageRef[]
}

export interface AiRetrieveCitation {
  uid: string
  folderId: string
  msgIdHeader?: string
  /** Scope account holding the ref (mirrors the production bridge tagging). */
  accountId: string
}

export interface AiRetrieveResponse {
  /** Echo — the caller discards the result on mismatch (stale scope). */
  scope: AiScope
  /** Headers+snippet ONLY; bodies NEVER appear here. */
  envelopes: MailEnvelope[]
  /** fetch-bodies only, cited set only. */
  bodies?: { ref: MailMessageRef; accountId: string; text: string }[]
  citations: AiRetrieveCitation[]
}

function envKey(ref: MailMessageRef): string {
  return `${ref.folderId}/${ref.uid}`
}

function envelopeHaystack(env: MailEnvelope): string {
  return [...env.from, ...env.to, env.subject, env.snippet].join('\n')
}

function toFakeCitation(env: MailEnvelope, accountId: string): AiRetrieveCitation {
  return {
    uid: env.ref.uid,
    folderId: env.ref.folderId,
    msgIdHeader: `<${env.ref.uid}@example.com>`,
    accountId,
  }
}

function checkFakeScope(scope: AiScope, known: string[]): string | null {
  if (!scope || !Array.isArray(scope.accountIds) || scope.accountIds.length < 1) {
    return 'mail: AI scope must select at least one account'
  }
  if (scope.accountIds.length > 8) return 'mail: AI scope selects too many accounts (max 8)'
  for (const id of scope.accountIds) {
    if (typeof id !== 'string' || id.length < 1 || !known.includes(id)) {
      return 'mail: unknown account in AI scope'
    }
  }
  return null
}

function clampFakeLimit(raw: number | undefined): number {
  if (typeof raw !== 'number' || !Number.isFinite(raw)) return 50
  return Math.min(Math.max(Math.floor(raw), 1), 200)
}

export class FakeMailApi implements MailApi {
  private accounts: MailAccountSummary[]
  private folders: Map<string, MailFolder[]>
  private envelopes: MailEnvelope[]
  private bodies: Map<string, MailBody>
  private security: Map<string, { replyTo?: string; auth?: 'pass' | 'fail' | 'unknown' }>
  /** Envelope owner per folder/uid key (outbox + AI fakes filter by account). */
  private readonly accountOf = new Map<string, string>()
  private outboxRecords: OutboxRecord[] = []
  private outboxSeq = 0
  private drafts: Map<string, string> = new Map()
  private online = true
  private offlineDetail?: string
  private readonly offlineHandlers = new Set<(status: OfflineStatus) => void>()
  private readonly syncHandlers = new Set<(event: MailEvent) => void>()

  constructor(seeds: FakeSeed[] = [defaultSeed()]) {
    this.accounts = seeds.map((seed) => ({
      accountId: seed.accountId,
      label: seed.label,
      sessionOnly: true,
    }))
    this.folders = new Map(seeds.map((seed) => [seed.accountId, [...seed.folders]]))
    this.envelopes = seeds.flatMap((seed) => seed.messages.map(toEnvelope))
    for (const seed of seeds) {
      for (const message of seed.messages) {
        this.accountOf.set(`${message.folderId}/${message.uid}`, seed.accountId)
      }
    }
    this.bodies = new Map(
      seeds.flatMap((seed) =>
        seed.messages.map(
          (message) => [`${message.folderId}/${message.uid}`, toBody(message)] as const,
        ),
      ),
    )
    this.security = new Map(
      seeds.flatMap((seed) =>
        seed.messages
          .filter((message) => message.replyTo !== undefined || message.auth !== undefined)
          .map(
            (message) =>
              [
                `${message.folderId}/${message.uid}`,
                { replyTo: message.replyTo, auth: message.auth },
              ] as const,
          ),
      ),
    )
    for (const seed of seeds) {
      for (const extra of seed.serverExtra ?? []) {
        this.bodies.set(`${extra.folderId}/${extra.uid}`, toBody(extra))
      }
    }
    this.serverExtra = seeds.flatMap((seed) => (seed.serverExtra ?? []).map(toEnvelope))
  }

  private serverExtra: MailEnvelope[]

  /** Phishing/auth metadata the envelope shape cannot carry (fake-only sidecar). */
  fakeSecurity(ref: MailMessageRef): { replyTo?: string; auth?: 'pass' | 'fail' | 'unknown' } {
    return this.security.get(`${ref.folderId}/${ref.uid}`) ?? {}
  }

  async openMessage(_mailRef: string): Promise<string | null> {
    return null
  }

  async listAccounts(): Promise<MailAccountSummary[]> {
    return [...this.accounts]
  }

  async addAccount(label: string): Promise<MailAccountSummary> {
    const account: MailAccountSummary = {
      accountId: `fake-${this.accounts.length + 1}`,
      label,
      sessionOnly: true,
    }
    this.accounts.push(account)
    this.folders.set(account.accountId, [])
    return account
  }

  async removeAccount(accountId: string): Promise<void> {
    this.accounts = this.accounts.filter((account) => account.accountId !== accountId)
    this.folders.delete(accountId)
  }

  async listFolders(accountId: string): Promise<MailFolder[]> {
    return [...(this.folders.get(accountId) ?? [])]
  }

  async listMessages(
    accountId: string,
    folderId: string,
    limit: number,
    cursor?: string,
  ): Promise<MessagePage> {
    void accountId
    const start = cursor === undefined ? 0 : Number.parseInt(cursor, 10)
    const inFolder = this.envelopes
      .filter((env) => env.ref.folderId === folderId)
      .sort((a, b) => (a.date < b.date ? 1 : -1))
    const page = inFolder.slice(start, start + limit)
    const next = start + limit < inFolder.length ? String(start + limit) : undefined
    return { envelopes: page, cursor: next }
  }

  async fetchBody(ref: MailMessageRef): Promise<MailBody> {
    return this.bodies.get(`${ref.folderId}/${ref.uid}`) ?? { text: '' }
  }

  async searchLocal(accountIds: string[], query: MailQuery): Promise<MailEnvelope[]> {
    void accountIds
    return this.envelopes.filter((env) => matchesQuery(env, query))
  }

  async searchServer(
    accountId: string,
    folderId: string,
    query: MailQuery,
  ): Promise<MailEnvelope[]> {
    void accountId
    void folderId
    const local = this.envelopes.filter((env) => matchesQuery(env, query))
    const remote = this.serverExtra.filter((env) => matchesQuery(env, query))
    return [...local, ...remote]
  }

  async composeSend(req: SendRequest): Promise<{ outboxId: string }> {
    this.outboxSeq += 1
    const id = `fake-outbox-${this.outboxSeq}`
    const owner = this.accounts[0]?.accountId ?? 'demo'
    this.outboxRecords.push({
      id,
      accountId: owner,
      req,
      state: 'queued',
      attempts: 0,
      queuedAt: new Date().toISOString(),
    })
    return { outboxId: id }
  }

  async saveDraft(accountId: string, threadKey: string, text: string): Promise<void> {
    this.drafts.set(`${accountId}/${threadKey}`, text)
  }

  async outboxStatus(accountId: string): Promise<OutboxStatus> {
    const mine = this.outboxRecords.filter((item) => item.accountId === accountId)
    return {
      queued: mine.filter((item) => item.state === 'queued' || item.state === 'sending').length,
      failed: mine.filter((item) => item.state === 'failed').length,
    }
  }

  async move(ref: MailMessageRef, toFolderId: string): Promise<void> {
    const env = this.envelopes.find(
      (item) => item.ref.uid === ref.uid && item.ref.folderId === ref.folderId,
    )
    if (env) env.ref.folderId = toFolderId
  }

  async setFlags(ref: MailMessageRef, flags: Partial<MailFlags>): Promise<void> {
    const env = this.envelopes.find(
      (item) => item.ref.uid === ref.uid && item.ref.folderId === ref.folderId,
    )
    if (env) env.flags = { ...env.flags, ...flags }
  }

  async deleteMessage(ref: MailMessageRef): Promise<void> {
    this.envelopes = this.envelopes.filter(
      (item) => !(item.ref.uid === ref.uid && item.ref.folderId === ref.folderId),
    )
  }

  onSyncEvent(handler: (event: MailEvent) => void): () => void {
    this.syncHandlers.add(handler)
    return () => {
      this.syncHandlers.delete(handler)
    }
  }

  /** Emit a sync event to subscribers (test/dev driver). */
  emitSync(event: MailEvent): void {
    for (const handler of this.syncHandlers) handler(event)
  }

  // P4 sync/outbox additions (Y owns the contract; fake stubs for dev).
  async syncNow(accountId: string, folderId?: string): Promise<{ started: boolean }> {
    if (!this.online) {
      this.emitSync({
        kind: 'sync-state',
        accountId,
        folderId: folderId ?? 'inbox',
        state: 'error',
        detail: 'offline',
      })
      return { started: false }
    }
    this.emitSync({
      kind: 'sync-state',
      accountId,
      folderId: folderId ?? 'inbox',
      state: 'syncing',
    })
    // Settle back to idle on a microtask so a panel busy flag cannot stick.
    void Promise.resolve().then(() => {
      this.emitSync({
        kind: 'sync-state',
        accountId,
        folderId: folderId ?? 'inbox',
        state: 'idle',
      })
    })
    return { started: true }
  }

  async outboxList(accountId: string): Promise<OutboxRecord[]> {
    return this.outboxRecords
      .filter((item) => item.accountId === accountId)
      .sort((a, b) =>
        a.queuedAt === b.queuedAt ? (a.id < b.id ? 1 : -1) : a.queuedAt < b.queuedAt ? 1 : -1,
      )
  }

  async outboxRetry(
    accountId: string,
    outboxId: string,
  ): Promise<{ outboxId: string; state: string }> {
    const record = this.outboxRecords.find(
      (item) => item.id === outboxId && item.accountId === accountId,
    )
    if (record && (record.state === 'failed' || record.state === 'cancelled')) {
      record.state = 'queued'
      record.attempts = 0
      record.lastError = undefined
    }
    return { outboxId, state: 'queued' }
  }

  async outboxCancel(accountId: string, outboxId: string): Promise<{ cancelled: boolean }> {
    const record = this.outboxRecords.find(
      (item) => item.id === outboxId && item.accountId === accountId,
    )
    if (record && (record.state === 'queued' || record.state === 'failed')) {
      record.state = 'cancelled'
      return { cancelled: true }
    }
    return { cancelled: false }
  }

  onOfflineStatus(handler: (status: OfflineStatus) => void): () => void {
    this.offlineHandlers.add(handler)
    // Replay the current value on subscribe (spec section 6).
    handler({ online: this.online, detail: this.offlineDetail })
    return () => {
      this.offlineHandlers.delete(handler)
    }
  }

  /** Flip connectivity and notify subscribers (test/dev driver). */
  setOnline(online: boolean, detail?: string): void {
    this.online = online
    this.offlineDetail = detail
    const status: OfflineStatus = { online, detail }
    for (const handler of this.offlineHandlers) handler(status)
  }

  /** Seed demo outbox rows (queued/sending/failed) for shell probing. */
  seedOutboxDemo(accountId = 'demo'): void {
    const seeds: Array<{
      id: string
      subject: string
      to: string[]
      state: OutboxRecord['state']
      lastError?: string
    }> = [
      { id: 'demo-o1', subject: 'Queued note', to: ['a@example.com'], state: 'queued' },
      { id: 'demo-o2', subject: 'Sending note', to: ['b@example.com'], state: 'sending' },
      {
        id: 'demo-o3',
        subject: 'Failed note',
        to: ['c@example.com'],
        state: 'failed',
        lastError: 'smtp: connection refused',
      },
    ]
    for (const [index, seed] of seeds.entries()) {
      if (this.outboxRecords.some((item) => item.id === seed.id)) continue
      this.outboxRecords.push({
        id: seed.id,
        accountId,
        req: { to: seed.to, subject: seed.subject, text: '' },
        state: seed.state,
        attempts: seed.state === 'failed' ? 2 : 0,
        lastError: seed.lastError,
        queuedAt: `2026-10-15T04:00:${String(index).padStart(2, '0')}Z`,
      })
    }
  }

  // ── P5 AI retrieval fake (UI lane owns; K lane consumes structurally) ──
  // Same spec section 4 shape the W lane bridges in production
  // (`mail:ai-retrieve`): headers+snippet envelopes only, bodies for the
  // cited set only, scope echoed back. Unknown accounts are rejected.

  async aiRetrieve(request: AiRetrieveRequest): Promise<AiRetrieveResponse> {
    const scopeError = checkFakeScope(
      request.scope,
      this.accounts.map((a) => a.accountId),
    )
    if (scopeError) throw new Error(scopeError)
    const scope: AiScope = { accountIds: [...request.scope.accountIds] }
    const limit = clampFakeLimit(request.limit)
    if (request.op === 'triage-headers') {
      const query = request.query?.trim() ?? ''
      if (!query) throw new Error('mail: triage-headers requires a query')
      const needle = query.toLowerCase()
      const envelopes = this.envelopes
        .filter((env) => scope.accountIds.includes(this.accountOf.get(envKey(env.ref)) ?? ''))
        .filter((env) => envelopeHaystack(env).toLowerCase().includes(needle))
        .slice(0, limit)
      const cite = (env: MailEnvelope): AiRetrieveCitation =>
        toFakeCitation(env, this.accountOf.get(envKey(env.ref)) ?? '')
      return { scope, envelopes, citations: envelopes.map(cite) }
    }
    const refs = request.refs ?? []
    if (refs.length === 0) throw new Error(`mail: ${request.op} requires refs`)
    for (const ref of refs) {
      const owner = this.accountOf.get(envKey(ref))
      if (!owner || !scope.accountIds.includes(owner)) {
        throw new Error('mail: ref outside the AI scope')
      }
    }
    const envelopes = refs
      .map((ref) =>
        this.envelopes.find((env) => env.ref.uid === ref.uid && env.ref.folderId === ref.folderId),
      )
      .filter((env): env is MailEnvelope => env !== undefined)
    if (request.op === 'cite-resolve') {
      const cite = (env: MailEnvelope): AiRetrieveCitation =>
        toFakeCitation(env, this.accountOf.get(envKey(env.ref)) ?? '')
      return { scope, envelopes, citations: envelopes.map(cite) }
    }
    const bodies = refs
      .map((ref) => {
        const body = this.bodies.get(envKey(ref))
        if (!body) return null
        return { ref, accountId: this.accountOf.get(envKey(ref)) ?? '', text: body.text }
      })
      .filter(
        (entry): entry is { ref: MailMessageRef; accountId: string; text: string } =>
          entry !== null,
      )
    const cite = (env: MailEnvelope): AiRetrieveCitation =>
      toFakeCitation(env, this.accountOf.get(envKey(env.ref)) ?? '')
    return { scope, envelopes, bodies, citations: envelopes.map(cite) }
  }

  async attachmentFetch(): Promise<{ ok: boolean; path?: string }> {
    return { ok: false }
  }

  onAttachmentProgress(): () => void {
    return noopUnsubscribe()
  }

  setDirty(): void {
    // Fake ignores the dirty push; the guard path is covered by shell tests.
  }

  onCloseSaveRequest(): () => void {
    return noopUnsubscribe()
  }

  sendCloseSaveResult(): void {
    // No-op in the fake; nothing is awaiting a close verdict.
  }

  async getLanguage(): Promise<'zh'> {
    return 'zh'
  }

  onLanguageChanged(): () => void {
    return noopUnsubscribe()
  }

  async getTheme(): Promise<'system'> {
    return 'system'
  }

  onThemeChanged(): () => void {
    return noopUnsubscribe()
  }

  async getAiPanelPrefs(): Promise<typeof DEFAULT_AI_PANEL_PREFS> {
    return DEFAULT_AI_PANEL_PREFS
  }

  onAiPanelPrefsChanged(): () => void {
    return noopUnsubscribe()
  }

  onChromePressed(): () => void {
    return noopUnsubscribe()
  }

  async getAiSettings(): Promise<never> {
    throw new Error('mail: AI settings are stubbed in P3')
  }

  async aiGskStatus(): Promise<never> {
    throw new Error('mail: AI status is stubbed in P3')
  }

  async aiStream(): Promise<never> {
    throw new Error('mail: AI streaming is stubbed in P3')
  }

  async aiStreamCancel(): Promise<void> {
    // No stream exists in the fake; cancelling is a no-op.
  }

  onAiStream(): () => void {
    return noopUnsubscribe()
  }
}

let sharedFake: FakeMailApi | null = null

export function sharedFakeApi(): FakeMailApi {
  sharedFake ??= new FakeMailApi()
  return sharedFake
}

/** Preload bridge when present, seeded fakes otherwise (plain vite dev). */
export function resolveMailApi(): MailApi {
  if (typeof window !== 'undefined' && window.mailApi) return window.mailApi
  return sharedFakeApi()
}
