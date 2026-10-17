/**
 * Scripted fixtures for sync-engine tests (Y lane). Fakes only: no
 * network, no real accounts, no credentials. A ScriptedDriver plays one
 * queued SyncFolderResult per fetchIncremental call and records sends,
 * intent attempts, and pause/resume transitions for assertions.
 */
import type { SyncDriverPort, SyncFolderResult, SyncStorePort } from '../src/main/sync/engine'
import type {
  IntentAttempt,
  MailEnvelope,
  MailFlags,
  MailIntent,
  MailMessageRef,
  SendRequest,
  SyncState,
} from '@genoffice/mail-protocols'
import type { AccountKind } from '../src/main/sync/caps'

let counter = 1

export interface TestEnvelopeOverrides extends Omit<Partial<MailEnvelope>, 'flags' | 'ref'> {
  flags?: Partial<MailFlags>
  ref?: Partial<MailMessageRef>
}

export function testEnvelope(uid: string, overrides?: TestEnvelopeOverrides): MailEnvelope {
  const flags: MailFlags = { seen: false, flagged: false, answered: false, draft: false }
  const { flags: flagOverrides, ref: refOverrides, ...rest } = overrides ?? {}
  return {
    from: ['sender@example.com'],
    to: ['reader@example.com'],
    cc: [],
    subject: `Subject ${uid}`,
    date: '2026-10-14T00:00:00.000Z',
    labels: [],
    snippet: `snippet ${uid}`,
    hasAttach: false,
    ...rest,
    ref: { uid, folderId: 'inbox', threadId: `thread-${uid}`, ...refOverrides },
    flags: { ...flags, ...flagOverrides },
  }
}

export function testSendRequest(overrides?: Partial<SendRequest>): SendRequest {
  return {
    to: ['reader@example.com'],
    subject: 'Hello',
    text: 'Body text',
    ...overrides,
  }
}

/**
 * Standalone MailStore fake with the scoped folder purge a resync
 * prefers. Mirrors the InMemoryMailStore shape (which keeps its maps
 * private) so resync tests can assert folder rows replaced while
 * outbox, drafts, and intents survive.
 */
export class FakeSyncStore implements SyncStorePort {
  private envelopes = new Map<string, MailEnvelope>()
  private drafts = new Map<string, string>()
  private outbox = 0
  private intents = new Map<string, MailIntent>()
  private sync = new Map<string, SyncState>()
  private nextId = 1

  upsertEnvelopes(envs: MailEnvelope[]): void {
    for (const env of envs) {
      this.envelopes.set(`${env.ref.folderId}/${env.ref.uid}`, env)
    }
  }

  envelopesIn(folderId: string): MailEnvelope[] {
    return [...this.envelopes.values()].filter((env) => env.ref.folderId === folderId)
  }

  purgeFolder(accountId: string, folderId: string): void {
    void accountId
    for (const [key, env] of this.envelopes) {
      if (env.ref.folderId === folderId) this.envelopes.delete(key)
    }
  }

  getThread(threadId: string): MailEnvelope[] {
    return [...this.envelopes.values()].filter((env) => env.ref.threadId === threadId)
  }

  enqueueSend(_req: SendRequest): string {
    this.outbox += 1
    return `store-outbox-${this.outbox}`
  }

  saveDraft(accountId: string, threadKey: string, text: string): void {
    this.drafts.set(`${accountId}/${threadKey}`, text)
  }

  draftFor(accountId: string, threadKey: string): string | undefined {
    return this.drafts.get(`${accountId}/${threadKey}`)
  }

  queueIntent(it: Omit<MailIntent, 'id' | 'state'>): string {
    const id = `intent-${this.nextId++}`
    this.intents.set(id, { ...it, id, state: 'queued' })
    return id
  }

  pendingIntents(accountId: string): MailIntent[] {
    return [...this.intents.values()].filter(
      (it) => it.accountId === accountId && it.state === 'queued',
    )
  }

  markIntent(id: string, state: 'acked' | 'failed'): void {
    const current = this.intents.get(id)
    if (current) this.intents.set(id, { ...current, state })
  }

  intentState(id: string): MailIntent['state'] | undefined {
    return this.intents.get(id)?.state
  }

  setSyncState(s: SyncState): void {
    this.sync.set(`${s.accountId}/${s.folderId}`, s)
  }

  getSyncState(accountId: string, folderId: string): SyncState | undefined {
    return this.sync.get(`${accountId}/${folderId}`)
  }

  fts(query: string, accountIds: string[]): MailEnvelope[] {
    const needle = query.trim().toLowerCase()
    if (!needle) return []
    void accountIds
    return [...this.envelopes.values()].filter((env) =>
      [env.subject, env.snippet, ...env.from, ...env.to].join(' ').toLowerCase().includes(needle),
    )
  }
}

export interface ScriptedDriverOptions {
  accountId?: string
  kind?: AccountKind
  folderId?: string
  results?: SyncFolderResult[]
  intentOutcomes?: IntentAttempt[]
  confirmSentIds?: Set<string>
  failSendWith?: string
}

export function syncState(accountId: string, overrides?: Partial<SyncState>): SyncState {
  return {
    accountId,
    folderId: 'inbox',
    lastSyncAt: '2026-10-14T00:00:00.000Z',
    ...overrides,
  }
}

export function incrementalResult(envelopes: MailEnvelope[], state: SyncState): SyncFolderResult {
  return { decision: { kind: 'incremental' }, envelopes, state }
}

export class ScriptedDriver implements SyncDriverPort {
  accountId: string
  kind: AccountKind
  folderId: string
  readonly sent: Array<{ req: SendRequest; outboxId: string }> = []
  readonly intentAttempts: MailIntent[] = []
  readonly refetched: string[] = []
  readonly confirmed: string[] = []
  paused = 0
  resumed = 0
  fetchCalls = 0
  private readonly results: SyncFolderResult[]
  private readonly intentOutcomes: IntentAttempt[]
  private readonly confirmSentIds: Set<string>
  failSendWith?: string

  constructor(options?: ScriptedDriverOptions) {
    this.accountId = options?.accountId ?? `account-${counter++}`
    this.kind = options?.kind ?? 'imap'
    this.folderId = options?.folderId ?? 'inbox'
    this.results = [...(options?.results ?? [])]
    this.intentOutcomes = [...(options?.intentOutcomes ?? [])]
    this.confirmSentIds = options?.confirmSentIds ?? new Set()
    this.failSendWith = options?.failSendWith
  }

  queueResult(result: SyncFolderResult): void {
    this.results.push(result)
  }

  clearResults(): void {
    this.results.length = 0
  }

  async fetchIncremental(stored: SyncState | undefined): Promise<SyncFolderResult> {
    void stored
    this.fetchCalls += 1
    const next = this.results.shift()
    if (!next) throw new Error('scripted driver has no queued result')
    return next
  }

  async send(req: SendRequest, outboxId: string): Promise<void> {
    if (this.failSendWith !== undefined) throw new Error(this.failSendWith)
    this.sent.push({ req, outboxId })
    this.confirmSentIds.add(outboxId)
  }

  async confirmSent(outboxId: string): Promise<boolean> {
    this.confirmed.push(outboxId)
    return this.confirmSentIds.has(outboxId)
  }

  async attemptIntent(intent: MailIntent): Promise<IntentAttempt> {
    this.intentAttempts.push(intent)
    return this.intentOutcomes.shift() ?? 'acked'
  }

  async refetch(targetUid: string): Promise<void> {
    this.refetched.push(targetUid)
  }

  pause(): void {
    this.paused += 1
  }

  resume(): void {
    this.resumed += 1
  }
}
