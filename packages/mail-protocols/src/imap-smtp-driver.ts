/**
 * ImapSmtpDriver: MailProvider over IMAP (fetch/sync) + SMTP (send).
 *
 * Pure TypeScript: no Electron, no safeStorage, no direct sockets. The
 * main-process caller (apps/mail/src/main, per locked decision D6) injects
 * the connection objects and the credentials it unwrapped from the vault
 * (credential-vault.ts, the single safeStorage caller per R2). This driver
 * NEVER imports safeStorage or shell.openExternal; the security-eyes lane
 * (Alya) verifies that with a grep before merge.
 *
 * Sync mechanics (amended spec 3.2/4.3):
 * - IDLE where the server advertises it, else 30-60 s poll with backoff.
 * - CONDSTORE/MODSEQ where advertised, else UID-range fetch.
 * - UIDVALIDITY change means full folder re-fetch. Resync never deletes
 *   queued outbox sends, drafts, or intents: the driver owns none of those
 *   tables (they live in MailStore/sync-engine, P3 lane), so there is
 *   nothing here that could delete them; the resync fixture pins that.
 * - Offline moves/flags/deletes queue as idempotent intents
 *   (targetUid + expectedModseq). Stale modseq re-fetches then retries
 *   once, else surfaces a conflict; never a silent merge guess.
 */
import {
  MailBody,
  MailEnvelope,
  MailEvent,
  MailFlags,
  MailFolder,
  MailIntent,
  MailMessageRef,
  MailProvider,
  MailQuery,
  SendRequest,
  StagedAttachment,
  SyncState,
} from './types.js'
import { OUTGOING_WARN_BYTES } from './attachments.js'

export const DEFAULT_POLL_INTERVAL_MS = 45_000
export const MAX_POLL_BACKOFF_MS = 300_000
/** Hard reject above this total size unless the caller overrides the caps. */
export const MAX_TOTAL_BYTES = 50 * 1024 * 1024
export const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024
export const MAX_RECIPIENTS = 100

export type ImapSmtpErrorCode =
  | 'offline'
  | 'oversize'
  | 'too-many-recipients'
  | 'attachment-too-large'
  | 'conflict'
  | 'network'
  | 'auth'

export class ImapSmtpError extends Error {
  readonly code: ImapSmtpErrorCode
  constructor(code: ImapSmtpErrorCode, message: string, options?: { cause?: unknown }) {
    super(message)
    this.name = 'ImapSmtpError'
    this.code = code
    if (options?.cause !== undefined) {
      ;(this as { cause?: unknown }).cause = options.cause
    }
  }
}

/** Credentials unwrapped by the main-process vault caller. Opaque here. */
export interface InjectedCredentials {
  user: string
  password: string
}

export interface ImapConnectionOptions {
  host: string
  port: number
  secure: boolean
  auth: InjectedCredentials
}

export interface SmtpConnectionOptions {
  host: string
  port: number
  secure: boolean
  auth: InjectedCredentials
  /** RFC 5322 From header used for outgoing mail. */
  from: string
}

/** Per-driver send limits, enforced in the driver and surfaced to compose. */
export interface ImapSmtpSendCaps {
  /** Above this total, send proceeds with a warning. Default 25 MiB. */
  warnTotalBytes?: number
  /** Above this total, send is rejected. Default 50 MiB. */
  maxTotalBytes?: number
  /** Above this per-file size, send is rejected. Default 25 MiB. */
  maxAttachmentBytes?: number
  /** Above this recipient count, send is rejected. Default 100. */
  maxRecipients?: number
}

export interface ImapSmtpAccountConfig {
  accountId: string
  imap: ImapConnectionOptions
  smtp: SmtpConnectionOptions
  /** Poll cadence when the server does not advertise IDLE. Default 45 s. */
  pollIntervalMs?: number
  /** Backoff ceiling for failed polls. Default 5 min. */
  maxPollBackoffMs?: number
  sendCaps?: ImapSmtpSendCaps
}

export interface ImapMailboxInfo {
  path: string
  name: string
  role: MailFolder['role']
  unread: number
}

export interface ImapMailboxStatus {
  uidValidity: number
  highestModseq?: string
  exists: number
}

/** One fetched row, before mapping to the shared envelope shape. */
export interface ImapEnvelopeData {
  uid: string
  modseq?: string
  from: string[]
  to: string[]
  cc: string[]
  subject: string
  date: string
  seen: boolean
  flagged: boolean
  answered: boolean
  draft: boolean
  msgIdHeader?: string
  labels?: string[]
  snippet: string
  hasAttach: boolean
}

/**
 * Narrow seam around an IMAP connection. The production implementation
 * wraps imapflow (see imap-transport.ts); tests inject FakeImapServer.
 */
export interface ImapClientLike {
  connect(): Promise<void>
  close(): Promise<void>
  supports(capability: string): boolean
  listMailboxes(): Promise<ImapMailboxInfo[]>
  selectMailbox(path: string): Promise<ImapMailboxStatus>
  fetchRange(path: string, afterUid: string | undefined, limit: number): Promise<ImapEnvelopeData[]>
  fetchSinceModseq(path: string, modseq: string): Promise<ImapEnvelopeData[]>
  fetchBody(path: string, uid: string): Promise<{ text: string }>
  /** Returns the new MODSEQ when the server reports one. */
  setFlags(
    path: string,
    uid: string,
    flags: Partial<MailFlags>,
    options?: { expectedModseq?: string },
  ): Promise<string | undefined>
  moveMessage(
    path: string,
    uid: string,
    destPath: string,
    options?: { expectedModseq?: string },
  ): Promise<void>
  deleteMessage(path: string, uid: string): Promise<void>
  searchMessages(path: string, query: MailQuery): Promise<ImapEnvelopeData[]>
  idle(onEvent: (ev: IdleMailboxEvent) => void): { stop(): void }
}

export interface IdleMailboxEvent {
  path: string
  uid: string
  kind: 'new' | 'updated' | 'deleted'
}

export interface SmtpAttachmentPayload {
  filename: string
  contentType: string
  /** Streamed from disk at send time; preferred for large files. */
  path?: string
  content?: Uint8Array
}

export interface SmtpSendMail {
  from: string
  to: string[]
  cc?: string[]
  bcc?: string[]
  subject: string
  text: string
  html?: string
  attachments: SmtpAttachmentPayload[]
}

/** Narrow seam around an SMTP transport (nodemailer in production). */
export interface SmtpTransportLike {
  send(mail: SmtpSendMail): Promise<{ messageId: string }>
  close(): Promise<void>
}

/** Minimal stores, implemented by the sync engine (P3 lane). In-memory defaults ship here. */
export interface SyncStateStore {
  getSyncState(accountId: string, folderId: string): SyncState | undefined
  setSyncState(state: SyncState): void
}

export interface IntentStore {
  queueIntent(intent: Omit<MailIntent, 'id' | 'state'>): string
  pendingIntents(accountId: string): MailIntent[]
  markIntent(id: string, state: 'acked' | 'failed'): void
}

export interface OutboxStore {
  enqueueSend(accountId: string, req: SendRequest): string
  markOutbox?(id: string, state: 'sent' | 'failed', detail?: string): void
}

export interface ImapSmtpDriverDeps {
  createImapClient: (options: ImapConnectionOptions) => ImapClientLike
  createSmtpTransport: (options: SmtpConnectionOptions) => SmtpTransportLike
  syncStore: SyncStateStore
  intentStore: IntentStore
  outbox: OutboxStore
  /** Optional event tap used by watch() and by list() resync detection. */
  onEvent?: (ev: MailEvent) => void
}

export interface SendResult {
  outboxId: string
  /** Present when the message actually left via SMTP (online path). */
  messageId?: string
  /** Non-blocking notices, e.g. the >25 MB outgoing warning. */
  warnings: string[]
}

export interface DrainResult {
  acked: string[]
  failed: Array<{ id: string; error: string }>
}

export class InMemorySyncStateStore implements SyncStateStore {
  private readonly states = new Map<string, SyncState>()
  getSyncState(accountId: string, folderId: string): SyncState | undefined {
    return this.states.get(`${accountId}/${folderId}`)
  }
  setSyncState(state: SyncState): void {
    this.states.set(`${state.accountId}/${state.folderId}`, state)
  }
}

export class InMemoryIntentStore implements IntentStore {
  private seq = 0
  readonly intents = new Map<string, MailIntent>()
  queueIntent(intent: Omit<MailIntent, 'id' | 'state'>): string {
    const id = `intent-${++this.seq}`
    this.intents.set(id, { ...intent, id, state: 'queued' })
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
}

export class InMemoryOutboxStore implements OutboxStore {
  private seq = 0
  readonly entries = new Map<string, { accountId: string; req: SendRequest; state: string }>()
  enqueueSend(accountId: string, req: SendRequest): string {
    const id = `outbox-${++this.seq}`
    this.entries.set(id, { accountId, req, state: 'queued' })
    return id
  }
  markOutbox(id: string, state: 'sent' | 'failed', detail?: string): void {
    const current = this.entries.get(id)
    if (current) this.entries.set(id, { ...current, state, ...(detail ? { detail } : {}) })
  }
}

function utf8Bytes(value: string): number {
  return Buffer.byteLength(value, 'utf8')
}

export function sendPayloadBytes(req: SendRequest): number {
  let total = utf8Bytes(req.text) + utf8Bytes(req.html ?? '')
  for (const att of req.attachments ?? []) total += att.size
  return total
}

function recipientCount(req: SendRequest): number {
  return req.to.length + (req.cc?.length ?? 0) + (req.bcc?.length ?? 0)
}

function threadIdFor(folderId: string, row: ImapEnvelopeData): string {
  return row.msgIdHeader ?? `${folderId}:${row.uid}`
}

/** Canonical event ref for a uid-only server signal (no envelope fetched). */
function eventRef(folderId: string, uid: string): MailMessageRef {
  return { uid, folderId, threadId: `${folderId}:${uid}` }
}

export function toEnvelope(folderId: string, row: ImapEnvelopeData): MailEnvelope {
  return {
    ref: {
      uid: row.uid,
      folderId,
      msgIdHeader: row.msgIdHeader,
      threadId: threadIdFor(folderId, row),
    },
    from: row.from,
    to: row.to,
    cc: row.cc,
    subject: row.subject,
    date: row.date,
    flags: { seen: row.seen, flagged: row.flagged, answered: row.answered, draft: row.draft },
    labels: row.labels ?? [],
    snippet: row.snippet,
    hasAttach: row.hasAttach,
  }
}

interface QueuedIntentPayload {
  folderId: string
  flags?: Partial<MailFlags>
  toFolderId?: string
}

function parseIntentPayload(intent: MailIntent): QueuedIntentPayload {
  return JSON.parse(intent.payloadJson) as QueuedIntentPayload
}

export class ImapSmtpDriver implements MailProvider {
  private readonly config: ImapSmtpAccountConfig
  private readonly deps: ImapSmtpDriverDeps
  private imap: ImapClientLike | undefined
  private smtp: SmtpTransportLike | undefined
  private connected = false
  private readonly pollWatermarks = new Map<string, string>()

  constructor(config: ImapSmtpAccountConfig, deps: ImapSmtpDriverDeps) {
    this.config = config
    this.deps = deps
  }

  get isConnected(): boolean {
    return this.connected
  }

  async connect(): Promise<void> {
    if (this.connected) return
    this.imap = this.deps.createImapClient(this.config.imap)
    this.smtp = this.deps.createSmtpTransport(this.config.smtp)
    await this.imap.connect()
    this.connected = true
  }

  async disconnect(): Promise<void> {
    this.connected = false
    const imap = this.imap
    const smtp = this.smtp
    this.imap = undefined
    this.smtp = undefined
    if (imap) await imap.close()
    if (smtp) await smtp.close()
  }

  private requireImap(): ImapClientLike {
    if (!this.connected || !this.imap) {
      throw new ImapSmtpError('offline', 'IMAP client is not connected')
    }
    return this.imap
  }

  private emit(ev: MailEvent): void {
    this.deps.onEvent?.(ev)
  }

  private caps(): Required<ImapSmtpSendCaps> {
    return {
      warnTotalBytes: this.config.sendCaps?.warnTotalBytes ?? OUTGOING_WARN_BYTES,
      maxTotalBytes: this.config.sendCaps?.maxTotalBytes ?? MAX_TOTAL_BYTES,
      maxAttachmentBytes: this.config.sendCaps?.maxAttachmentBytes ?? MAX_ATTACHMENT_BYTES,
      maxRecipients: this.config.sendCaps?.maxRecipients ?? MAX_RECIPIENTS,
    }
  }

  /** Validates send caps. Returns warnings; throws on hard violations. */
  checkSendCaps(req: SendRequest): string[] {
    const caps = this.caps()
    const warnings: string[] = []
    const recipients = recipientCount(req)
    if (recipients === 0) {
      throw new ImapSmtpError('too-many-recipients', 'Cannot send mail with no recipients')
    }
    if (recipients > caps.maxRecipients) {
      throw new ImapSmtpError(
        'too-many-recipients',
        `Recipient count ${recipients} exceeds per-driver cap ${caps.maxRecipients}`,
      )
    }
    for (const att of req.attachments ?? []) {
      if (att.size > caps.maxAttachmentBytes) {
        throw new ImapSmtpError(
          'attachment-too-large',
          `Attachment ${att.filename} (${att.size} bytes) exceeds per-driver cap ${caps.maxAttachmentBytes}`,
        )
      }
    }
    const total = sendPayloadBytes(req)
    if (total > caps.maxTotalBytes) {
      throw new ImapSmtpError(
        'oversize',
        `Outgoing mail (${total} bytes) exceeds per-driver cap ${caps.maxTotalBytes}`,
      )
    }
    if (total > caps.warnTotalBytes) {
      warnings.push(`Outgoing mail is large (${total} bytes); some servers reject mail above 25 MB`)
    }
    return warnings
  }

  async listFolders(): Promise<MailFolder[]> {
    const imap = this.requireImap()
    const boxes = await imap.listMailboxes()
    return boxes.map((box) => ({
      id: box.path,
      name: box.name,
      role: box.role,
      unread: box.unread,
    }))
  }

  async list(
    folderId: string,
    range: { afterUid?: string; limit: number },
  ): Promise<MailEnvelope[]> {
    const imap = this.requireImap()
    const status = await imap.selectMailbox(folderId)
    const stored = this.deps.syncStore.getSyncState(this.config.accountId, folderId)

    if (stored && stored.uidValidity !== status.uidValidity) {
      const rows = await this.fetchAll(imap, folderId)
      this.deps.syncStore.setSyncState({
        accountId: this.config.accountId,
        folderId,
        uidValidity: status.uidValidity,
        highestModseq: status.highestModseq ?? stored.highestModseq,
        lastSyncAt: new Date().toISOString(),
      })
      this.emit({
        kind: 'resync',
        accountId: this.config.accountId,
        folderId,
        reason: 'uidvalidity-changed',
      })
      return rows.map((row) => toEnvelope(folderId, row))
    }

    if (imap.supports('CONDSTORE') && stored?.highestModseq && range.afterUid === undefined) {
      const delta = await imap.fetchSinceModseq(folderId, stored.highestModseq)
      const highest = highestModseqOf(delta) ?? status.highestModseq ?? stored.highestModseq
      this.deps.syncStore.setSyncState({
        accountId: this.config.accountId,
        folderId,
        uidValidity: status.uidValidity,
        highestModseq: highest,
        lastSyncAt: new Date().toISOString(),
      })
      return delta.map((row) => toEnvelope(folderId, row))
    }

    const rows = await imap.fetchRange(folderId, range.afterUid, range.limit)
    const highest = highestModseqOf(rows) ?? status.highestModseq ?? stored?.highestModseq
    this.deps.syncStore.setSyncState({
      accountId: this.config.accountId,
      folderId,
      uidValidity: status.uidValidity,
      highestModseq: highest,
      lastSyncAt: new Date().toISOString(),
    })
    return rows.map((row) => toEnvelope(folderId, row))
  }

  private async fetchAll(imap: ImapClientLike, folderId: string): Promise<ImapEnvelopeData[]> {
    const all: ImapEnvelopeData[] = []
    let cursor: string | undefined
    for (;;) {
      const page = await imap.fetchRange(folderId, cursor, 200)
      if (page.length === 0) break
      all.push(...page)
      if (page.length < 200) break
      cursor = page[page.length - 1]?.uid
      if (cursor === undefined) break
    }
    return all
  }

  async search(folderId: string, query: MailQuery): Promise<MailEnvelope[]> {
    const imap = this.requireImap()
    const rows = await imap.searchMessages(folderId, query)
    return rows.map((row) => toEnvelope(folderId, row))
  }

  async fetch(ref: MailMessageRef): Promise<MailBody> {
    const imap = this.requireImap()
    const body = await imap.fetchBody(ref.folderId, ref.uid)
    // Full HTML stays content-addressed on disk (P3 store); the driver returns text.
    return { text: body.text }
  }

  async send(req: SendRequest): Promise<SendResult> {
    const warnings = this.checkSendCaps(req)
    const outboxId = this.deps.outbox.enqueueSend(this.config.accountId, req)
    if (!this.connected || !this.smtp) {
      // Offline: the item stays queued (client wins for outbox, spec 4.3.2).
      return { outboxId, warnings }
    }
    try {
      const { messageId } = await this.transmit(req)
      this.deps.outbox.markOutbox?.(outboxId, 'sent')
      return { outboxId, messageId, warnings }
    } catch (err) {
      this.deps.outbox.markOutbox?.(outboxId, 'failed', messageOf(err))
      throw err
    }
  }

  /** Transmits one request over SMTP. Public so the sync engine can flush the outbox. */
  async transmit(req: SendRequest): Promise<{ messageId: string }> {
    if (!this.connected || !this.smtp) {
      throw new ImapSmtpError('offline', 'SMTP transport is not connected')
    }
    return this.smtp.send({
      from: this.config.smtp.from,
      to: req.to,
      cc: req.cc,
      bcc: req.bcc,
      subject: req.subject,
      text: req.text,
      html: req.html,
      attachments: (req.attachments ?? []).map(toSmtpAttachment),
    })
  }

  async move(ref: MailMessageRef, toFolderId: string): Promise<void> {
    if (!this.connected) {
      this.queueOfflineIntent('move', ref, { folderId: ref.folderId, toFolderId })
      return
    }
    await this.requireImap().moveMessage(ref.folderId, ref.uid, toFolderId)
  }

  async setFlags(ref: MailMessageRef, flags: Partial<MailFlags>): Promise<void> {
    if (!this.connected) {
      this.queueOfflineIntent('flag', ref, { folderId: ref.folderId, flags })
      return
    }
    await this.requireImap().setFlags(ref.folderId, ref.uid, flags)
  }

  async delete(ref: MailMessageRef): Promise<void> {
    if (!this.connected) {
      this.queueOfflineIntent('delete', ref, { folderId: ref.folderId })
      return
    }
    await this.requireImap().deleteMessage(ref.folderId, ref.uid)
  }

  private queueOfflineIntent(
    kind: MailIntent['kind'],
    ref: MailMessageRef,
    payload: QueuedIntentPayload,
  ): void {
    const stored = this.deps.syncStore.getSyncState(this.config.accountId, ref.folderId)
    this.deps.intentStore.queueIntent({
      accountId: this.config.accountId,
      kind,
      targetUid: ref.uid,
      expectedModseq: stored?.highestModseq,
      payloadJson: JSON.stringify(payload),
    })
  }

  /**
   * Replays queued offline intents. Each intent is idempotent; a stale
   * modseq re-fetches then retries once, else the intent fails and a
   * conflict event surfaces (spec 4.3.3, never a silent merge guess).
   */
  async drainIntents(): Promise<DrainResult> {
    const imap = this.requireImap()
    const result: DrainResult = { acked: [], failed: [] }
    for (const intent of this.deps.intentStore.pendingIntents(this.config.accountId)) {
      try {
        await this.executeIntent(imap, intent, false)
        this.deps.intentStore.markIntent(intent.id, 'acked')
        result.acked.push(intent.id)
      } catch (err) {
        this.deps.intentStore.markIntent(intent.id, 'failed')
        result.failed.push({ id: intent.id, error: messageOf(err) })
      }
    }
    return result
  }

  private async executeIntent(
    imap: ImapClientLike,
    intent: MailIntent,
    retried: boolean,
  ): Promise<void> {
    const payload = parseIntentPayload(intent)
    const guarded =
      intent.expectedModseq !== undefined ? { expectedModseq: intent.expectedModseq } : undefined
    try {
      if (intent.kind === 'move' && payload.toFolderId) {
        await imap.moveMessage(payload.folderId, intent.targetUid, payload.toFolderId, guarded)
      } else if (intent.kind === 'flag' && payload.flags) {
        await imap.setFlags(payload.folderId, intent.targetUid, payload.flags, guarded)
      } else if (intent.kind === 'delete') {
        await imap.deleteMessage(payload.folderId, intent.targetUid)
      } else {
        throw new ImapSmtpError('conflict', `Intent ${intent.id} has an unsupported payload`)
      }
    } catch (err) {
      if (err instanceof StaleModseqError && !retried) {
        // Re-fetch the target window, then retry exactly once.
        await imap.fetchRange(payload.folderId, uidBefore(intent.targetUid), 2)
        await this.executeIntent(imap, intent, true)
        return
      }
      if (err instanceof StaleModseqError) {
        this.emit({
          kind: 'conflict',
          accountId: this.config.accountId,
          folderId: payload.folderId,
          uid: intent.targetUid,
          detail: `Intent ${intent.id} still stale after re-fetch; left for the user`,
        })
      }
      throw err
    }
  }

  watch(handler: (ev: MailEvent) => void): { stop(): void } {
    const imap = this.requireImap()
    const notify = (ev: MailEvent): void => {
      this.emit(ev)
      handler(ev)
    }
    if (imap.supports('IDLE')) {
      const sub = imap.idle((ev) => {
        if (ev.kind === 'new')
          notify({
            kind: 'messages-added',
            accountId: this.config.accountId,
            folderId: ev.path,
            refs: [eventRef(ev.path, ev.uid)],
          })
        else if (ev.kind === 'updated')
          notify({
            kind: 'flags-changed',
            accountId: this.config.accountId,
            ref: eventRef(ev.path, ev.uid),
            flags: {},
          })
        else
          notify({
            kind: 'message-deleted',
            accountId: this.config.accountId,
            ref: eventRef(ev.path, ev.uid),
          })
      })
      return { stop: () => sub.stop() }
    }
    return this.startPollLoop(notify)
  }

  private startPollLoop(notify: (ev: MailEvent) => void): { stop(): void } {
    const imap = this.requireImap()
    const base = this.config.pollIntervalMs ?? DEFAULT_POLL_INTERVAL_MS
    const ceiling = this.config.maxPollBackoffMs ?? MAX_POLL_BACKOFF_MS
    let stopped = false
    let backoff = base
    let timer: ReturnType<typeof setTimeout> | undefined

    const pollOnce = async (): Promise<void> => {
      const folders = await imap.listMailboxes()
      for (const folder of folders) {
        const status = await imap.selectMailbox(folder.path)
        const stored = this.deps.syncStore.getSyncState(this.config.accountId, folder.path)
        if (stored && stored.uidValidity !== status.uidValidity) {
          await this.fetchAll(imap, folder.path)
          this.deps.syncStore.setSyncState({
            accountId: this.config.accountId,
            folderId: folder.path,
            uidValidity: status.uidValidity,
            highestModseq: status.highestModseq ?? stored.highestModseq,
            lastSyncAt: new Date().toISOString(),
          })
          notify({
            kind: 'resync',
            accountId: this.config.accountId,
            folderId: folder.path,
            reason: 'uidvalidity-changed',
          })
          continue
        }
        const watermark = this.pollWatermarks.get(folder.path)
        let rows: ImapEnvelopeData[]
        if (imap.supports('CONDSTORE') && stored?.highestModseq) {
          rows = await imap.fetchSinceModseq(folder.path, stored.highestModseq)
        } else {
          rows = await imap.fetchRange(folder.path, watermark, 200)
        }
        for (const row of rows) {
          const seen = this.pollWatermarks.get(folder.path)
          if (seen === undefined || row.uid > seen) {
            notify({
              kind: 'messages-added',
              accountId: this.config.accountId,
              folderId: folder.path,
              refs: [
                {
                  uid: row.uid,
                  folderId: folder.path,
                  msgIdHeader: row.msgIdHeader,
                  threadId: threadIdFor(folder.path, row),
                },
              ],
            })
          } else {
            notify({
              kind: 'flags-changed',
              accountId: this.config.accountId,
              ref: {
                uid: row.uid,
                folderId: folder.path,
                msgIdHeader: row.msgIdHeader,
                threadId: threadIdFor(folder.path, row),
              },
              flags: { seen: row.seen, flagged: row.flagged },
            })
          }
        }
        const maxUid = maxUidOf(rows) ?? watermark
        if (maxUid !== undefined) this.pollWatermarks.set(folder.path, maxUid)
        const highest = highestModseqOf(rows) ?? status.highestModseq ?? stored?.highestModseq
        this.deps.syncStore.setSyncState({
          accountId: this.config.accountId,
          folderId: folder.path,
          uidValidity: status.uidValidity,
          highestModseq: highest,
          lastSyncAt: new Date().toISOString(),
        })
      }
    }

    const tick = async (): Promise<void> => {
      if (stopped) return
      try {
        await pollOnce()
        backoff = base
      } catch (err) {
        notify({
          kind: 'sync-state',
          accountId: this.config.accountId,
          folderId: '*',
          state: 'error',
          detail: messageOf(err),
        })
        backoff = Math.min(backoff * 2, ceiling)
      }
      if (!stopped) timer = setTimeout(() => void tick(), backoff)
    }
    timer = setTimeout(() => void tick(), base)
    return {
      stop: () => {
        stopped = true
        if (timer) clearTimeout(timer)
      },
    }
  }
}

/** Raised when an intent's expectedModseq no longer matches the server. */
export class StaleModseqError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'StaleModseqError'
  }
}

function toSmtpAttachment(att: StagedAttachment): SmtpAttachmentPayload {
  if (att.path) {
    // Handed to nodemailer as a path so it streams the file at send time.
    return { filename: att.filename, contentType: att.mime, path: att.path }
  }
  return { filename: att.filename, contentType: att.mime, content: att.data ?? new Uint8Array() }
}

function highestModseqOf(rows: ImapEnvelopeData[]): string | undefined {
  let best: string | undefined
  for (const row of rows) {
    if (row.modseq !== undefined && (best === undefined || row.modseq > best)) best = row.modseq
  }
  return best
}

function maxUidOf(rows: ImapEnvelopeData[]): string | undefined {
  let best: string | undefined
  for (const row of rows) {
    if (best === undefined || row.uid > best) best = row.uid
  }
  return best
}

/** UID cursor just before the target, so a re-fetch window includes it. */
function uidBefore(uid: string): string | undefined {
  const n = Number(uid)
  return Number.isInteger(n) && n > 0 ? String(n - 1) : undefined
}

function messageOf(err: unknown): string {
  return err instanceof Error ? err.message : String(err)
}
