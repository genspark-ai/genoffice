/**
 * GraphDriver: MailProvider over Microsoft Graph (Outlook mail).
 *
 * Pure TypeScript: no Electron, no safeStorage, no direct sockets. The
 * main-process caller (apps/mail/src/main, per locked decision D6) injects
 * the GraphTransport (built with a vault-backed GraphAuthProvider) and the
 * staged-attachment reader. This driver NEVER imports safeStorage,
 * electron, or shell.openExternal; the security-eyes lane (Alya) verifies
 * that with a grep before merge. OAuth scope, PKCE system-browser flow,
 * token refresh, and invalid_grant wipe + re-auth all live in main.
 *
 * Sync mechanics (amended spec 3.2/4.3): delta-query sync per folder.
 * An expired delta token reports resyncRequired (never throws, never
 * retries the dead token); the sync engine then runs a full resync via
 * decideSync({ deltaExpired: true }). Resync never deletes queued outbox
 * sends, drafts, or intents: the driver owns none of those tables (they
 * live in MailStore/sync-engine, P3 lane), so there is nothing here that
 * could delete them; the delta-expiry fixture pins that.
 *
 * Mail only (D10): calendar/contacts endpoints are never called.
 */

import { readFile } from 'node:fs/promises'
import { makeSnippet, singleLine, stripTags } from './mime'
import {
  GRAPH_MAIL_SCOPES,
  GraphServiceError,
  isAuthFailure,
  isDeltaExpiredError,
  isInvalidGrant,
  type GraphTransport,
} from './graph-transport'
import type {
  MailBody,
  MailEnvelope,
  MailEvent,
  MailFlags,
  MailFolder,
  MailMessageRef,
  MailProvider,
  MailQuery,
  SendRequest,
} from './types'

export { GRAPH_MAIL_SCOPES }

export const DEFAULT_POLL_INTERVAL_MS = 45_000
export const MAX_POLL_BACKOFF_MS = 300_000

/**
 * Graph sendMail JSON ceiling is 4 MiB per request (larger sends need an
 * upload session, out of P1 scope). Warn fires below the ceiling so
 * compose can confirm before the driver rejects.
 */
export const GRAPH_WARN_TOTAL_BYTES = 3 * 1024 * 1024
export const GRAPH_MAX_TOTAL_BYTES = 4 * 1024 * 1024
export const GRAPH_MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024
/** Exchange Online caps a message at 500 recipients. */
export const GRAPH_MAX_RECIPIENTS = 500

export type GraphDriverErrorCode =
  | 'offline'
  | 'oversize'
  | 'too-many-recipients'
  | 'attachment-too-large'
  | 'conflict'
  | 'network'
  | 'auth'
  | 'delta-expired'

export class GraphDriverError extends Error {
  readonly code: GraphDriverErrorCode
  /**
   * True when the identity layer reported invalid_grant: main must wipe
   * the stored credential blob and require re-auth (spec 7.1). The driver
   * never wipes or re-authenticates itself.
   */
  readonly needsReauth: boolean
  constructor(
    code: GraphDriverErrorCode,
    message: string,
    options?: { cause?: unknown; needsReauth?: boolean },
  ) {
    super(message)
    this.name = 'GraphDriverError'
    this.code = code
    this.needsReauth = options?.needsReauth ?? false
    if (options?.cause !== undefined) {
      ;(this as { cause?: unknown }).cause = options.cause
    }
  }
}

/** Per-driver send limits, enforced in the driver and surfaced to compose. */
export interface GraphSendCaps {
  /** Above this total, checkSendSize reports warn. Default 3 MiB. */
  warnTotalBytes?: number
  /** Above this total, send is rejected. Default 4 MiB (Graph ceiling). */
  maxTotalBytes?: number
  /** Above this per-file size, send is rejected. Default 3 MiB. */
  maxAttachmentBytes?: number
  /** Above this recipient count, send is rejected. Default 500. */
  maxRecipients?: number
}

export interface GraphDriverConfig {
  accountId: string
  transport: GraphTransport
  /** Reads staged attachment bytes by StagedAttachment.path. Default: node fs. */
  readStagedFile?: (path: string) => Promise<Uint8Array>
  /** Generates outbox ids returned by send(). Default: time + counter. */
  newId?: () => string
  sendCaps?: GraphSendCaps
  /** Poll cadence for watch(). Default 45 s. */
  pollIntervalMs?: number
  /** Backoff ceiling for failed polls. Default 5 min. */
  maxPollBackoffMs?: number
  /** Folder ids watched by watch(). Default ['inbox']. */
  watchFolders?: string[]
}

/** Minimal Graph message shape the driver reads. */
export interface GraphMessage {
  id: string
  conversationId?: string
  internetMessageId?: string
  subject?: string
  from?: { emailAddress?: { address?: string; name?: string } }
  toRecipients?: Array<{ emailAddress?: { address?: string } }>
  ccRecipients?: Array<{ emailAddress?: { address?: string } }>
  receivedDateTime?: string
  isRead?: boolean
  flag?: { flagStatus?: string }
  categories?: string[]
  hasAttachments?: boolean
  bodyPreview?: string
  body?: { contentType?: string; content?: string }
  parentFolderId?: string
  '@removed'?: { reason?: string }
}

export interface GraphFolder {
  id: string
  displayName?: string
  unreadItemCount?: number
}

interface GraphPage<T> {
  value?: T[]
  '@odata.nextLink'?: string
  '@odata.deltaLink'?: string
}

export interface DeltaSyncResult {
  envelopes: MailEnvelope[]
  removed: MailMessageRef[]
  flagChanges: Array<{ ref: MailMessageRef; flags: Partial<MailFlags> }>
  /** Fresh delta link for SyncState.deltaToken. Absent when resyncRequired. */
  deltaLink?: string
  /**
   * True when the server rejected the delta token: run a full resync
   * (decideSync deltaExpired) instead of retrying. Queued outbox, drafts,
   * and intents are preserved by the caller; this result deletes nothing.
   */
  resyncRequired: boolean
}

export interface SendSizeVerdict {
  totalBytes: number
  recipients: number
  verdict: 'ok' | 'warn'
  warnings: string[]
}

const MESSAGE_SELECT =
  'id,conversationId,internetMessageId,subject,from,toRecipients,ccRecipients,' +
  'receivedDateTime,isRead,flag,categories,hasAttachments,bodyPreview,body,parentFolderId'

/** Well-known folder names (Graph path segments) mapped to MailFolder roles. */
const WELL_KNOWN_ROLES: Array<{ name: string; role: MailFolder['role'] }> = [
  { name: 'inbox', role: 'inbox' },
  { name: 'sentitems', role: 'sent' },
  { name: 'drafts', role: 'drafts' },
  { name: 'deleteditems', role: 'trash' },
  { name: 'junkemail', role: 'spam' },
  { name: 'archive', role: 'archive' },
]

function utf8Bytes(text: string): number {
  return new TextEncoder().encode(text).byteLength
}

function addrOf(list: Array<{ emailAddress?: { address?: string } }> | undefined): string[] {
  return (list ?? []).map((r) => r.emailAddress?.address ?? '').filter((a) => a !== '')
}

function toFlags(msg: GraphMessage): MailFlags {
  return {
    seen: msg.isRead ?? false,
    flagged: (msg.flag?.flagStatus ?? 'notFlagged') === 'flagged',
    // Graph exposes no answered/draft bits; those flags stay local to MailStore.
    answered: false,
    draft: false,
  }
}

function isNotFound(err: unknown): boolean {
  return err instanceof GraphServiceError && err.status === 404
}

function mapServiceError(err: unknown, ref?: string): GraphDriverError {
  if (err instanceof GraphDriverError) return err
  if (isDeltaExpiredError(err)) {
    return new GraphDriverError(
      'delta-expired',
      'Graph delta token expired; full resync required',
      { cause: err },
    )
  }
  if (err instanceof GraphServiceError) {
    if (isInvalidGrant(err)) {
      return new GraphDriverError(
        'auth',
        'Graph credential rejected (invalid_grant); wipe and re-auth in main',
        {
          cause: err,
          needsReauth: true,
        },
      )
    }
    if (isAuthFailure(err)) {
      return new GraphDriverError('auth', `Graph authentication failed: ${err.code}`, {
        cause: err,
      })
    }
    if (err.status === 404) {
      return new GraphDriverError(
        'conflict',
        `Graph target gone${ref === undefined ? '' : `: ${ref}`}`,
        { cause: err },
      )
    }
    if (err.status === 429 || err.status >= 500 || err.status === 0) {
      return new GraphDriverError('network', `Graph transport failure: ${err.code}`, { cause: err })
    }
    return new GraphDriverError('network', `Graph request failed: ${err.code}`, { cause: err })
  }
  return new GraphDriverError('network', 'Graph request failed', { cause: err })
}

export class GraphDriver implements MailProvider {
  private readonly accountId: string
  private readonly transport: GraphTransport
  private readonly readStagedFile: (path: string) => Promise<Uint8Array>
  private readonly newId: () => string
  private readonly caps: Required<GraphSendCaps>
  private readonly pollIntervalMs: number
  private readonly maxPollBackoffMs: number
  private readonly watchFolders: string[]
  private connected = false
  private idCounter = 0

  constructor(config: GraphDriverConfig) {
    this.accountId = config.accountId
    this.transport = config.transport
    this.readStagedFile = config.readStagedFile ?? ((path) => readFile(path))
    let counter = 0
    this.newId =
      config.newId ??
      (() => {
        counter += 1
        return `graph-${Date.now().toString(36)}-${counter}`
      })
    this.caps = {
      warnTotalBytes: config.sendCaps?.warnTotalBytes ?? GRAPH_WARN_TOTAL_BYTES,
      maxTotalBytes: config.sendCaps?.maxTotalBytes ?? GRAPH_MAX_TOTAL_BYTES,
      maxAttachmentBytes: config.sendCaps?.maxAttachmentBytes ?? GRAPH_MAX_ATTACHMENT_BYTES,
      maxRecipients: config.sendCaps?.maxRecipients ?? GRAPH_MAX_RECIPIENTS,
    }
    this.pollIntervalMs = config.pollIntervalMs ?? DEFAULT_POLL_INTERVAL_MS
    this.maxPollBackoffMs = config.maxPollBackoffMs ?? MAX_POLL_BACKOFF_MS
    this.watchFolders = config.watchFolders ?? ['inbox']
  }

  private toEnvelope(msg: GraphMessage, folderId: string): MailEnvelope {
    const text =
      msg.body?.contentType === 'text'
        ? (msg.body.content ?? '')
        : stripTags(msg.body?.content ?? '')
    return {
      ref: {
        uid: msg.id,
        folderId: msg.parentFolderId ?? folderId,
        msgIdHeader: msg.internetMessageId,
        threadId: msg.conversationId ?? msg.id,
      },
      from: msg.from?.emailAddress?.address === undefined ? [] : [msg.from.emailAddress.address],
      to: addrOf(msg.toRecipients),
      cc: addrOf(msg.ccRecipients),
      subject: singleLine(msg.subject),
      date: msg.receivedDateTime ?? '',
      flags: toFlags(msg),
      labels: msg.categories ?? [],
      snippet: makeSnippet(msg.bodyPreview ?? text),
      hasAttach: msg.hasAttachments ?? false,
    }
  }

  async connect(): Promise<void> {
    try {
      // Identity probe only (D10 mail scope still needs the user record).
      await this.transport.get('/me?$select=id')
      this.connected = true
    } catch (err) {
      throw mapServiceError(err)
    }
  }

  async disconnect(): Promise<void> {
    this.connected = false
  }

  private requireConnected(): void {
    if (!this.connected) throw new GraphDriverError('offline', 'Graph driver is not connected')
  }

  async listFolders(): Promise<MailFolder[]> {
    this.requireConnected()
    try {
      const page = await this.transport.get<GraphPage<GraphFolder>>('/me/mailFolders?$top=100')
      const folders = page.value ?? []
      // Resolve well-known names to ids so roles survive localized display names.
      const roleById = new Map<string, MailFolder['role']>()
      for (const wk of WELL_KNOWN_ROLES) {
        try {
          const known = await this.transport.get<GraphFolder>(
            `/me/mailFolders/${wk.name}?$select=id`,
          )
          if (known.id !== undefined) roleById.set(known.id, wk.role)
        } catch (err) {
          if (!isNotFound(err)) throw err
        }
      }
      return folders.map((f) => ({
        id: f.id,
        name: f.displayName ?? f.id,
        role: roleById.get(f.id) ?? 'custom',
        unread: f.unreadItemCount ?? 0,
      }))
    } catch (err) {
      throw mapServiceError(err)
    }
  }

  async list(
    folderId: string,
    range: { afterUid?: string; limit: number },
  ): Promise<MailEnvelope[]> {
    this.requireConnected()
    try {
      let path: string
      if (range.afterUid === undefined) {
        path =
          `/me/mailFolders/${encodeURIComponent(folderId)}/messages` +
          `?$top=${range.limit}&$orderby=receivedDateTime desc&$select=${MESSAGE_SELECT}`
      } else if (range.afterUid.startsWith('https://')) {
        // Opaque service paging cursor (a nextLink returned earlier).
        path = range.afterUid
      } else {
        // Message-id cursor: page by date below the anchor message.
        const anchor = await this.transport.get<GraphMessage>(
          `/me/messages/${encodeURIComponent(range.afterUid)}?$select=receivedDateTime`,
        )
        const date = anchor.receivedDateTime ?? ''
        path =
          `/me/mailFolders/${encodeURIComponent(folderId)}/messages` +
          `?$top=${range.limit}&$orderby=receivedDateTime desc` +
          `&$filter=receivedDateTime lt ${encodeURIComponent(date)}&$select=${MESSAGE_SELECT}`
      }
      const page = await this.transport.get<GraphPage<GraphMessage>>(path)
      return (page.value ?? []).map((m) => this.toEnvelope(m, folderId))
    } catch (err) {
      throw mapServiceError(err)
    }
  }

  async search(folderId: string, query: MailQuery): Promise<MailEnvelope[]> {
    this.requireConnected()
    const parts: string[] = []
    if (query.text !== undefined && query.text !== '')
      parts.push(`"${query.text.replace(/"/g, '')}"`)
    if (query.from !== undefined && query.from !== '') parts.push(`from:${query.from}`)
    if (query.since !== undefined && query.since !== '') parts.push(`received>=${query.since}`)
    if (query.onlyUnread === true) parts.push('isRead:false')
    try {
      const path =
        `/me/mailFolders/${encodeURIComponent(folderId)}/messages` +
        `?$search=${encodeURIComponent(parts.join(' '))}&$top=25&$select=${MESSAGE_SELECT}`
      const page = await this.transport.get<GraphPage<GraphMessage>>(path, {
        headers: { ConsistencyLevel: 'eventual' },
      })
      return (page.value ?? []).map((m) => this.toEnvelope(m, folderId))
    } catch (err) {
      throw mapServiceError(err)
    }
  }

  async fetch(ref: MailMessageRef): Promise<MailBody> {
    this.requireConnected()
    try {
      const msg = await this.transport.get<GraphMessage>(
        `/me/messages/${encodeURIComponent(ref.uid)}?$select=${MESSAGE_SELECT}`,
      )
      const html = msg.body?.contentType === 'text' ? undefined : msg.body?.content
      const text =
        msg.body?.contentType === 'text' ? (msg.body.content ?? '') : stripTags(html ?? '')
      return { text }
    } catch (err) {
      throw mapServiceError(err, ref.uid)
    }
  }

  /**
   * Full HTML body for sync-engine persistence (content-addressed on disk
   * by MailStore; never kept in the DB). Undefined for text-only bodies.
   */
  async fetchMessageHtml(ref: MailMessageRef): Promise<string | undefined> {
    this.requireConnected()
    try {
      const msg = await this.transport.get<GraphMessage>(
        `/me/messages/${encodeURIComponent(ref.uid)}?$select=body`,
      )
      if (msg.body?.contentType === 'text') return undefined
      return msg.body?.content
    } catch (err) {
      throw mapServiceError(err, ref.uid)
    }
  }

  /**
   * Compose pre-send verdict: ok or warn with human-readable warnings.
   * Rejections throw from send(); this method never throws for size.
   */
  checkSendSize(req: SendRequest): SendSizeVerdict {
    const recipients = req.to.length + (req.cc?.length ?? 0) + (req.bcc?.length ?? 0)
    let totalBytes = utf8Bytes(req.text) + (req.html === undefined ? 0 : utf8Bytes(req.html))
    for (const a of req.attachments ?? []) totalBytes += a.size
    const warnings: string[] = []
    if (totalBytes > this.caps.warnTotalBytes) {
      warnings.push(
        `Outgoing mail is ${(totalBytes / 1048576).toFixed(1)} MiB; Graph sendMail caps at 4 MiB`,
      )
    }
    return { totalBytes, recipients, verdict: warnings.length === 0 ? 'ok' : 'warn', warnings }
  }

  private checkSendCapsOrThrow(req: SendRequest): void {
    const recipients = req.to.length + (req.cc?.length ?? 0) + (req.bcc?.length ?? 0)
    if (recipients > this.caps.maxRecipients) {
      throw new GraphDriverError(
        'too-many-recipients',
        `Recipient count ${recipients} exceeds the Graph limit of ${this.caps.maxRecipients}`,
      )
    }
    for (const a of req.attachments ?? []) {
      if (a.size > this.caps.maxAttachmentBytes) {
        throw new GraphDriverError(
          'attachment-too-large',
          `Attachment ${a.filename} exceeds the per-file limit of ${this.caps.maxAttachmentBytes} bytes`,
        )
      }
    }
    let totalBytes = utf8Bytes(req.text) + (req.html === undefined ? 0 : utf8Bytes(req.html))
    for (const a of req.attachments ?? []) totalBytes += a.size
    if (totalBytes > this.caps.maxTotalBytes) {
      throw new GraphDriverError(
        'oversize',
        `Outgoing mail size ${totalBytes} exceeds the Graph sendMail ceiling of ${this.caps.maxTotalBytes} bytes`,
      )
    }
  }

  async send(req: SendRequest): Promise<{ outboxId: string }> {
    this.requireConnected()
    this.checkSendCapsOrThrow(req)
    try {
      const toRecipients = req.to.map((address) => ({ emailAddress: { address } }))
      const ccRecipients = (req.cc ?? []).map((address) => ({ emailAddress: { address } }))
      const bccRecipients = (req.bcc ?? []).map((address) => ({ emailAddress: { address } }))
      const attachments: unknown[] = []
      for (const a of req.attachments ?? []) {
        // Canonical payload policy: staged path wins; staged.data is the
        // test/tiny-inline fallback honored by every driver.
        const bytes = a.data ?? (await this.readStagedFile(a.path))
        attachments.push({
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: a.filename,
          contentType: a.mime,
          contentBytes: Buffer.from(bytes).toString('base64'),
        })
      }
      const message: Record<string, unknown> = {
        subject: req.subject,
        body: {
          contentType: req.html === undefined ? 'Text' : 'HTML',
          content: req.html ?? req.text,
        },
        toRecipients,
      }
      if (ccRecipients.length > 0) message['ccRecipients'] = ccRecipients
      if (bccRecipients.length > 0) message['bccRecipients'] = bccRecipients
      if (attachments.length > 0) message['attachments'] = attachments
      await this.transport.post('/me/sendMail', { message, saveToSentItems: 'true' })
      this.idCounter += 1
      return { outboxId: this.newId() }
    } catch (err) {
      throw mapServiceError(err)
    }
  }

  async move(ref: MailMessageRef, toFolderId: string): Promise<void> {
    this.requireConnected()
    try {
      await this.transport.post(`/me/messages/${encodeURIComponent(ref.uid)}/move`, {
        destinationId: toFolderId,
      })
    } catch (err) {
      throw mapServiceError(err, ref.uid)
    }
  }

  async setFlags(ref: MailMessageRef, flags: Partial<MailFlags>): Promise<void> {
    this.requireConnected()
    const patch: Record<string, unknown> = {}
    if (flags.seen !== undefined) patch['isRead'] = flags.seen
    if (flags.flagged !== undefined)
      patch['flag'] = { flagStatus: flags.flagged ? 'flagged' : 'notFlagged' }
    // answered/draft have no Graph property; they stay local to MailStore.
    if (Object.keys(patch).length === 0) return
    try {
      await this.transport.patch(`/me/messages/${encodeURIComponent(ref.uid)}`, patch)
    } catch (err) {
      throw mapServiceError(err, ref.uid)
    }
  }

  async delete(ref: MailMessageRef): Promise<void> {
    this.requireConnected()
    try {
      await this.transport.del(`/me/messages/${encodeURIComponent(ref.uid)}`)
    } catch (err) {
      throw mapServiceError(err, ref.uid)
    }
  }

  /**
   * Delta-query sync for one folder. Pass the stored SyncState.deltaToken;
   * omit it for the first sync. Expired tokens yield resyncRequired (no
   * throw). The returned deltaLink persists as the next SyncState.deltaToken.
   */
  async deltaSync(folderId: string, deltaToken?: string): Promise<DeltaSyncResult> {
    this.requireConnected()
    const empty: DeltaSyncResult = {
      envelopes: [],
      removed: [],
      flagChanges: [],
      resyncRequired: false,
    }
    try {
      const path =
        deltaToken === undefined
          ? `/me/mailFolders/${encodeURIComponent(folderId)}/messages/delta?$select=${MESSAGE_SELECT}`
          : `/me/mailFolders/${encodeURIComponent(folderId)}/messages/delta?$deltatoken=${encodeURIComponent(deltaToken)}&$select=${MESSAGE_SELECT}`
      const page = await this.transport.get<GraphPage<GraphMessage>>(path)
      const result: DeltaSyncResult = { ...empty, deltaLink: page['@odata.deltaLink'] }
      for (const m of page.value ?? []) {
        const ref: MailMessageRef = {
          uid: m.id,
          folderId: m.parentFolderId ?? folderId,
          msgIdHeader: m.internetMessageId,
          threadId: m.conversationId ?? m.id,
        }
        if (m['@removed'] !== undefined) {
          result.removed.push(ref)
          continue
        }
        result.envelopes.push(this.toEnvelope(m, folderId))
        result.flagChanges.push({ ref, flags: toFlags(m) })
      }
      return result
    } catch (err) {
      if (isDeltaExpiredError(err)) return { ...empty, resyncRequired: true }
      throw mapServiceError(err)
    }
  }

  watch(handler: (ev: MailEvent) => void): { stop(): void } {
    let stopped = false
    let timer: ReturnType<typeof setTimeout> | undefined
    let backoffMs = this.pollIntervalMs
    const seen = new Map<string, Map<string, { isRead: boolean; flagStatus: string }>>()
    let baselineDone = false

    const poll = async (): Promise<void> => {
      if (stopped) return
      try {
        handler({
          kind: 'sync-state',
          accountId: this.accountId,
          folderId: this.watchFolders[0] ?? 'inbox',
          state: 'syncing',
        })
        for (const folderId of this.watchFolders) {
          const page = await this.transport.get<GraphPage<GraphMessage>>(
            `/me/mailFolders/${encodeURIComponent(folderId)}/messages?$top=25&$orderby=receivedDateTime desc&$select=${MESSAGE_SELECT}`,
          )
          const folderSeen = seen.get(folderId) ?? new Map()
          const next = new Map<string, { isRead: boolean; flagStatus: string }>()
          const added: MailMessageRef[] = []
          for (const m of page.value ?? []) {
            const flagStatus = m.flag?.flagStatus ?? 'notFlagged'
            next.set(m.id, { isRead: m.isRead ?? false, flagStatus })
            const prev = folderSeen.get(m.id)
            if (prev === undefined) {
              if (baselineDone) {
                added.push({
                  uid: m.id,
                  folderId: m.parentFolderId ?? folderId,
                  msgIdHeader: m.internetMessageId,
                  threadId: m.conversationId ?? m.id,
                })
              }
            } else if (prev.isRead !== (m.isRead ?? false) || prev.flagStatus !== flagStatus) {
              if (baselineDone) {
                handler({
                  kind: 'flags-changed',
                  accountId: this.accountId,
                  ref: {
                    uid: m.id,
                    folderId: m.parentFolderId ?? folderId,
                    msgIdHeader: m.internetMessageId,
                    threadId: m.conversationId ?? m.id,
                  },
                  flags: toFlags(m),
                })
              }
            }
          }
          seen.set(folderId, next)
          if (baselineDone && added.length > 0) {
            handler({ kind: 'messages-added', accountId: this.accountId, folderId, refs: added })
          }
        }
        baselineDone = true
        backoffMs = this.pollIntervalMs
        handler({
          kind: 'sync-state',
          accountId: this.accountId,
          folderId: this.watchFolders[0] ?? 'inbox',
          state: 'idle',
        })
      } catch {
        backoffMs = Math.min(backoffMs * 2, this.maxPollBackoffMs)
        try {
          handler({
            kind: 'sync-state',
            accountId: this.accountId,
            folderId: this.watchFolders[0] ?? 'inbox',
            state: 'error',
            detail: 'Graph poll failed; backing off',
          })
        } catch {
          // Handler failures never break the poll loop.
        }
      }
      if (!stopped) {
        timer = setTimeout(() => void poll(), backoffMs)
        timer.unref?.()
      }
    }

    void poll()
    return {
      stop: () => {
        stopped = true
        if (timer !== undefined) clearTimeout(timer)
      },
    }
  }
}
