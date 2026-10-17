/**
 * GmailDriver: MailProvider over the Gmail API v1 (googleapis).
 *
 * Pure-TS class with injected auth and transport. The driver NEVER touches
 * Electron safeStorage, refresh tokens, or client secrets: the main process
 * owns the PKCE flow, token refresh, and the credential vault, and hands
 * this driver a short-lived access token per call through GmailAuthProvider.
 * Scheduling and secrets-injection stay in apps/mail/src/main (spec D6).
 *
 * Sync uses users.history.list from the stored historyId. An expired
 * historyId (HTTP 404) forces a full resync through decideSync; resync
 * only rewrites envelopes and sync state, it never deletes queued
 * outbox sends, local drafts, or pending intents (the driver holds no
 * reference to those queues at all). An invalid_grant Auth error is
 * surfaced as GmailAuthExpiredError so main can wipe the vault blob and
 * require re-auth; the driver never retries into weaker storage.
 */

import { randomUUID } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { parseMimeMessage } from './mime'
import { formatAddress, parseAddressList } from './addresses'
import { OUTGOING_WARN_BYTES } from './attachments'
import { decideSync } from './sync'
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
  SyncState,
} from './types'

/** Least-privilege OAuth scope for the Gmail driver. Full mail scope is forbidden. */
export const GMAIL_MODIFY_SCOPE = 'https://www.googleapis.com/auth/gmail.modify'

/** Re-exported for driver consumers building requests and handling events. */
export type { MailEvent, SendRequest } from './types'

/** Gmail API per-message send cap (base64url raw payload, 35 MB). */
export const GMAIL_MAX_MESSAGE_BYTES = 35 * 1024 * 1024

/** Compose warning threshold; mirrors the spec 25 MB outgoing warning. */
export const GMAIL_WARN_BYTES = OUTGOING_WARN_BYTES

/** Conservative per-message recipient cap enforced before transmit. */
export const GMAIL_MAX_RECIPIENTS_PER_MESSAGE = 500

/** System label ids Gmail always reports alongside user labels. */
const FLAG_LABELS = new Set(['UNREAD', 'IMPORTANT'])

/** Gmail system label id to MailFolder role. Unlisted labels are custom. */
const LABEL_ROLES: Record<string, MailFolder['role']> = {
  INBOX: 'inbox',
  SENT: 'sent',
  DRAFT: 'drafts',
  TRASH: 'trash',
  SPAM: 'spam',
}

/** Minimal Gmail message header shape (subset of gmail_v1.Schema$MessagePartHeader). */
export interface GmailHeader {
  name?: string | null
  value?: string | null
}

/** Minimal Gmail payload shape (subset of gmail_v1.Schema$MessagePart). */
export interface GmailPayload {
  mimeType?: string | null
  filename?: string | null
  headers?: GmailHeader[] | null
  body?: { data?: string | null; size?: number | null } | null
  parts?: GmailPayload[] | null
}

/** Minimal Gmail message shape (subset of gmail_v1.Schema$Message). */
export interface GmailMessage {
  id?: string | null
  threadId?: string | null
  labelIds?: string[] | null
  snippet?: string | null
  historyId?: string | null
  sizeEstimate?: number | null
  payload?: GmailPayload | null
}

/** Minimal Gmail label shape (subset of gmail_v1.Schema$Label). */
export interface GmailLabel {
  id?: string | null
  name?: string | null
  messagesUnread?: number | null
  type?: string | null
}

/**
 * Short-lived access-token source, implemented in the main process.
 * The driver calls it per request; tokens stay memory-only.
 */
export interface GmailAuthProvider {
  accountId: string
  getAccessToken(): Promise<string>
}

/**
 * Narrow transport over the Gmail API. The production adapter
 * (src/gmail-transport.ts) implements it with googleapis; tests inject
 * an in-memory fake. No real mailbox credentials appear in tests.
 */
export interface GmailTransport {
  listLabels(): Promise<GmailLabel[]>
  listMessages(args: {
    labelIds?: string[]
    q?: string
    maxResults: number
    pageToken?: string
  }): Promise<{
    messages?: { id?: string | null; threadId?: string | null }[]
    nextPageToken?: string | null
  }>
  getMessage(id: string, format: 'metadata' | 'full'): Promise<GmailMessage>
  sendMessage(rawBase64Url: string): Promise<{ id?: string | null }>
  modifyMessage(id: string, addLabelIds: string[], removeLabelIds: string[]): Promise<void>
  trashMessage(id: string): Promise<void>
  listHistory(startHistoryId: string): Promise<{
    historyId: string
    added?: { message?: GmailMessage | null }[]
    deleted?: { message?: GmailMessage | null }[]
    labelsChanged?: { message?: GmailMessage | null }[]
  }>
  getCurrentHistoryId(): Promise<string>
}

/** Sync-state persistence owned by the main process (SQLite sync_state). */
export interface GmailSyncHooks {
  load(accountId: string): SyncState | undefined
  save(state: SyncState): void
}

/** Thrown when Google rejects the access token with invalid_grant. */
export class GmailAuthExpiredError extends Error {
  constructor() {
    super('gmail access token rejected (invalid_grant): wipe vault blob and re-auth')
    this.name = 'GmailAuthExpiredError'
  }
}

/** Thrown when a send request breaks a per-driver send limit. */
export class GmailSendLimitError extends Error {
  readonly code: 'too-large' | 'too-many-recipients'
  constructor(code: 'too-large' | 'too-many-recipients', detail: string) {
    super(`gmail send rejected (${code}): ${detail}`)
    this.name = 'GmailSendLimitError'
    this.code = code
  }
}

/** True when a Gmail API failure means the stored historyId expired (HTTP 404). */
export function isHistoryExpiredError(err: unknown): boolean {
  if (!err || typeof err !== 'object') return false
  const code = (err as { code?: unknown }).code
  return code === 404 || code === '404'
}

/** True when a Gmail API failure is an invalid_grant auth rejection. */
export function isInvalidGrantError(err: unknown): boolean {
  if (!err || typeof err !== 'object') return false
  const rec = err as { code?: unknown; message?: unknown; response?: unknown }
  const hay = `${String(rec.message ?? '')} ${JSON.stringify(rec.response ?? '')}`
  if (/invalid_grant/i.test(hay)) return true
  return rec.code === 401 || rec.code === '401'
}

/** Map one Gmail API failure to the driver-level error policy. */
export function mapGmailError(err: unknown): unknown {
  if (isInvalidGrantError(err)) return new GmailAuthExpiredError()
  return err
}

/** Verdict of the per-driver send-limit check, surfaced in compose. */
export type SendLimitVerdict = 'ok' | 'warn' | 'too-large' | 'too-many-recipients'

/** Count every recipient across to/cc/bcc. */
export function countSendRecipients(req: SendRequest): number {
  return req.to.length + (req.cc?.length ?? 0) + (req.bcc?.length ?? 0)
}

/**
 * Rough encoded size of a send request: UTF-8 body bytes times base64
 * overhead plus staged attachment bytes. Staged files are measured by
 * their recorded size; the driver never reads secrets to estimate.
 */
export function estimateSendBytes(req: SendRequest): number {
  const textBytes = Buffer.byteLength(req.text, 'utf8') + Buffer.byteLength(req.html ?? '', 'utf8')
  const attachBytes = (req.attachments ?? []).reduce((sum, a) => sum + a.size, 0)
  return Math.ceil(((textBytes + attachBytes) * 4) / 3) + 4096
}

/**
 * Enforce per-driver send limits. Returns ok/warn, or the rejection
 * reason when the request must not be transmitted.
 */
export function checkSendLimits(req: SendRequest, estimatedBytes?: number): SendLimitVerdict {
  if (countSendRecipients(req) > GMAIL_MAX_RECIPIENTS_PER_MESSAGE) return 'too-many-recipients'
  const bytes = estimatedBytes ?? estimateSendBytes(req)
  if (bytes > GMAIL_MAX_MESSAGE_BYTES) return 'too-large'
  if (bytes > GMAIL_WARN_BYTES) return 'warn'
  return 'ok'
}

/** Map a Gmail label to a MailFolder; UNREAD/IMPORTANT are flags, not folders. */
export function gmailLabelToFolder(label: GmailLabel): MailFolder | undefined {
  const id = label.id ?? ''
  if (id === '' || FLAG_LABELS.has(id)) return undefined
  return {
    id,
    name: label.name ?? id,
    role: LABEL_ROLES[id] ?? 'custom',
    unread: label.messagesUnread ?? 0,
  }
}

/** Map Gmail label ids to MailFlags. Gmail has no answered flag; it reads false. */
export function gmailLabelsToFlags(labelIds: string[] | null | undefined): MailFlags {
  const labels = new Set(labelIds ?? [])
  return {
    seen: !labels.has('UNREAD'),
    flagged: labels.has('STARRED'),
    answered: false,
    draft: labels.has('DRAFT'),
  }
}

/** Read one header value case-insensitively; first match wins. */
function headerOf(headers: GmailHeader[] | null | undefined, name: string): string {
  const want = name.toLowerCase()
  for (const h of headers ?? []) {
    if ((h.name ?? '').toLowerCase() === want) return h.value ?? ''
  }
  return ''
}

/** Normalize an address-list header to display strings; never throws. */
function addressStrings(header: string): string[] {
  if (header.trim() === '') return []
  try {
    return parseAddressList(header).map(formatAddress)
  } catch {
    return [header]
  }
}

/** Collect every body data chunk of a MIME type from a payload tree. */
function collectBodies(part: GmailPayload | null | undefined, mime: string, out: string[]): void {
  if (!part) return
  if ((part.mimeType ?? '').toLowerCase().startsWith(mime) && part.body?.data)
    out.push(part.body.data)
  for (const child of part.parts ?? []) collectBodies(child, mime, out)
}

/** True when any payload part carries a filename (attachment present). */
function payloadHasAttachment(part: GmailPayload | null | undefined): boolean {
  if (!part) return false
  if ((part.filename ?? '') !== '') return true
  return (part.parts ?? []).some(payloadHasAttachment)
}

/** Decode base64url Gmail bodies to UTF-8 text. */
export function decodeGmailBody(data: string): string {
  const b64 = data.replace(/-/g, '+').replace(/_/g, '/')
  return Buffer.from(b64, 'base64').toString('utf8')
}

/** Strip HTML tags for the plain-text fallback; the raw HTML stays on disk via fetchRaw. */
function stripHtml(html: string): string {
  return html
    .replace(/<script[\s\S]*?<\/script\s*>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style\s*>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

/** Raw message fetch: text plus optional HTML for the main layer to persist. */
export interface GmailRawBody {
  text: string
  html?: string
}

/** Extract text (and HTML when present) from a full Gmail message payload. */
export function extractGmailBody(message: GmailMessage): GmailRawBody {
  const plains: string[] = []
  const htmls: string[] = []
  collectBodies(message.payload, 'text/plain', plains)
  collectBodies(message.payload, 'text/html', htmls)
  const text =
    plains.length > 0
      ? plains.map(decodeGmailBody).join('\n')
      : htmls.length > 0
        ? stripHtml(htmls.map(decodeGmailBody).join('\n'))
        : (message.snippet ?? '')
  const html = htmls.length > 0 ? htmls.map(decodeGmailBody).join('\n') : undefined
  return html === undefined ? { text } : { text, html }
}

/** Convert a Gmail message (metadata or full) to a MailEnvelope view-model. */
export function gmailMessageToEnvelope(
  message: GmailMessage,
  folderId: string,
  headers?: GmailHeader[],
): MailEnvelope {
  const hdrs = headers ?? message.payload?.headers ?? []
  const uid = message.id ?? ''
  return {
    ref: { uid, folderId, threadId: message.threadId ?? uid },
    from: addressStrings(headerOf(hdrs, 'from')),
    to: addressStrings(headerOf(hdrs, 'to')),
    cc: addressStrings(headerOf(hdrs, 'cc')),
    subject: headerOf(hdrs, 'subject'),
    date: headerOf(hdrs, 'date'),
    flags: gmailLabelsToFlags(message.labelIds),
    labels: [...(message.labelIds ?? [])],
    snippet: message.snippet ?? '',
    hasAttach: payloadHasAttachment(message.payload),
  }
}

/** Build the Gmail search query for a MailQuery (server expansion is explicit). */
export function buildGmailQuery(query: MailQuery): string {
  const parts: string[] = []
  if (query.text && query.text.trim() !== '') parts.push(query.text.trim())
  if (query.from && query.from.trim() !== '') parts.push(`from:${query.from.trim()}`)
  if (query.since && query.since.trim() !== '') {
    const day = query.since.slice(0, 10).replace(/-/g, '/')
    parts.push(`after:${day}`)
  }
  if (query.onlyUnread === true) parts.push('is:unread')
  return parts.join(' ')
}

/** Encode a header value with RFC 2047 when it leaves ASCII. */
function encodeHeaderValue(value: string): string {
  if (/^[\x20-\x7e]*$/.test(value)) return value
  return `=?UTF-8?B?${Buffer.from(value, 'utf8').toString('base64')}?=`
}

export interface RawAttachmentInput {
  filename: string
  mime: string
  bytes: Uint8Array
}

export interface RawMessageInput {
  from: string
  to: string[]
  cc?: string[]
  bcc?: string[]
  subject: string
  text: string
  html?: string
  attachments?: RawAttachmentInput[]
}

/**
 * Build an RFC 5322 message for users.messages.send. Pure and deterministic
 * so tests can round-trip it through the MIME parser. Callers pass ASCII
 * or UTF-8 strings; the builder folds headers and encodes bodies.
 */
export function buildRawMessage(input: RawMessageInput): string {
  const boundaryMix = `genoffice-mixed-${randomUUID().replace(/-/g, '')}`
  const boundaryAlt = `genoffice-alt-${randomUUID().replace(/-/g, '')}`
  const lines: string[] = []
  lines.push(`From: ${input.from}`)
  lines.push(`To: ${input.to.join(', ')}`)
  if (input.cc && input.cc.length > 0) lines.push(`Cc: ${input.cc.join(', ')}`)
  if (input.bcc && input.bcc.length > 0) lines.push(`Bcc: ${input.bcc.join(', ')}`)
  lines.push(`Subject: ${encodeHeaderValue(input.subject)}`)
  lines.push('MIME-Version: 1.0')
  const hasAttach = (input.attachments?.length ?? 0) > 0
  const textB64 = Buffer.from(input.text, 'utf8').toString('base64')
  const htmlB64 =
    input.html !== undefined ? Buffer.from(input.html, 'utf8').toString('base64') : undefined
  const pushAlt = (target: string[]): void => {
    target.push(`--${boundaryAlt}`)
    target.push('Content-Type: text/plain; charset=UTF-8')
    target.push('Content-Transfer-Encoding: base64')
    target.push('')
    target.push(textB64)
    if (htmlB64 !== undefined) {
      target.push(`--${boundaryAlt}`)
      target.push('Content-Type: text/html; charset=UTF-8')
      target.push('Content-Transfer-Encoding: base64')
      target.push('')
      target.push(htmlB64)
    }
    target.push(`--${boundaryAlt}--`)
  }
  if (!hasAttach && htmlB64 === undefined) {
    lines.push('Content-Type: text/plain; charset=UTF-8')
    lines.push('Content-Transfer-Encoding: base64')
    lines.push('')
    lines.push(textB64)
    return lines.join('\r\n')
  }
  if (!hasAttach) {
    lines.push(`Content-Type: multipart/alternative; boundary="${boundaryAlt}"`)
    lines.push('')
    pushAlt(lines)
    return lines.join('\r\n')
  }
  lines.push(`Content-Type: multipart/mixed; boundary="${boundaryMix}"`)
  lines.push('')
  lines.push(`--${boundaryMix}`)
  lines.push(`Content-Type: multipart/alternative; boundary="${boundaryAlt}"`)
  lines.push('')
  pushAlt(lines)
  for (const att of input.attachments ?? []) {
    const b64 = Buffer.from(att.bytes).toString('base64')
    const safeName = att.filename.replace(/["\r\n]/g, '_')
    lines.push(`--${boundaryMix}`)
    lines.push(`Content-Type: ${att.mime}; name="${safeName}"`)
    lines.push('Content-Transfer-Encoding: base64')
    lines.push(`Content-Disposition: attachment; filename="${safeName}"`)
    lines.push('')
    lines.push(b64)
  }
  lines.push(`--${boundaryMix}--`)
  return lines.join('\r\n')
}

/** Encode an RFC 5322 string for users.messages.send (base64url). */
export function encodeRawMessage(raw: string): string {
  return Buffer.from(raw, 'utf8')
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '')
}

export interface GmailDriverOptions {
  auth: GmailAuthProvider
  transport: GmailTransport
  /** Outbox id factory (injectable for deterministic tests). */
  makeOutboxId?: () => string
  /** Watch poll interval in ms; default 30 s. */
  pollMs?: number
  /** Account address used as the From header on send. */
  fromAddress?: string
}

/**
 * GmailDriver implementing MailProvider over an injected GmailTransport.
 * All network errors pass through mapGmailError so invalid_grant always
 * surfaces as GmailAuthExpiredError for the main-process vault policy.
 */
export class GmailDriver implements MailProvider {
  private readonly accountId: string
  private readonly auth: GmailAuthProvider
  private readonly transport: GmailTransport
  private readonly makeOutboxId: () => string
  private readonly pollMs: number
  private readonly fromAddress: string
  private connected = false
  private handlers = new Set<(ev: MailEvent) => void>()
  private timer: ReturnType<typeof setInterval> | undefined
  private hooks: GmailSyncHooks | undefined

  constructor(options: GmailDriverOptions) {
    this.auth = options.auth
    this.accountId = options.auth.accountId
    this.transport = options.transport
    this.makeOutboxId = options.makeOutboxId ?? (() => `gmail-outbox-${randomUUID()}`)
    this.pollMs = options.pollMs ?? 30000
    this.fromAddress = options.fromAddress ?? ''
  }

  /** Attach sync-state persistence (owned by main); required for watch(). */
  setSyncHooks(hooks: GmailSyncHooks): void {
    this.hooks = hooks
  }

  /** Touch the auth provider so missing credentials fail fast on connect. */
  async connect(): Promise<void> {
    await this.auth.getAccessToken()
    this.connected = true
  }

  async disconnect(): Promise<void> {
    this.stopWatch()
    this.connected = false
  }

  private ensureConnected(): void {
    if (!this.connected) throw new Error('gmail driver is not connected')
  }

  private emit(ev: MailEvent): void {
    for (const handler of this.handlers) handler(ev)
  }

  private async call<T>(fn: () => Promise<T>): Promise<T> {
    try {
      return await fn()
    } catch (err) {
      throw mapGmailError(err)
    }
  }

  async listFolders(): Promise<MailFolder[]> {
    this.ensureConnected()
    const labels = await this.call(() => this.transport.listLabels())
    const folders: MailFolder[] = []
    for (const label of labels) {
      const folder = gmailLabelToFolder(label)
      if (folder) folders.push(folder)
    }
    return folders
  }

  private async hydrateEnvelopes(
    ids: { id?: string | null; threadId?: string | null }[],
    folderId: string,
  ): Promise<MailEnvelope[]> {
    const envelopes: MailEnvelope[] = []
    for (const entry of ids) {
      if (!entry.id) continue
      const message = await this.call(() =>
        this.transport.getMessage(entry.id as string, 'metadata'),
      )
      const hdrs = message.payload?.headers ?? []
      envelopes.push(gmailMessageToEnvelope(message, folderId, hdrs))
    }
    return envelopes
  }

  async list(
    folderId: string,
    range: { afterUid?: string; limit: number },
  ): Promise<MailEnvelope[]> {
    this.ensureConnected()
    const page = await this.call(() =>
      this.transport.listMessages({
        labelIds: [folderId],
        maxResults: range.limit,
        pageToken: range.afterUid,
      }),
    )
    return this.hydrateEnvelopes(page.messages ?? [], folderId)
  }

  async search(folderId: string, query: MailQuery): Promise<MailEnvelope[]> {
    this.ensureConnected()
    const page = await this.call(() =>
      this.transport.listMessages({
        labelIds: [folderId],
        q: buildGmailQuery(query),
        maxResults: 50,
      }),
    )
    return this.hydrateEnvelopes(page.messages ?? [], folderId)
  }

  async fetch(ref: MailMessageRef): Promise<MailBody> {
    this.ensureConnected()
    const message = await this.call(() => this.transport.getMessage(ref.uid, 'full'))
    return { text: extractGmailBody(message).text }
  }

  /**
   * Raw fetch for the main layer: text plus HTML when the message is
   * multipart. Main persists the HTML content-addressed on disk and
   * keeps only the htmlRef handle in the store.
   */
  async fetchRaw(ref: MailMessageRef): Promise<GmailRawBody> {
    this.ensureConnected()
    const message = await this.call(() => this.transport.getMessage(ref.uid, 'full'))
    return extractGmailBody(message)
  }

  async send(req: SendRequest): Promise<{ outboxId: string }> {
    this.ensureConnected()
    const verdict = checkSendLimits(req)
    if (verdict === 'too-large') {
      throw new GmailSendLimitError(
        'too-large',
        `encoded size above the ${GMAIL_MAX_MESSAGE_BYTES} byte Gmail cap`,
      )
    }
    if (verdict === 'too-many-recipients') {
      throw new GmailSendLimitError(
        'too-many-recipients',
        `recipient count above the ${GMAIL_MAX_RECIPIENTS_PER_MESSAGE} per-message cap`,
      )
    }
    const attachments: RawAttachmentInput[] = (req.attachments ?? []).map((staged) => ({
      filename: staged.filename,
      mime: staged.mime,
      // Canonical payload policy: staged path wins; staged.data is the
      // test/tiny-inline fallback honored by every driver.
      bytes: staged.data ?? readFileSync(staged.path),
    }))
    const raw = buildRawMessage({
      from: this.fromAddress,
      to: req.to,
      cc: req.cc,
      bcc: req.bcc,
      subject: req.subject,
      text: req.text,
      html: req.html,
      attachments,
    })
    if (Buffer.byteLength(raw, 'utf8') > GMAIL_MAX_MESSAGE_BYTES) {
      throw new GmailSendLimitError(
        'too-large',
        'built message above the Gmail 35 MB cap after MIME framing',
      )
    }
    await this.call(() => this.transport.sendMessage(encodeRawMessage(raw)))
    return { outboxId: this.makeOutboxId() }
  }

  async move(ref: MailMessageRef, toFolderId: string): Promise<void> {
    this.ensureConnected()
    const remove = ref.folderId === toFolderId ? [] : [ref.folderId]
    await this.call(() => this.transport.modifyMessage(ref.uid, [toFolderId], remove))
  }

  async setFlags(ref: MailMessageRef, flags: Partial<MailFlags>): Promise<void> {
    this.ensureConnected()
    const add: string[] = []
    const remove: string[] = []
    if (flags.seen === true) remove.push('UNREAD')
    if (flags.seen === false) add.push('UNREAD')
    if (flags.flagged === true) add.push('STARRED')
    if (flags.flagged === false) remove.push('STARRED')
    // Gmail has no answered flag; answered/draft intent is recorded
    // nowhere server-side and is intentionally left untouched here.
    if (add.length === 0 && remove.length === 0) return
    await this.call(() => this.transport.modifyMessage(ref.uid, add, remove))
  }

  async delete(ref: MailMessageRef): Promise<void> {
    this.ensureConnected()
    await this.call(() => this.transport.trashMessage(ref.uid))
  }

  /** Full resync: re-list the inbox and persist the fresh historyId. */
  private async fullResync(reason: 'history-expired' | 'initial'): Promise<void> {
    this.emit({
      kind: 'sync-state',
      accountId: this.accountId,
      folderId: 'INBOX',
      state: reason === 'initial' ? 'syncing' : 'error',
      detail: reason === 'initial' ? undefined : 'history-expired',
    })
    const envelopes = await this.list('INBOX', { limit: 50 })
    this.emit({
      kind: 'messages-added',
      accountId: this.accountId,
      folderId: 'INBOX',
      refs: envelopes.map((env) => env.ref),
    })
    if (this.hooks) {
      const historyId = await this.call(() => this.transport.getCurrentHistoryId())
      this.hooks.save({
        accountId: this.accountId,
        folderId: 'INBOX',
        historyId,
        lastSyncAt: new Date().toISOString(),
      })
    }
    this.emit({ kind: 'sync-state', accountId: this.accountId, folderId: 'INBOX', state: 'idle' })
  }

  /**
   * One incremental sync step. Public so tests and the main scheduler can
   * drive it deterministically; watch() calls it on an interval.
   */
  async pollOnce(): Promise<void> {
    this.ensureConnected()
    if (!this.hooks) return
    const stored = this.hooks.load(this.accountId)
    if (!stored?.historyId) {
      await this.fullResync('initial')
      return
    }
    let history
    try {
      history = await this.call(() => this.transport.listHistory(stored.historyId as string))
    } catch (err) {
      if (isHistoryExpiredError(err)) {
        // Expired historyId means full resync. decideSync records the
        // reason; the driver still owns no outbox/draft/intent queue,
        // so those queues survive by construction.
        const decision = decideSync({ stored, historyExpired: true })
        if (decision.kind !== 'full-resync') throw err
        await this.fullResync('history-expired')
        return
      }
      throw err
    }
    for (const entry of history.added ?? []) {
      const message = entry.message
      if (!message?.id) continue
      this.emit({
        kind: 'messages-added',
        accountId: this.accountId,
        folderId: 'INBOX',
        refs: [{ uid: message.id, folderId: 'INBOX', threadId: message.threadId ?? message.id }],
      })
    }
    for (const entry of history.deleted ?? []) {
      const message = entry.message
      if (!message?.id) continue
      this.emit({
        kind: 'message-deleted',
        accountId: this.accountId,
        ref: { uid: message.id, folderId: 'INBOX', threadId: message.threadId ?? message.id },
      })
    }
    for (const entry of history.labelsChanged ?? []) {
      const message = entry.message
      if (!message?.id) continue
      this.emit({
        kind: 'flags-changed',
        accountId: this.accountId,
        ref: { uid: message.id, folderId: 'INBOX', threadId: message.threadId ?? message.id },
        flags: gmailLabelsToFlags(message.labelIds),
      })
    }
    this.hooks.save({
      ...stored,
      historyId: history.historyId,
      lastSyncAt: new Date().toISOString(),
    })
    this.emit({ kind: 'sync-state', accountId: this.accountId, folderId: 'INBOX', state: 'idle' })
  }

  watch(handler: (ev: MailEvent) => void): { stop(): void } {
    this.handlers.add(handler)
    if (this.timer === undefined) {
      void this.pollOnce().catch(() => undefined)
      this.timer = setInterval(() => {
        void this.pollOnce().catch(() => undefined)
      }, this.pollMs)
      if (typeof this.timer.unref === 'function') this.timer.unref()
    }
    return {
      stop: () => {
        this.handlers.delete(handler)
        if (this.handlers.size === 0) this.stopWatch()
      },
    }
  }

  private stopWatch(): void {
    if (this.timer !== undefined) clearInterval(this.timer)
    this.timer = undefined
  }
}

/**
 * Parse a sent raw message back to a view-model for the round-trip
 * matrix row. Reuses the Lane A MIME parser so compose output and
 * reader input share one interpretation.
 */
export async function parseSentRaw(raw: string): Promise<{ subject: string; text: string }> {
  const parsed = await parseMimeMessage(raw)
  return { subject: parsed.subject, text: parsed.text }
}
