/**
 * Production transports for ImapSmtpDriver.
 *
 * - ImapFlowAdapter wraps imapflow 1.5.x (MIT).
 * - NodemailerSmtpAdapter wraps nodemailer >= 4 (MIT-0) with attachments
 *   streamed from disk paths at send time.
 *
 * Both take injected credentials from the main-process vault caller; neither
 * touches safeStorage (R2). These adapters are deliberately thin: retry,
 * resync, and intent policy lives in the driver, which the fakes cover.
 */
import { ImapFlow } from 'imapflow'
import type {
  ExpungeEvent,
  FetchMessageObject,
  FetchQueryObject,
  FlagsEvent,
  ListResponse,
  MessageAddressObject,
  MessageStructureObject,
  SearchObject,
  StoreOptions,
} from 'imapflow'
import nodemailer from 'nodemailer'
import type { Transporter } from 'nodemailer'
import {
  StaleModseqError,
  type IdleMailboxEvent,
  type ImapClientLike,
  type ImapConnectionOptions,
  type ImapEnvelopeData,
  type ImapMailboxInfo,
  type ImapMailboxStatus,
  type SmtpConnectionOptions,
  type SmtpSendMail,
  type SmtpTransportLike,
} from './imap-smtp-driver.js'
import type { MailFolder, MailQuery } from './types.js'

const ENVELOPE_QUERY: FetchQueryObject = {
  uid: true,
  envelope: true,
  flags: true,
  internalDate: true,
  bodyStructure: true,
  bodyParts: [{ key: 'TEXT', maxLength: 280 }],
}

function addresses(addrs: MessageAddressObject[] | undefined): string[] {
  return (addrs ?? []).map((a) => a.address ?? a.name ?? '').filter((a) => a.length > 0)
}

function dateOf(msg: FetchMessageObject): string {
  const headerDate = msg.envelope?.date
  if (headerDate instanceof Date && !Number.isNaN(headerDate.getTime())) {
    return headerDate.toISOString()
  }
  if (msg.internalDate !== undefined) {
    const internal = new Date(msg.internalDate)
    if (!Number.isNaN(internal.getTime())) return internal.toISOString()
  }
  return new Date(0).toISOString()
}

function structureHasAttachment(node: MessageStructureObject): boolean {
  if (node.disposition === 'attachment') return true
  if (node.dispositionParameters?.['filename'] || node.parameters?.['name']) return true
  if (node.type === 'multipart') {
    return (node.childNodes ?? []).some(structureHasAttachment)
  }
  return node.type !== 'text' && node.type !== 'multipart'
}

function snippetOf(msg: FetchMessageObject): string {
  const parts = msg.bodyParts ? [...msg.bodyParts.values()] : []
  const first = parts[0]?.toString('utf8') ?? ''
  const collapsed = first.replace(/\s+/g, ' ').trim()
  if (collapsed.length > 0) return collapsed.slice(0, 160)
  return (msg.envelope?.subject ?? '').slice(0, 160)
}

function toData(msg: FetchMessageObject): ImapEnvelopeData {
  const flags = msg.flags ?? new Set<string>()
  return {
    uid: String(msg.uid),
    modseq: msg.modseq?.toString(),
    from: addresses(msg.envelope?.from),
    to: addresses(msg.envelope?.to),
    cc: addresses(msg.envelope?.cc),
    subject: msg.envelope?.subject ?? '',
    date: dateOf(msg),
    seen: flags.has('\\Seen'),
    flagged: flags.has('\\Flagged'),
    answered: flags.has('\\Answered'),
    draft: flags.has('\\Draft'),
    msgIdHeader: msg.envelope?.messageId,
    labels: msg.labels ? [...msg.labels] : [],
    snippet: snippetOf(msg),
    hasAttach: msg.bodyStructure ? structureHasAttachment(msg.bodyStructure) : false,
  }
}

function roleFromBox(box: ListResponse): MailFolder['role'] {
  const special = box.specialUse ?? ''
  if (special === '\\Sent') return 'sent'
  if (special === '\\Drafts') return 'drafts'
  if (special === '\\Trash') return 'trash'
  if (special === '\\Junk') return 'spam'
  if (special === '\\Archive') return 'archive'
  const upper = box.path.toUpperCase()
  const leaf = upper.split(box.delimiter).pop() ?? upper
  if (leaf === 'INBOX' || special === '\\Inbox') return 'inbox'
  if (leaf === 'SENT') return 'sent'
  if (leaf === 'DRAFTS') return 'drafts'
  if (leaf === 'TRASH') return 'trash'
  if (leaf === 'JUNK' || leaf === 'SPAM') return 'spam'
  if (leaf === 'ARCHIVE') return 'archive'
  return 'custom'
}

export class ImapFlowAdapter implements ImapClientLike {
  private readonly client: ImapFlow
  private selectedPath: string | undefined

  constructor(private readonly options: ImapConnectionOptions) {
    this.client = new ImapFlow({
      host: options.host,
      port: options.port,
      secure: options.secure,
      logger: false,
      auth: { user: options.auth.user, pass: options.auth.password },
    })
  }

  async connect(): Promise<void> {
    await this.client.connect()
  }

  async close(): Promise<void> {
    try {
      await this.client.logout()
    } catch {
      this.client.close()
    }
    this.selectedPath = undefined
  }

  supports(capability: string): boolean {
    return this.client.capabilities.has(capability.toUpperCase())
  }

  async listMailboxes(): Promise<ImapMailboxInfo[]> {
    const listed = await this.client.list({ statusQuery: { unseen: true } })
    return listed
      .filter((box) => !box.flags.has('\\Noselect'))
      .map((box) => ({
        path: box.path,
        name: box.name,
        role: roleFromBox(box),
        unread: box.status?.unseen ?? 0,
      }))
  }

  async selectMailbox(path: string): Promise<ImapMailboxStatus> {
    const box = await this.client.mailboxOpen(path)
    this.selectedPath = path
    return {
      uidValidity: Number(box.uidValidity),
      highestModseq: box.highestModseq?.toString(),
      exists: box.exists,
    }
  }

  private async ensureSelected(path: string): Promise<void> {
    if (this.selectedPath !== path) {
      await this.client.mailboxOpen(path)
      this.selectedPath = path
    }
  }

  async fetchRange(
    path: string,
    afterUid: string | undefined,
    limit: number,
  ): Promise<ImapEnvelopeData[]> {
    await this.ensureSelected(path)
    const start = afterUid === undefined ? 1 : Number(afterUid) + 1
    const rows: ImapEnvelopeData[] = []
    for await (const msg of this.client.fetch(`${start}:*`, { ...ENVELOPE_QUERY }, { uid: true })) {
      rows.push(toData(msg))
      if (rows.length >= limit) break
    }
    return rows
  }

  async fetchSinceModseq(path: string, modseq: string): Promise<ImapEnvelopeData[]> {
    await this.ensureSelected(path)
    const rows: ImapEnvelopeData[] = []
    const it = this.client.fetch(
      '1:*',
      { ...ENVELOPE_QUERY },
      {
        uid: true,
        changedSince: BigInt(modseq),
      },
    )
    for await (const msg of it) rows.push(toData(msg))
    return rows
  }

  async fetchBody(path: string, uid: string): Promise<{ text: string }> {
    await this.ensureSelected(path)
    const msg = await this.client.fetchOne(Number(uid), { bodyParts: ['TEXT'] }, { uid: true })
    if (!msg) throw new Error(`No message ${uid} in ${path}`)
    const parts = msg.bodyParts ? [...msg.bodyParts.values()] : []
    // Full MIME decode (base64/quoted-printable charsets) is Lane A's
    // postal-mime path; here the raw text part is decoded as UTF-8.
    return { text: parts[0]?.toString('utf8') ?? '' }
  }

  async setFlags(
    path: string,
    uid: string,
    flags: { seen?: boolean; flagged?: boolean; answered?: boolean; draft?: boolean },
    options?: { expectedModseq?: string },
  ): Promise<string | undefined> {
    await this.ensureSelected(path)
    const storeOptions: StoreOptions = { uid: true }
    if (this.supports('CONDSTORE') && options?.expectedModseq !== undefined) {
      storeOptions.unchangedSince = BigInt(options.expectedModseq)
    }
    const add: string[] = []
    const remove: string[] = []
    const put = (value: boolean | undefined, flag: string): void => {
      if (value === true) add.push(flag)
      else if (value === false) remove.push(flag)
    }
    put(flags.seen, '\\Seen')
    put(flags.flagged, '\\Flagged')
    put(flags.answered, '\\Answered')
    put(flags.draft, '\\Draft')
    if (add.length > 0) {
      const ok = await this.client.messageFlagsAdd(Number(uid), add, storeOptions)
      if (!ok) throw new StaleModseqError(`Flags for ${uid} in ${path} are stale`)
    }
    if (remove.length > 0) {
      const ok = await this.client.messageFlagsRemove(Number(uid), remove, storeOptions)
      if (!ok) throw new StaleModseqError(`Flags for ${uid} in ${path} are stale`)
    }
    if (!this.supports('CONDSTORE')) return undefined
    const after = await this.client.fetchOne(Number(uid), {}, { uid: true })
    return after ? after.modseq?.toString() : undefined
  }

  async moveMessage(
    path: string,
    uid: string,
    destPath: string,
    options?: { expectedModseq?: string },
  ): Promise<void> {
    await this.ensureSelected(path)
    if (this.supports('CONDSTORE') && options?.expectedModseq !== undefined) {
      const current = await this.client.fetchOne(Number(uid), {}, { uid: true })
      if (!current) throw new Error(`No message ${uid} in ${path}`)
      if (current.modseq?.toString() !== options.expectedModseq) {
        throw new StaleModseqError(`Message ${uid} in ${path} changed on the server`)
      }
    }
    const moved = await this.client.messageMove(Number(uid), destPath, { uid: true })
    if (!moved) throw new Error(`Move of ${uid} from ${path} to ${destPath} failed`)
  }

  async deleteMessage(path: string, uid: string): Promise<void> {
    await this.ensureSelected(path)
    // Sets \Deleted; physical removal follows the server's CLOSE semantics.
    await this.client.messageDelete(Number(uid), { uid: true })
  }

  async searchMessages(path: string, query: MailQuery): Promise<ImapEnvelopeData[]> {
    await this.ensureSelected(path)
    const search: SearchObject = {}
    if (query.text) search.text = query.text
    if (query.from) search.from = query.from
    if (query.to) search.to = query.to
    if (query.subject) search.subject = query.subject
    if (query.seen !== undefined) search.seen = query.seen
    if (query.flagged !== undefined) search.flagged = query.flagged
    if (query.since) search.since = query.since
    if (query.before) search.before = query.before
    if (Object.keys(search).length === 0) search.all = true
    const uids = await this.client.search(search, { uid: true })
    if (!uids || uids.length === 0) return []
    const rows: ImapEnvelopeData[] = []
    const it = this.client.fetch(uids, { ...ENVELOPE_QUERY }, { uid: true })
    for await (const msg of it) rows.push(toData(msg))
    const found =
      query.hasAttach === undefined ? rows : rows.filter((r) => r.hasAttach === query.hasAttach)
    return found
  }

  idle(onEvent: (ev: IdleMailboxEvent) => void): { stop(): void } {
    let stopped = false
    const path = this.selectedPath ?? 'INBOX'

    const onExists = async (): Promise<void> => {
      if (stopped) return
      try {
        await this.ensureSelected(path)
        const status = await this.client.status(path, { uidNext: true })
        const uidNext = status?.uidNext
        if (uidNext === undefined) return
        const rows: ImapEnvelopeData[] = []
        const it = this.client.fetch(`${uidNext - 1}:*`, { uid: true }, { uid: true })
        for await (const msg of it) rows.push(toData(msg))
        for (const row of rows) {
          if (Number(row.uid) >= uidNext - 1 && !stopped) {
            onEvent({ path, uid: row.uid, kind: 'new' })
          }
        }
      } catch {
        // Transient races (e.g. close during fetch) end with the idle loop.
      }
    }
    const onExpunge = (ev: ExpungeEvent): void => {
      if (!stopped && ev.uid !== undefined) {
        onEvent({ path: ev.path, uid: String(ev.uid), kind: 'deleted' })
      }
    }
    const onFlags = (ev: FlagsEvent): void => {
      if (!stopped && ev.uid !== undefined) {
        onEvent({ path: ev.path, uid: String(ev.uid), kind: 'updated' })
      }
    }
    const existsListener = (): void => {
      void onExists()
    }

    void (async () => {
      try {
        await this.ensureSelected(path)
      } catch {
        return
      }
      this.client.on('exists', existsListener)
      this.client.on('expunge', onExpunge)
      this.client.on('flags', onFlags)
      try {
        while (!stopped) {
          await this.client.idle()
        }
      } catch {
        // Connection closed or IDLE unsupported mid-loop; watchers stop.
      } finally {
        this.client.removeListener('exists', existsListener)
        this.client.removeListener('expunge', onExpunge)
        this.client.removeListener('flags', onFlags)
      }
    })()

    return {
      stop: () => {
        stopped = true
        // Any command breaks IDLE; the loop then exits and cleans up.
        void this.client.noop().catch(() => undefined)
      },
    }
  }
}

export class NodemailerSmtpAdapter implements SmtpTransportLike {
  private readonly transporter: Transporter

  constructor(
    private readonly options: SmtpConnectionOptions,
    transportOverride?: Transporter,
  ) {
    this.transporter =
      transportOverride ??
      nodemailer.createTransport({
        host: options.host,
        port: options.port,
        secure: options.secure,
        auth: { user: options.auth.user, pass: options.auth.password },
      })
  }

  async send(mail: SmtpSendMail): Promise<{ messageId: string }> {
    const info = await this.transporter.sendMail({
      from: mail.from,
      to: mail.to,
      cc: mail.cc,
      bcc: mail.bcc,
      subject: mail.subject,
      text: mail.text,
      html: mail.html,
      attachments: mail.attachments.map((a) =>
        a.path
          ? { filename: a.filename, contentType: a.contentType, path: a.path }
          : {
              filename: a.filename,
              contentType: a.contentType,
              content: Buffer.from(a.content ?? new Uint8Array()),
            },
      ),
    })
    return { messageId: info.messageId }
  }

  async close(): Promise<void> {
    this.transporter.close()
  }
}

export function createImapFlowClient(options: ImapConnectionOptions): ImapClientLike {
  return new ImapFlowAdapter(options)
}

export function createNodemailerTransport(options: SmtpConnectionOptions): SmtpTransportLike {
  return new NodemailerSmtpAdapter(options)
}
