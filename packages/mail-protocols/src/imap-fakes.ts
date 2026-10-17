/**
 * In-memory IMAP/SMTP doubles for Lane B tests and local staging.
 *
 * These fakes implement the same narrow seams as the production adapters
 * (ImapClientLike / SmtpTransportLike) so the round-trip matrix row,
 * UIDVALIDITY resync, and offline-intent fixtures run with no network and
 * no real mailbox credentials. They are exported for the P3 sync-engine
 * lane, which can reuse them for its own fixtures.
 */
import {
  IdleMailboxEvent,
  ImapClientLike,
  ImapEnvelopeData,
  ImapMailboxInfo,
  ImapMailboxStatus,
  SmtpSendMail,
  SmtpTransportLike,
  StaleModseqError,
} from './imap-smtp-driver.js'
import type { MailFolder, MailQuery } from './types.js'

export interface FakeMessageInit {
  from: string[]
  to: string[]
  cc?: string[]
  subject: string
  date?: string
  seen?: boolean
  flagged?: boolean
  answered?: boolean
  draft?: boolean
  msgIdHeader?: string
  labels?: string[]
  snippet?: string
  hasAttach?: boolean
  text?: string
}

interface FakeStoredMessage extends FakeMessageInit {
  uid: string
  modseq: number
}

interface FakeMailbox {
  uidValidity: number
  messages: FakeStoredMessage[]
  nextUid: number
  modseq: number
  movedFrom: Map<string, string>
}

export interface FakeImapOptions {
  idle?: boolean
  condstore?: boolean
}

function roleForPath(path: string): MailFolder['role'] {
  const upper = path.toUpperCase()
  const leaf = upper.split('/').pop() ?? upper
  if (leaf === 'INBOX') return 'inbox'
  if (leaf === 'SENT') return 'sent'
  if (leaf === 'DRAFTS') return 'drafts'
  if (leaf === 'TRASH') return 'trash'
  if (leaf === 'JUNK' || leaf === 'SPAM') return 'spam'
  if (leaf === 'ARCHIVE') return 'archive'
  return 'custom'
}

export class FakeImapServer implements ImapClientLike {
  private readonly capabilities = new Set<string>()
  private readonly mailboxes = new Map<string, FakeMailbox>()
  private readonly idleSubs = new Set<(ev: IdleMailboxEvent) => void>()
  private connected = false

  /** Test hooks. */
  failOnSelect?: Error
  failOnList?: Error
  /** UIDs that raise StaleModseqError exactly once, then behave normally. */
  readonly staleOnce = new Set<string>()
  /** When true, every flagged/move op raises StaleModseqError (conflict path). */
  alwaysStale = false

  /** Call counters so tests can assert CONDSTORE vs UID-range paths. */
  fetchRangeCalls = 0
  fetchSinceModseqCalls = 0
  selectCalls = 0

  constructor(options?: FakeImapOptions) {
    if (options?.idle !== false) this.capabilities.add('IDLE')
    if (options?.condstore !== false) this.capabilities.add('CONDSTORE')
  }

  setCapability(name: string, on: boolean): void {
    if (on) this.capabilities.add(name.toUpperCase())
    else this.capabilities.delete(name.toUpperCase())
  }

  async connect(): Promise<void> {
    this.connected = true
  }

  async close(): Promise<void> {
    this.connected = false
    this.idleSubs.clear()
  }

  supports(capability: string): boolean {
    return this.capabilities.has(capability.toUpperCase())
  }

  private mailbox(path: string): FakeMailbox {
    let box = this.mailboxes.get(path)
    if (!box) {
      box = { uidValidity: 1, messages: [], nextUid: 1, modseq: 0, movedFrom: new Map() }
      this.mailboxes.set(path, box)
    }
    return box
  }

  seedMessage(path: string, init: FakeMessageInit): string {
    const box = this.mailbox(path)
    const uid = String(box.nextUid++)
    box.modseq += 1
    box.messages.push({
      from: init.from,
      to: init.to,
      cc: init.cc ?? [],
      subject: init.subject,
      date: init.date ?? new Date().toISOString(),
      seen: init.seen ?? false,
      flagged: init.flagged ?? false,
      answered: init.answered ?? false,
      draft: init.draft ?? false,
      msgIdHeader: init.msgIdHeader,
      labels: init.labels ?? [],
      snippet: init.snippet ?? init.subject,
      hasAttach: init.hasAttach ?? false,
      text: init.text ?? init.subject,
      uid,
      modseq: box.modseq,
    })
    return uid
  }

  /** Seeds a message and notifies IDLE subscribers, like a real delivery. */
  simulateIncoming(path: string, init: FakeMessageInit): string {
    const uid = this.seedMessage(path, init)
    for (const sub of this.idleSubs) sub({ path, uid, kind: 'new' })
    return uid
  }

  simulateFlagChange(path: string, uid: string): void {
    for (const sub of this.idleSubs) sub({ path, uid, kind: 'updated' })
  }

  setUidValidity(path: string, uidValidity: number): void {
    this.mailbox(path).uidValidity = uidValidity
  }

  touchMessage(path: string, uid: string, patch: Partial<FakeMessageInit>): void {
    const box = this.mailbox(path)
    const found = box.messages.find((m) => m.uid === uid)
    if (!found) throw new Error(`FakeImapServer: no message ${uid} in ${path}`)
    Object.assign(found, patch)
    box.modseq += 1
    found.modseq = box.modseq
  }

  messageModseq(path: string, uid: string): string | undefined {
    const box = this.mailbox(path)
    return box.messages.find((m) => m.uid === uid)?.modseq.toString()
  }

  messageCount(path: string): number {
    return this.mailbox(path).messages.length
  }

  async listMailboxes(): Promise<ImapMailboxInfo[]> {
    if (this.failOnList) throw this.failOnList
    return [...this.mailboxes.keys()].map((path) => {
      const box = this.mailboxes.get(path)
      const unread = box?.messages.filter((m) => !m.seen).length ?? 0
      const name = path.split('/').pop() ?? path
      return { path, name, role: roleForPath(path), unread }
    })
  }

  async selectMailbox(path: string): Promise<ImapMailboxStatus> {
    this.selectCalls += 1
    if (this.failOnSelect) throw this.failOnSelect
    const box = this.mailbox(path)
    return {
      uidValidity: box.uidValidity,
      highestModseq: this.supports('CONDSTORE') ? String(box.modseq) : undefined,
      exists: box.messages.length,
    }
  }

  async fetchRange(
    path: string,
    afterUid: string | undefined,
    limit: number,
  ): Promise<ImapEnvelopeData[]> {
    this.fetchRangeCalls += 1
    const box = this.mailbox(path)
    const after = afterUid === undefined ? 0 : Number(afterUid)
    return box.messages
      .filter((m) => Number(m.uid) > after)
      .slice(0, limit)
      .map((m) => toData(m, this.supports('CONDSTORE')))
  }

  async fetchSinceModseq(path: string, modseq: string): Promise<ImapEnvelopeData[]> {
    this.fetchSinceModseqCalls += 1
    const box = this.mailbox(path)
    return box.messages.filter((m) => m.modseq > Number(modseq)).map((m) => toData(m, true))
  }

  async fetchBody(path: string, uid: string): Promise<{ text: string }> {
    const found = this.mailbox(path).messages.find((m) => m.uid === uid)
    if (!found) throw new Error(`FakeImapServer: no message ${uid} in ${path}`)
    return { text: found.text ?? found.subject }
  }

  async setFlags(
    path: string,
    uid: string,
    flags: { seen?: boolean; flagged?: boolean; answered?: boolean; draft?: boolean },
    _options?: { expectedModseq?: string },
  ): Promise<string | undefined> {
    // The fake signals staleness through explicit test hooks (staleOnce /
    // alwaysStale), not through the expected value, so replays stay
    // deterministic. The production adapter honors expectedModseq via
    // CONDSTORE UNCHANGEDSINCE.
    this.throwIfStale(uid)
    const box = this.mailbox(path)
    const found = box.messages.find((m) => m.uid === uid)
    if (!found) return undefined
    if (flags.seen !== undefined) found.seen = flags.seen
    if (flags.flagged !== undefined) found.flagged = flags.flagged
    if (flags.answered !== undefined) found.answered = flags.answered
    if (flags.draft !== undefined) found.draft = flags.draft
    box.modseq += 1
    found.modseq = box.modseq
    this.simulateFlagChange(path, uid)
    return this.supports('CONDSTORE') ? String(found.modseq) : undefined
  }

  async moveMessage(
    path: string,
    uid: string,
    destPath: string,
    _options?: { expectedModseq?: string },
  ): Promise<void> {
    this.throwIfStale(uid)
    const src = this.mailbox(path)
    const index = src.messages.findIndex((m) => m.uid === uid)
    if (index === -1) {
      // Idempotent replay: already moved to the same destination.
      if (src.movedFrom.get(uid) === destPath) return
      throw new Error(`FakeImapServer: no message ${uid} in ${path}`)
    }
    const [moved] = src.messages.splice(index, 1)
    if (!moved) throw new Error(`FakeImapServer: no message ${uid} in ${path}`)
    const dest = this.mailbox(destPath)
    const newUid = String(dest.nextUid++)
    dest.modseq += 1
    dest.messages.push({ ...moved, uid: newUid, modseq: dest.modseq })
    src.movedFrom.set(uid, destPath)
    for (const sub of this.idleSubs) sub({ path, uid, kind: 'deleted' })
  }

  async deleteMessage(path: string, uid: string): Promise<void> {
    const box = this.mailbox(path)
    const index = box.messages.findIndex((m) => m.uid === uid)
    // Deleting a missing message is a no-op (idempotent replay).
    if (index === -1) return
    box.messages.splice(index, 1)
    for (const sub of this.idleSubs) sub({ path, uid, kind: 'deleted' })
  }

  async searchMessages(path: string, query: MailQuery): Promise<ImapEnvelopeData[]> {
    const box = this.mailbox(path)
    const needle = query.text?.toLowerCase()
    const fromFilter = query.from?.toLowerCase()
    const toFilter = query.to?.toLowerCase()
    const subjectFilter = query.subject?.toLowerCase()
    const sinceDay = query.since ? dayStart(query.since) : undefined
    const beforeDay = query.before ? dayStart(query.before) : undefined
    return box.messages
      .filter((m) => {
        if (needle) {
          const hay = `${m.subject} ${m.snippet} ${m.from.join(' ')}`.toLowerCase()
          if (!hay.includes(needle)) return false
        }
        if (fromFilter && !m.from.some((a) => a.toLowerCase().includes(fromFilter))) return false
        if (toFilter && !m.to.some((a) => a.toLowerCase().includes(toFilter))) return false
        if (subjectFilter && !m.subject.toLowerCase().includes(subjectFilter)) return false
        if (query.seen !== undefined && m.seen !== query.seen) return false
        if (query.flagged !== undefined && m.flagged !== query.flagged) return false
        if (query.hasAttach !== undefined && m.hasAttach !== query.hasAttach) return false
        if (sinceDay !== undefined) {
          const date = m.date ? dayStart(m.date) : undefined
          if (date === undefined || date < sinceDay) return false
        }
        if (beforeDay !== undefined) {
          const date = m.date ? dayStart(m.date) : undefined
          if (date === undefined || date >= beforeDay) return false
        }
        return true
      })
      .map((m) => toData(m, this.supports('CONDSTORE')))
  }

  idle(onEvent: (ev: IdleMailboxEvent) => void): { stop(): void } {
    this.idleSubs.add(onEvent)
    return { stop: () => void this.idleSubs.delete(onEvent) }
  }

  private throwIfStale(uid: string): void {
    if (this.alwaysStale) {
      throw new StaleModseqError(`FakeImapServer: modseq for ${uid} is stale`)
    }
    if (this.staleOnce.has(uid)) {
      this.staleOnce.delete(uid)
      throw new StaleModseqError(`FakeImapServer: modseq for ${uid} is stale`)
    }
  }
}

/** Day-granularity floor for SINCE/BEFORE comparisons; undefined when unparsable. */
function dayStart(value: string): number | undefined {
  const time = new Date(value).getTime()
  if (Number.isNaN(time)) return undefined
  const day = new Date(time)
  day.setHours(0, 0, 0, 0)
  return day.getTime()
}

function toData(m: FakeStoredMessage, withModseq: boolean): ImapEnvelopeData {
  return {
    uid: m.uid,
    modseq: withModseq ? String(m.modseq) : undefined,
    from: m.from,
    to: m.to,
    cc: m.cc ?? [],
    subject: m.subject,
    date: m.date ?? new Date(0).toISOString(),
    seen: m.seen ?? false,
    flagged: m.flagged ?? false,
    answered: m.answered ?? false,
    draft: m.draft ?? false,
    msgIdHeader: m.msgIdHeader,
    labels: m.labels ?? [],
    snippet: m.snippet ?? m.subject,
    hasAttach: m.hasAttach ?? false,
  }
}

export class FakeSmtpServer implements SmtpTransportLike {
  readonly sent: SmtpSendMail[] = []
  failNext?: Error
  private seq = 0

  async send(mail: SmtpSendMail): Promise<{ messageId: string }> {
    if (this.failNext) {
      const err = this.failNext
      this.failNext = undefined
      throw err
    }
    this.sent.push(structuredClone(mail))
    this.seq += 1
    return { messageId: `<fake-${this.seq}@genoffice.local>` }
  }

  async close(): Promise<void> {}

  /** Delivers a captured send into a fake IMAP mailbox, closing the test loop. */
  deliverTo(imap: FakeImapServer, path: string, index = -1): string {
    const captured = this.sent.at(index)
    if (!captured) throw new Error('FakeSmtpServer: nothing sent yet')
    return imap.seedMessage(path, {
      from: [captured.from],
      to: captured.to,
      cc: captured.cc,
      subject: captured.subject,
      msgIdHeader: `<fake-${this.seq}@genoffice.local>`,
      snippet: captured.text.slice(0, 120),
      hasAttach: captured.attachments.length > 0,
      text: captured.text,
    })
  }
}
