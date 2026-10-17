/**
 * In-memory Gmail API fake for driver tests. No network, no credentials:
 * the auth provider returns a constant fake token and every transport call
 * runs against seeded fixtures.
 */

import type {
  GmailAuthProvider,
  GmailLabel,
  GmailMessage,
  GmailPayload,
  GmailSyncHooks,
  GmailTransport,
} from '../src/gmail'
import type { SyncState } from '../src/types'

export function b64urlText(text: string): string {
  return Buffer.from(text, 'utf8')
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '')
}

export interface SeedMessage {
  id: string
  threadId: string
  labelIds: string[]
  headers: Record<string, string>
  text: string
  html?: string
  attachments?: { filename: string; mime: string; text: string }[]
}

function seedPayload(seed: SeedMessage, full: boolean): GmailPayload {
  const headers = Object.entries(seed.headers).map(([name, value]) => ({ name, value }))
  if (!full) {
    // Metadata format keeps headers plus the part skeleton (filenames,
    // no body data), matching the real Gmail API metadata shape.
    return {
      mimeType: 'multipart/mixed',
      filename: '',
      headers,
      parts: (seed.attachments ?? []).map((att) => ({
        mimeType: att.mime,
        filename: att.filename,
        headers: [],
        parts: [],
      })),
    }
  }
  const parts: GmailPayload[] = [
    {
      mimeType: 'text/plain',
      filename: '',
      headers: [],
      body: { data: b64urlText(seed.text), size: seed.text.length },
      parts: [],
    },
  ]
  if (seed.html !== undefined) {
    parts.push({
      mimeType: 'text/html',
      filename: '',
      headers: [],
      body: { data: b64urlText(seed.html), size: seed.html.length },
      parts: [],
    })
  }
  for (const att of seed.attachments ?? []) {
    parts.push({
      mimeType: att.mime,
      filename: att.filename,
      headers: [],
      body: { data: b64urlText(att.text), size: att.text.length },
      parts: [],
    })
  }
  return {
    mimeType: 'multipart/mixed',
    filename: '',
    headers: Object.entries(seed.headers).map(([name, value]) => ({ name, value })),
    parts,
  }
}

export function fakeAuth(accountId = 'gmail-test'): GmailAuthProvider {
  return {
    accountId,
    getAccessToken: async () => 'fake-token',
  }
}

export function memorySyncHooks(): GmailSyncHooks & { states: Map<string, SyncState> } {
  const states = new Map<string, SyncState>()
  return {
    states,
    load: (accountId: string) => states.get(accountId),
    save: (state: SyncState) => {
      states.set(state.accountId, state)
    },
  }
}

export class FakeGmailTransport implements GmailTransport {
  seeds = new Map<string, SeedMessage>()
  sentRaw: string[] = []
  modified: { id: string; add: string[]; remove: string[] }[] = []
  trashed: string[] = []
  historyId = '9000'
  historyError: unknown = undefined
  stagedHistory: { added: GmailMessage[]; deleted: GmailMessage[]; labelsChanged: GmailMessage[] } =
    {
      added: [],
      deleted: [],
      labelsChanged: [],
    }

  constructor(seeds: SeedMessage[] = []) {
    for (const seed of seeds) this.seeds.set(seed.id, seed)
  }

  async listLabels(): Promise<GmailLabel[]> {
    return [
      { id: 'INBOX', name: 'INBOX', messagesUnread: 2, type: 'system' },
      { id: 'SENT', name: 'SENT', messagesUnread: 0, type: 'system' },
      { id: 'DRAFT', name: 'DRAFT', messagesUnread: 0, type: 'system' },
      { id: 'TRASH', name: 'TRASH', messagesUnread: 0, type: 'system' },
      { id: 'SPAM', name: 'SPAM', messagesUnread: 0, type: 'system' },
      { id: 'UNREAD', name: 'UNREAD', messagesUnread: 2, type: 'system' },
      { id: 'IMPORTANT', name: 'IMPORTANT', messagesUnread: 0, type: 'system' },
      { id: 'Label_1', name: 'Work', messagesUnread: 1, type: 'user' },
    ]
  }

  private matchesQuery(seed: SeedMessage, q: string): boolean {
    const hay = `${seed.headers.Subject ?? ''} ${seed.text} ${(seed.headers.From ?? '').toLowerCase()}`
    const tokens = q.match(/from:\S+|is:unread|after:\S+|\S+/g) ?? []
    for (const token of tokens) {
      if (token.startsWith('from:')) {
        const want = token.slice('from:'.length).toLowerCase()
        if (!(seed.headers.From ?? '').toLowerCase().includes(want)) return false
      } else if (token === 'is:unread') {
        if (!seed.labelIds.includes('UNREAD')) return false
      } else if (token.startsWith('after:')) {
        continue
      } else if (!hay.toLowerCase().includes(token.toLowerCase())) {
        return false
      }
    }
    return true
  }

  async listMessages(args: {
    labelIds?: string[]
    q?: string
    maxResults: number
    pageToken?: string
  }): Promise<{
    messages?: { id?: string | null; threadId?: string | null }[]
    nextPageToken?: string | null
  }> {
    let all = [...this.seeds.values()]
    if (args.labelIds && args.labelIds.length > 0) {
      all = all.filter((seed) => args.labelIds?.every((label) => seed.labelIds.includes(label)))
    }
    if (args.q && args.q.trim() !== '')
      all = all.filter((seed) => this.matchesQuery(seed, args.q as string))
    const start = args.pageToken ? Number.parseInt(args.pageToken, 10) || 0 : 0
    const slice = all.slice(start, start + args.maxResults)
    const next = start + args.maxResults < all.length ? String(start + args.maxResults) : undefined
    return {
      messages: slice.map((seed) => ({ id: seed.id, threadId: seed.threadId })),
      nextPageToken: next,
    }
  }

  private toMessage(seed: SeedMessage, full: boolean): GmailMessage {
    return {
      id: seed.id,
      threadId: seed.threadId,
      labelIds: [...seed.labelIds],
      snippet: seed.text.slice(0, 100),
      historyId: this.historyId,
      sizeEstimate: seed.text.length,
      payload: seedPayload(seed, full),
    }
  }

  async getMessage(id: string, format: 'metadata' | 'full'): Promise<GmailMessage> {
    const seed = this.seeds.get(id)
    if (!seed) throw { code: 404, message: 'Requested entity was not found.' }
    return this.toMessage(seed, format === 'full')
  }

  async sendMessage(rawBase64Url: string): Promise<{ id?: string | null }> {
    this.sentRaw.push(rawBase64Url)
    return { id: `sent-${this.sentRaw.length}` }
  }

  async modifyMessage(id: string, addLabelIds: string[], removeLabelIds: string[]): Promise<void> {
    const seed = this.seeds.get(id)
    if (!seed) throw { code: 404, message: 'Requested entity was not found.' }
    this.modified.push({ id, add: addLabelIds, remove: removeLabelIds })
    const labels = new Set(seed.labelIds)
    for (const label of addLabelIds) labels.add(label)
    for (const label of removeLabelIds) labels.delete(label)
    seed.labelIds = [...labels]
  }

  async trashMessage(id: string): Promise<void> {
    const seed = this.seeds.get(id)
    if (!seed) throw { code: 404, message: 'Requested entity was not found.' }
    this.trashed.push(id)
    const labels = new Set(seed.labelIds)
    labels.delete('INBOX')
    labels.add('TRASH')
    seed.labelIds = [...labels]
  }

  async listHistory(startHistoryId: string): Promise<{
    historyId: string
    added?: { message?: GmailMessage | null }[]
    deleted?: { message?: GmailMessage | null }[]
    labelsChanged?: { message?: GmailMessage | null }[]
  }> {
    if (this.historyError !== undefined) throw this.historyError
    if (startHistoryId !== this.historyId) {
      throw { code: 404, message: 'Requested entity was not found.' }
    }
    const slim = (m: GmailMessage) => ({ message: m })
    return {
      historyId: this.historyId,
      added: this.stagedHistory.added.map(slim),
      deleted: this.stagedHistory.deleted.map(slim),
      labelsChanged: this.stagedHistory.labelsChanged.map(slim),
    }
  }

  async getCurrentHistoryId(): Promise<string> {
    return this.historyId
  }
}

export function seedInbox(): SeedMessage[] {
  return [
    {
      id: 'm1',
      threadId: 't1',
      labelIds: ['INBOX', 'UNREAD', 'CATEGORY_PERSONAL'],
      headers: {
        From: 'Alice Example <alice@example.com>',
        To: 'me@example.com',
        Subject: 'Invoice for October',
        Date: 'Tue, 14 Oct 2026 10:00:00 +0800',
      },
      text: 'Hello, the October invoice is attached.',
      attachments: [{ filename: 'invoice.txt', mime: 'text/plain', text: 'invoice total 100' }],
    },
    {
      id: 'm2',
      threadId: 't1',
      labelIds: ['INBOX', 'STARRED'],
      headers: {
        From: 'Bob Example <bob@example.com>',
        To: 'me@example.com',
        Cc: 'carol@example.com',
        Subject: 'Re: Invoice for October',
        Date: 'Tue, 14 Oct 2026 11:00:00 +0800',
      },
      text: 'Looks good, approved.',
      html: '<p>Looks <b>good</b>, approved.</p>',
    },
  ]
}
