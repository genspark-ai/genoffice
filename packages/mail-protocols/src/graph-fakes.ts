/**
 * Staging fake for GraphDriver tests: an in-memory GraphTransport with no
 * network and no credentials. This is the "staging" side of the round-trip
 * matrix row (graph / staging-fake): seed folders and messages, drive the
 * full MailProvider surface, and assert on captured sends/moves/patches.
 * It also simulates expired delta tokens (410 resyncRequired), paging
 * nextLinks, and one-shot auth failures including invalid_grant.
 */

import { GraphServiceError, type GraphRequestInit, type GraphTransport } from './graph-transport'
import type { GraphFolder, GraphMessage } from './graph-driver'

export interface FakeSentMail {
  message: Record<string, unknown>
  saveToSentItems: unknown
}

export interface FakePatch {
  path: string
  body: unknown
}

function clone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T
}

let fakeIdCounter = 0

export function makeGraphMessage(over: Partial<GraphMessage> & { id: string }): GraphMessage {
  fakeIdCounter += 1
  return {
    conversationId: over.id,
    subject: '',
    receivedDateTime: '2026-10-14T10:00:00Z',
    isRead: false,
    flag: { flagStatus: 'notFlagged' },
    categories: [],
    hasAttachments: false,
    ...over,
  }
}

export class FakeGraphTransport implements GraphTransport {
  folders: GraphFolder[] = [
    { id: 'inbox', displayName: 'Inbox', unreadItemCount: 1 },
    { id: 'sentitems', displayName: 'Sent Items', unreadItemCount: 0 },
    { id: 'drafts', displayName: 'Drafts', unreadItemCount: 0 },
    { id: 'deleteditems', displayName: 'Deleted Items', unreadItemCount: 0 },
    { id: 'junkemail', displayName: 'Junk Email', unreadItemCount: 0 },
    { id: 'archive', displayName: 'Archive', unreadItemCount: 0 },
    { id: 'custom-1', displayName: 'Receipts', unreadItemCount: 0 },
  ]

  messages = new Map<string, GraphMessage[]>()
  sentMail: FakeSentMail[] = []
  patches: FakePatch[] = []
  deletes: string[] = []
  moves: Array<{ id: string; destinationId: unknown }> = []

  /** Queued outbox/draft/intent handles owned by the TEST, never the driver. */
  leftovers: { outbox: string[]; intents: string[] } = { outbox: [], intents: [] }

  private deltaVersion = 1
  private deltaChanges: GraphMessage[] = []
  private deltaRemovals: string[] = []
  private failNext: GraphServiceError | undefined

  /** Next request throws this service error once (auth-failure simulation). */
  failNextWith(err: GraphServiceError): void {
    this.failNext = err
  }

  /** Invalidate the current delta token without issuing a new one. */
  expireDelta(): void {
    this.deltaVersion += 1
  }

  currentDeltaToken(folderId: string): string {
    return `delta_${folderId}_v${this.deltaVersion}`
  }

  pushDeltaChange(msg: GraphMessage): void {
    this.deltaChanges.push(clone(msg))
  }

  seed(folderId: string, msgs: GraphMessage[]): void {
    this.messages.set(
      folderId,
      msgs.map((m) => ({ ...clone(m), parentFolderId: folderId })),
    )
  }

  allIn(folderId: string): GraphMessage[] {
    return (this.messages.get(folderId) ?? []).map(clone)
  }

  private takeFailure(): GraphServiceError | undefined {
    const err = this.failNext
    this.failNext = undefined
    return err
  }

  private findMessage(
    id: string,
  ): { folderId: string; index: number; msg: GraphMessage } | undefined {
    for (const [folderId, list] of this.messages) {
      const index = list.findIndex((m) => m.id === id)
      if (index >= 0) return { folderId, index, msg: list[index] }
    }
    return undefined
  }

  private split(path: string): { clean: string; query: URLSearchParams } {
    if (path.startsWith('https://')) {
      const url = new URL(path)
      return { clean: url.pathname.replace(/^\/v1\.0/, ''), query: url.searchParams }
    }
    const url = new URL(path, 'https://graph.microsoft.com')
    return { clean: url.pathname, query: url.searchParams }
  }

  async get<T>(path: string, _init?: GraphRequestInit): Promise<T> {
    const failure = this.takeFailure()
    if (failure !== undefined) throw failure
    const { clean, query } = this.split(path)

    if (clean === '/me') return clone({ id: 'user-1' }) as T

    const nextMatch = /^\/next$/.exec(clean)
    if (nextMatch !== null) {
      const folderId = query.get('folder') ?? 'inbox'
      const skip = Number(query.get('skip') ?? '0')
      const top = Number(query.get('top') ?? '25')
      return this.page(folderId, skip, top) as T
    }

    const wellKnown = /^\/me\/mailFolders\/([^/]+)$/.exec(clean)
    if (wellKnown !== null && !clean.endsWith('/messages') && !clean.endsWith('/messages/delta')) {
      const folder = this.folders.find((f) => f.id === decodeURIComponent(wellKnown[1]))
      if (folder === undefined)
        throw new GraphServiceError(404, 'ErrorItemNotFound', 'Folder not found')
      return clone({ id: folder.id }) as T
    }

    if (clean === '/me/mailFolders') {
      return clone({ value: this.folders }) as T
    }

    const deltaMatch = /^\/me\/mailFolders\/([^/]+)\/messages\/delta$/.exec(clean)
    if (deltaMatch !== null) {
      const folderId = decodeURIComponent(deltaMatch[1])
      const token = query.get('$deltatoken') ?? query.get('deltatoken')
      if (token !== null && token !== this.currentDeltaToken(folderId)) {
        throw new GraphServiceError(410, 'resyncRequired', 'Delta token expired')
      }
      if (token === null) {
        const value = this.allIn(folderId)
        return clone({
          value,
          '@odata.deltaLink': `https://graph.microsoft.com/v1.0/delta?token=${this.currentDeltaToken(folderId)}`,
        }) as T
      }
      this.deltaVersion += 1
      const value = [
        ...this.deltaChanges.map(clone),
        ...this.deltaRemovals.map((id) => ({ id, '@removed': { reason: 'deleted' } })),
      ]
      this.deltaChanges = []
      this.deltaRemovals = []
      return clone({
        value,
        '@odata.deltaLink': `https://graph.microsoft.com/v1.0/delta?token=${this.currentDeltaToken(folderId)}`,
      }) as T
    }

    const folderMessages = /^\/me\/mailFolders\/([^/]+)\/messages$/.exec(clean)
    if (folderMessages !== null) {
      const folderId = decodeURIComponent(folderMessages[1])
      let list = this.allIn(folderId)
      const search = query.get('$search')
      if (search !== null) list = this.applySearch(list, search)
      const filter = query.get('$filter')
      if (filter !== null) list = this.applyFilter(list, filter)
      list.sort((a, b) => (b.receivedDateTime ?? '').localeCompare(a.receivedDateTime ?? ''))
      const top = Number(query.get('$top') ?? '25')
      return this.pageList(folderId, list, 0, top) as T
    }

    const single = /^\/me\/messages\/([^/]+)$/.exec(clean)
    if (single !== null) {
      const found = this.findMessage(decodeURIComponent(single[1]))
      if (found === undefined)
        throw new GraphServiceError(404, 'ErrorItemNotFound', 'Message not found')
      return clone(found.msg) as T
    }

    throw new GraphServiceError(400, 'BadRequest', `Fake has no route for ${clean}`)
  }

  private page(
    folderId: string,
    skip: number,
    top: number,
  ): { value: GraphMessage[]; '@odata.nextLink'?: string } {
    const list = this.allIn(folderId).sort((a, b) =>
      (b.receivedDateTime ?? '').localeCompare(a.receivedDateTime ?? ''),
    )
    return this.pageList(folderId, list, skip, top)
  }

  private pageList(
    folderId: string,
    list: GraphMessage[],
    skip: number,
    top: number,
  ): { value: GraphMessage[]; '@odata.nextLink'?: string } {
    const slice = list.slice(skip, skip + top)
    const out: { value: GraphMessage[]; '@odata.nextLink'?: string } = { value: slice }
    if (skip + top < list.length) {
      out['@odata.nextLink'] =
        `https://graph.microsoft.com/v1.0/next?folder=${encodeURIComponent(folderId)}&skip=${skip + top}&top=${top}`
    }
    return out
  }

  private applySearch(list: GraphMessage[], search: string): GraphMessage[] {
    const fromMatch = /from:([^\s]+)/.exec(search)
    const unreadOnly = search.includes('isRead:false')
    const sinceMatch = /received>=([0-9-]+)/.exec(search)
    const quoted = /"([^"]*)"/.exec(search)
    return list.filter((m) => {
      if (fromMatch !== null && !(m.from?.emailAddress?.address ?? '').includes(fromMatch[1]))
        return false
      if (unreadOnly && (m.isRead ?? false)) return false
      if (sinceMatch !== null && (m.receivedDateTime ?? '') < sinceMatch[1]) return false
      if (quoted !== null) {
        const hay = `${m.subject ?? ''} ${m.bodyPreview ?? ''} ${m.body?.content ?? ''}`
        if (!hay.includes(quoted[1])) return false
      }
      return true
    })
  }

  private applyFilter(list: GraphMessage[], filter: string): GraphMessage[] {
    const ltMatch = /receivedDateTime lt (.+)/.exec(filter)
    if (ltMatch === null) return list
    const bound = decodeURIComponent(ltMatch[1]).trim()
    return list.filter((m) => (m.receivedDateTime ?? '') < bound)
  }

  async post<T>(path: string, body: unknown, _init?: GraphRequestInit): Promise<T> {
    const failure = this.takeFailure()
    if (failure !== undefined) throw failure
    const { clean } = this.split(path)

    if (clean === '/me/sendMail') {
      const payload = body as { message: Record<string, unknown>; saveToSentItems: unknown }
      this.sentMail.push({
        message: clone(payload.message),
        saveToSentItems: payload.saveToSentItems,
      })
      const sent = this.messages.get('sentitems') ?? []
      fakeIdCounter += 1
      sent.push(
        makeGraphMessage({
          id: `sent-${fakeIdCounter}`,
          subject: (payload.message['subject'] as string | undefined) ?? '',
          parentFolderId: 'sentitems',
        }),
      )
      this.messages.set('sentitems', sent)
      return clone({}) as T
    }

    const moveMatch = /^\/me\/messages\/([^/]+)\/move$/.exec(clean)
    if (moveMatch !== null) {
      const found = this.findMessage(decodeURIComponent(moveMatch[1]))
      if (found === undefined)
        throw new GraphServiceError(404, 'ErrorItemNotFound', 'Message not found')
      const destinationId = (body as { destinationId: unknown }).destinationId as string
      this.moves.push({ id: found.msg.id, destinationId })
      const [moved] = this.messages.get(found.folderId)?.splice(found.index, 1) ?? []
      if (moved !== undefined) {
        moved.parentFolderId = destinationId
        const target = this.messages.get(destinationId) ?? []
        target.push(moved)
        this.messages.set(destinationId, target)
      }
      return clone({ id: found.msg.id }) as T
    }

    throw new GraphServiceError(400, 'BadRequest', `Fake has no route for ${clean}`)
  }

  async patch<T>(path: string, body: unknown, _init?: GraphRequestInit): Promise<T> {
    const failure = this.takeFailure()
    if (failure !== undefined) throw failure
    const { clean } = this.split(path)
    const single = /^\/me\/messages\/([^/]+)$/.exec(clean)
    if (single === null)
      throw new GraphServiceError(400, 'BadRequest', `Fake has no route for ${clean}`)
    const found = this.findMessage(decodeURIComponent(single[1]))
    if (found === undefined)
      throw new GraphServiceError(404, 'ErrorItemNotFound', 'Message not found')
    const patch = body as { isRead?: boolean; flag?: { flagStatus?: string } }
    if (patch.isRead !== undefined) found.msg.isRead = patch.isRead
    if (patch.flag?.flagStatus !== undefined) found.msg.flag = { flagStatus: patch.flag.flagStatus }
    this.patches.push({ path: clean, body: clone(patch) })
    return clone({}) as T
  }

  async del(path: string, _init?: GraphRequestInit): Promise<void> {
    const failure = this.takeFailure()
    if (failure !== undefined) throw failure
    const { clean } = this.split(path)
    const single = /^\/me\/messages\/([^/]+)$/.exec(clean)
    if (single === null)
      throw new GraphServiceError(400, 'BadRequest', `Fake has no route for ${clean}`)
    const found = this.findMessage(decodeURIComponent(single[1]))
    if (found === undefined)
      throw new GraphServiceError(404, 'ErrorItemNotFound', 'Message not found')
    this.messages.get(found.folderId)?.splice(found.index, 1)
    this.deletes.push(found.msg.id)
    this.deltaRemovals.push(found.msg.id)
  }
}
