import type { AiStreamChunk } from '@genoffice/ai-provider'
import type { MailEnvelope, MailMessageRef } from '@genoffice/mail-protocols'
import type { AiRetrieveDeps, AiRetrieveStorePort } from './retrieve'

/**
 * Stubbed AI provider fixture for tests and stubbed-provider e2e (P5 spec
 * section 11). Implements the same AiRetrieveDeps surface as production:
 * two synthetic accounts with disjoint mail, scope-enforced reads, and a
 * deterministic streaming stub shaped like AiStreamChunk. Synthetic
 * example.com mail only — never real user mail.
 */

export interface StubSeedMessage {
  uid: string
  folderId: string
  from: string[]
  subject: string
  snippet: string
  text: string
  msgIdHeader?: string
}

export interface StubSeedAccount {
  accountId: string
  messages: StubSeedMessage[]
}

function toEnvelope(message: StubSeedMessage): MailEnvelope {
  return {
    ref: {
      uid: message.uid,
      folderId: message.folderId,
      ...(message.msgIdHeader !== undefined ? { msgIdHeader: message.msgIdHeader } : {}),
      threadId: `thread-${message.uid}`,
    },
    from: message.from,
    to: ['owner@example.com'],
    cc: [],
    subject: message.subject,
    date: '2026-10-01T09:00:00.000Z',
    flags: { seen: false, flagged: false, answered: false, draft: false },
    labels: [],
    snippet: message.snippet,
    hasAttach: false,
  }
}

function includes(haystack: string, needle: string): boolean {
  return haystack.toLowerCase().includes(needle.toLowerCase())
}

function stubPort(messages: StubSeedMessage[]): AiRetrieveStorePort {
  return {
    search(_accountId, query, opts) {
      const needle = query.trim().toLowerCase()
      if (!needle) return []
      return messages
        .filter((message) => opts.folderId === undefined || message.folderId === opts.folderId)
        .filter((message) =>
          includes([...message.from, message.subject, message.snippet].join(' '), needle),
        )
        .slice(0, opts.limit)
        .map(toEnvelope)
    },
    body(_accountId, folderId, uid) {
      return messages.find((message) => message.folderId === folderId && message.uid === uid)?.text
    },
    envelope(_accountId, folderId, uid) {
      const found = messages.find((message) => message.folderId === folderId && message.uid === uid)
      return found ? toEnvelope(found) : undefined
    },
  }
}

/**
 * Build stub retrieval deps from synthetic per-account seeds. Reads are
 * scoped by construction: each port serves exactly one account, and the
 * retrieval core never asks a port about another account.
 */
export function createStubAiDeps(seeds: StubSeedAccount[]): AiRetrieveDeps {
  const byAccount = new Map(seeds.map((seed) => [seed.accountId, stubPort(seed.messages)]))
  return {
    knownAccountIds: () => seeds.map((seed) => seed.accountId),
    storeFor: (accountId: string) => {
      const port = byAccount.get(accountId)
      if (!port) throw new Error(`mail: unknown stub account ${accountId}`)
      return port
    },
  }
}

/** Two-account leak fixture: disjoint synthetic mail for scoping tests. */
export function twoAccountStubSeeds(): StubSeedAccount[] {
  return [
    {
      accountId: 'stub-a',
      messages: [
        {
          uid: 'a1',
          folderId: 'inbox',
          from: ['billing@example.com'],
          subject: 'Invoice INV-A1 due Monday',
          snippet: 'Your invoice INV-A1 is due Monday.',
          text: 'Hello, your invoice INV-A1 for 120 credits is due Monday. Thank you.',
          msgIdHeader: '<a1@example.com>',
        },
      ],
    },
    {
      accountId: 'stub-b',
      messages: [
        {
          uid: 'b1',
          folderId: 'inbox',
          from: ['clinic@example.net'],
          subject: 'Appointment reminder B1',
          snippet: 'Your appointment B1 is confirmed for Friday.',
          text: 'Hello, your appointment B1 at the clinic is confirmed for Friday.',
          msgIdHeader: '<b1@example.net>',
        },
      ],
    },
  ]
}

/** Ref helper for the leak fixtures (folder + uid + thread triple). */
export function stubRef(folderId: string, uid: string): MailMessageRef {
  return { uid, folderId, threadId: `thread-${uid}` }
}

/** Ref helper carrying an exact-match account hint (D-H1-C3). */
export function stubRefHinted(folderId: string, uid: string, accountId: string): MailMessageRef {
  return { uid, folderId, threadId: `thread-${uid}`, accountId }
}

/**
 * Same-uid two-account fixture for exact-match tests: both scope members
 * hold `(inbox, dup1)` with disjoint synthetic mail, so resolution order
 * and hint preference are observable.
 */
export function sameUidTwoAccountSeeds(): StubSeedAccount[] {
  return [
    {
      accountId: 'stub-a',
      messages: [
        {
          uid: 'dup1',
          folderId: 'inbox',
          from: ['billing@example.com'],
          subject: 'Alpha invoice dup1',
          snippet: 'Alpha copy of dup1.',
          text: 'Alpha body for dup1, billing marker one.',
          msgIdHeader: '<dup1-a@example.com>',
        },
      ],
    },
    {
      accountId: 'stub-b',
      messages: [
        {
          uid: 'dup1',
          folderId: 'inbox',
          from: ['clinic@example.net'],
          subject: 'Beta reminder dup1',
          snippet: 'Beta copy of dup1.',
          text: 'Beta body for dup1, clinic marker two.',
          msgIdHeader: '<dup1-b@example.net>',
        },
      ],
    },
  ]
}

/**
 * Deterministic stubbed stream: splits text into word-bounded delta chunks
 * plus a terminal done chunk, all echoing the request id. Lets transport
 * tests assert chunk ordering and completion without a live provider.
 */
export function collectStubStream(requestId: string, text: string): AiStreamChunk[] {
  const chunks: AiStreamChunk[] = []
  const words = text.split(/\s+/).filter((word) => word !== '')
  let carried = ''
  for (const word of words) {
    carried += (carried === '' ? '' : ' ') + word
    if (carried.length >= 24) {
      chunks.push({ requestId, type: 'delta', text: carried })
      carried = ''
    }
  }
  if (carried !== '') chunks.push({ requestId, type: 'delta', text: carried })
  chunks.push({ requestId, type: 'done' })
  return chunks
}
