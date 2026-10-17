import { describe, expect, it } from 'vitest'
import { InMemoryMailStore } from '../src/main/mail-store'
import type { MailEnvelope } from '@genoffice/mail-protocols'

function envelope(uid: string, subject: string): MailEnvelope {
  return {
    ref: { uid, folderId: 'inbox', threadId: 't1' },
    from: ['a@example.com'],
    to: ['b@example.com'],
    cc: [],
    subject,
    date: '2026-10-14',
    flags: { seen: false, flagged: false, answered: false, draft: false },
    labels: [],
    snippet: 'hello world',
    hasAttach: false,
  }
}

describe('InMemoryMailStore stub', () => {
  it('stores envelopes, threads, drafts and outbox status', () => {
    const store = new InMemoryMailStore()
    store.upsertEnvelopes([envelope('1', 'Invoice'), envelope('2', 'Lunch')])
    expect(store.getThread('t1')).toHaveLength(2)
    expect(store.fts('invoice', ['a'])).toHaveLength(1)
    store.saveDraft('a', 't1', 'draft text')
    expect(store.draftFor('a', 't1')).toBe('draft text')
    const id = store.enqueueSend({ to: ['b@example.com'], subject: 's', text: 'b' })
    expect(id.startsWith('outbox-')).toBe(true)
    expect(store.outboxSize()).toBe(1)
  })

  it('queues and resolves offline intents idempotently', () => {
    const store = new InMemoryMailStore()
    const id = store.queueIntent({
      accountId: 'a',
      kind: 'move',
      targetUid: '1',
      payloadJson: '{}',
    })
    expect(store.pendingIntents('a')).toHaveLength(1)
    store.markIntent(id, 'acked')
    expect(store.pendingIntents('a')).toHaveLength(0)
  })

  it('round-trips sync state', () => {
    const store = new InMemoryMailStore()
    expect(store.getSyncState('a', 'inbox')).toBeUndefined()
    store.setSyncState({ accountId: 'a', folderId: 'inbox', lastSyncAt: '2026-10-14' })
    expect(store.getSyncState('a', 'inbox')?.lastSyncAt).toBe('2026-10-14')
  })
})
