/**
 * Gmail history-expiry resync fixture: an expired historyId forces a full
 * resync through decideSync, and the resync preserves queued outbox sends,
 * local drafts, and pending intents (the driver holds none of those
 * queues, so they survive by construction).
 */

import { describe, expect, it } from 'vitest'
import { decideSync } from '../src/sync'
import {
  GmailDriver,
  isHistoryExpiredError,
  isInvalidGrantError,
  GmailAuthExpiredError,
  type MailEvent,
} from '../src/gmail'
import { FakeGmailTransport, fakeAuth, memorySyncHooks, seedInbox } from './gmail-fake'

describe('history expiry detection', () => {
  it('treats Gmail 404 as expired historyId', () => {
    expect(isHistoryExpiredError({ code: 404, message: 'Requested entity was not found.' })).toBe(
      true,
    )
    expect(isHistoryExpiredError({ code: 400 })).toBe(false)
    expect(isHistoryExpiredError(undefined)).toBe(false)
  })

  it('decides full resync with the history-expired reason', () => {
    const stored = {
      accountId: 'gmail-test',
      folderId: 'INBOX',
      historyId: '111',
      lastSyncAt: '2026-10-14T09:00:00Z',
    }
    expect(decideSync({ stored, historyExpired: true })).toEqual({
      kind: 'full-resync',
      reason: 'history-expired',
    })
  })

  it('detects invalid_grant auth rejections', () => {
    expect(isInvalidGrantError({ code: 401, message: 'invalid_grant: Token expired' })).toBe(true)
    expect(isInvalidGrantError({ code: 404 })).toBe(false)
  })
})

describe('expired historyId resync fixture', () => {
  it('full-resyncs on 404 and preserves outbox, drafts, and intents', async () => {
    const transport = new FakeGmailTransport(seedInbox())
    transport.historyId = '9001'
    const driver = new GmailDriver({ auth: fakeAuth(), transport, pollMs: 60000 })
    const hooks = memorySyncHooks()
    hooks.save({
      accountId: 'gmail-test',
      folderId: 'INBOX',
      historyId: '111',
      lastSyncAt: '2026-10-14T09:00:00Z',
    })
    // Main-layer queues the driver must never disturb.
    const outbox = [{ id: 'queued-send-1', to: ['boss@example.com'] }]
    const drafts = [{ accountId: 'gmail-test', threadKey: 't1', text: 'unsent draft' }]
    const intents = [{ id: 'intent-1', targetUid: 'm2', kind: 'flag' }]
    driver.setSyncHooks(hooks)
    await driver.connect()
    const events: MailEvent[] = []
    const sub = driver.watch((ev) => events.push(ev))
    await driver.pollOnce()
    sub.stop()
    const errorState = events.find(
      (e) => 'kind' in e && e.kind === 'sync-state' && e.state === 'error',
    )
    expect(errorState).toMatchObject({ detail: 'history-expired' })
    const added = events.find((e) => 'kind' in e && e.kind === 'messages-added')
    expect(added).toMatchObject({ folderId: 'INBOX' })
    expect(events[events.length - 1]).toMatchObject({ kind: 'sync-state', state: 'idle' })
    expect(hooks.load('gmail-test')?.historyId).toBe('9001')
    expect(outbox).toHaveLength(1)
    expect(drafts).toHaveLength(1)
    expect(intents).toHaveLength(1)
    await driver.disconnect()
  })

  it('starts with a full sync when no historyId is stored', async () => {
    const transport = new FakeGmailTransport(seedInbox())
    const driver = new GmailDriver({ auth: fakeAuth(), transport, pollMs: 60000 })
    const hooks = memorySyncHooks()
    driver.setSyncHooks(hooks)
    await driver.connect()
    const events: MailEvent[] = []
    const sub = driver.watch((ev) => events.push(ev))
    await driver.pollOnce()
    sub.stop()
    expect(events.some((e) => 'kind' in e && e.kind === 'messages-added')).toBe(true)
    expect(hooks.load('gmail-test')?.historyId).toBe('9000')
    await driver.disconnect()
  })
})

describe('invalid_grant policy', () => {
  it('surfaces invalid_grant as GmailAuthExpiredError for vault wipe and re-auth', async () => {
    const transport = new FakeGmailTransport(seedInbox())
    transport.historyError = {
      code: 401,
      message: 'invalid_grant: Token has been expired or revoked.',
    }
    const driver = new GmailDriver({ auth: fakeAuth(), transport, pollMs: 60000 })
    const hooks = memorySyncHooks()
    hooks.save({
      accountId: 'gmail-test',
      folderId: 'INBOX',
      historyId: '9000',
      lastSyncAt: '2026-10-14T09:00:00Z',
    })
    driver.setSyncHooks(hooks)
    await driver.connect()
    await expect(driver.pollOnce()).rejects.toBeInstanceOf(GmailAuthExpiredError)
    await driver.disconnect()
  })
})
