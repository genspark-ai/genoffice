import { describe, expect, it } from 'vitest'
import type { MailIntent, SyncState } from '../src/types'
import {
  decideSync,
  intentKey,
  nextIntent,
  resolveIntentAttempt,
  type IntentStep,
} from '../src/sync'

const stored = (over: Partial<SyncState> = {}): SyncState => ({
  accountId: 'a1',
  folderId: 'inbox',
  uidValidity: 100,
  highestModseq: '500',
  lastSyncAt: '2026-10-14T09:00:00Z',
  ...over,
})

const intent = (over: Partial<MailIntent> = {}): MailIntent => ({
  id: 'i1',
  accountId: 'a1',
  kind: 'move',
  targetUid: '42',
  expectedModseq: '500',
  payloadJson: '{}',
  state: 'queued',
  ...over,
})

describe('decideSync', () => {
  it('starts initial with no stored state', () => {
    expect(decideSync({ stored: undefined })).toEqual({ kind: 'initial' })
  })

  it('stays incremental when validity holds and no expiry', () => {
    expect(decideSync({ stored: stored(), remoteUidValidity: 100 })).toEqual({
      kind: 'incremental',
    })
  })

  it('forces full re-fetch on UIDVALIDITY change', () => {
    expect(decideSync({ stored: stored(), remoteUidValidity: 101 })).toEqual({
      kind: 'full-resync',
      reason: 'uidvalidity-changed',
    })
  })

  it('forces full resync on expired Gmail historyId', () => {
    expect(decideSync({ stored: stored({ historyId: '999' }), historyExpired: true })).toEqual({
      kind: 'full-resync',
      reason: 'history-expired',
    })
  })

  it('forces full resync on expired Graph delta token', () => {
    expect(decideSync({ stored: stored({ deltaToken: 'tok' }), deltaExpired: true })).toEqual({
      kind: 'full-resync',
      reason: 'delta-expired',
    })
  })
})

describe('nextIntent', () => {
  it('replays oldest queued first and skips settled intents', () => {
    const step = nextIntent(
      [intent({ id: 'done', state: 'acked' }), intent({ id: 'i1' }), intent({ id: 'i2' })],
      new Map(),
    )
    expect(step?.intent.id).toBe('i1')
    expect(step?.attempt).toBe(1)
  })

  it('returns undefined when nothing is queued', () => {
    expect(nextIntent([intent({ state: 'failed' })], new Map())).toBeUndefined()
  })

  it('counts attempts from the caller map', () => {
    const step = nextIntent([intent({ id: 'i1' })], new Map([['i1', 1]]))
    expect(step?.attempt).toBe(2)
  })
})

describe('resolveIntentAttempt', () => {
  const step: IntentStep = { intent: intent(), attempt: 1 }

  it('acks on success', () => {
    expect(resolveIntentAttempt(step, 'acked')).toEqual({ action: 'ack' })
  })

  it('refetches then retries once on stale modseq', () => {
    expect(resolveIntentAttempt(step, 'stale-modseq')).toEqual({ action: 'retry-after-refetch' })
  })

  it('surfaces a conflict on the second stale modseq, never a merge guess', () => {
    expect(resolveIntentAttempt({ ...step, attempt: 2 }, 'stale-modseq')).toEqual({
      action: 'fail',
      reason: 'conflict',
    })
  })

  it('fails transport errors directly', () => {
    expect(resolveIntentAttempt(step, 'failed')).toEqual({ action: 'fail', reason: 'transport' })
  })
})

describe('intentKey', () => {
  it('binds target plus expected modseq for idempotency', () => {
    expect(intentKey('a1', '42', '500')).toBe(intentKey('a1', '42', '500'))
    expect(intentKey('a1', '42', '500')).not.toBe(intentKey('a1', '42', '501'))
    expect(intentKey('a1', '42')).not.toBe(intentKey('a1', '43'))
  })
})
