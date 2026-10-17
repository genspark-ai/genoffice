/**
 * matchesQuery: the shared local post-filter behind the P3 sync-engine
 * contract (all ten MailQuery fields honored on every driver). Pure,
 * no I/O, no credentials.
 */
import { describe, expect, it } from 'vitest'
import { matchesQuery } from '../src/query.js'
import type { MailEnvelope } from '../src/types.js'

function makeEnv(overrides?: Partial<MailEnvelope>): MailEnvelope {
  return {
    ref: { uid: '1', folderId: 'INBOX', threadId: 'INBOX:1' },
    from: ['a@example.local'],
    to: ['me@example.local'],
    cc: [],
    subject: 'Q1 invoice report',
    date: '2026-02-10T10:00:00.000Z',
    flags: { seen: false, flagged: true, answered: false, draft: false },
    labels: [],
    snippet: 'payment due this quarter',
    hasAttach: true,
    ...overrides,
  }
}

describe('matchesQuery', () => {
  it('matches everything on an empty query', () => {
    expect(matchesQuery(makeEnv(), {})).toBe(true)
  })

  it('matches text across from/to/subject/snippet, case-insensitively', () => {
    expect(matchesQuery(makeEnv(), { text: 'INVOICE' })).toBe(true)
    expect(matchesQuery(makeEnv(), { text: 'quarter' })).toBe(true)
    expect(matchesQuery(makeEnv(), { text: 'no-such-phrase' })).toBe(false)
  })

  it('matches from/to/subject substrings', () => {
    expect(matchesQuery(makeEnv(), { from: 'A@EXAMPLE' })).toBe(true)
    expect(matchesQuery(makeEnv(), { to: 'me@example' })).toBe(true)
    expect(matchesQuery(makeEnv(), { to: 'boss@example' })).toBe(false)
    expect(matchesQuery(makeEnv(), { subject: 'invoice report' })).toBe(true)
    expect(matchesQuery(makeEnv(), { subject: 'lunch' })).toBe(false)
  })

  it('matches flag and attachment predicates exactly', () => {
    expect(matchesQuery(makeEnv(), { seen: false, flagged: true, hasAttach: true })).toBe(true)
    expect(matchesQuery(makeEnv(), { seen: true })).toBe(false)
    expect(matchesQuery(makeEnv(), { flagged: false })).toBe(false)
    expect(matchesQuery(makeEnv(), { hasAttach: false })).toBe(false)
  })

  it('compares since/before at day granularity', () => {
    expect(matchesQuery(makeEnv(), { since: '2026-02-10' })).toBe(true)
    expect(matchesQuery(makeEnv(), { since: '2026-02-11' })).toBe(false)
    expect(matchesQuery(makeEnv(), { before: '2026-02-11' })).toBe(true)
    expect(matchesQuery(makeEnv(), { before: '2026-02-10' })).toBe(false)
    expect(matchesQuery(makeEnv(), { since: '2026-01-01', before: '2026-03-01' })).toBe(true)
  })

  it('ignores unparsable date clauses instead of failing closed', () => {
    expect(matchesQuery(makeEnv(), { since: 'not-a-date' })).toBe(true)
    expect(matchesQuery(makeEnv(), { before: 'not-a-date' })).toBe(true)
  })

  it('treats onlyUnread as unseen-only', () => {
    expect(matchesQuery(makeEnv(), { onlyUnread: true })).toBe(true)
    const seen = makeEnv({ flags: { seen: true, flagged: false, answered: false, draft: false } })
    expect(matchesQuery(seen, { onlyUnread: true })).toBe(false)
    expect(matchesQuery(seen, { onlyUnread: false })).toBe(true)
  })
})
