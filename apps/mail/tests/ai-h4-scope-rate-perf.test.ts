import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { beforeEach, describe, expect, it } from 'vitest'
import { MAIL_CHANNELS } from '../src/shared/ipc'
import type { AiRetrieveResponse } from '../src/shared/ipc'
import { mailAiRetrieveSchema } from '../src/shared/ipc-schemas'
import type { ComposeIpcBus } from '../src/main/compose-handlers'
import {
  aiRetrieve,
  createAiRetrieveRateLimiter,
  createStubAiDeps,
  isScopeStale,
  normalizeAiScope,
  registerAiRetrieveHandlers,
  resetAiRetrieveHandlers,
  resetSharedAiRetrieveRateLimiter,
  sameUidTwoAccountSeeds,
  stubRef,
  stubRefHinted,
  twoAccountStubSeeds,
} from '../src/main/ai'
import type { AiRetrieveDeps } from '../src/main/ai'
import { resetMailDemoSeed } from '../src/main/mail-read-handlers'
import { configureMailStoreDir, getMailStore, resetMailStoreConfig } from '../src/main/mail-store'
import { closeAllMailStores } from '../src/main/store/store'
import { containsPii, PROMPT_TEXT_CAP_CHARS, redactForPrompt } from '../src/main/sync/redact'

/**
 * H4 transport/scope hardening (H1 sections 4.2 + 6): scope-enforcement
 * fuzz (unknown/empty/stale scope, cross-account confusion), D-H1-C3
 * exact-match preference with two-account same-uid fixtures, P-9
 * token-bucket rate limits, and P-6..P-8 throughput budget pins.
 * Synthetic example.com mail only; stubbed stores throughout.
 */

function makeBus(): {
  bus: ComposeIpcBus
  invoke: (channel: string, payload?: unknown) => Promise<unknown>
} {
  const handlers = new Map<string, (payload: unknown) => Promise<unknown>>()
  return {
    bus: {
      handle: (channel, listener) => {
        handlers.set(channel, listener)
      },
    },
    invoke: async (channel, payload) => {
      const listener = handlers.get(channel)
      if (!listener) throw new Error(`no handler: ${channel}`)
      return listener(payload)
    },
  }
}

/** Stub deps with an isolated limiter so budget timing never throttles. */
function isolatedDeps(
  seeds: ReturnType<typeof twoAccountStubSeeds>,
  now: () => number = Date.now,
): AiRetrieveDeps {
  return { ...createStubAiDeps(seeds), rateLimiter: createAiRetrieveRateLimiter(now) }
}

beforeEach(() => {
  resetAiRetrieveHandlers()
  resetSharedAiRetrieveRateLimiter()
  resetMailDemoSeed()
  closeAllMailStores()
  resetMailStoreConfig()
})

describe('H4 scope-enforcement fuzz', () => {
  it('rejects unknown scope accounts before opening any store', async () => {
    const base = createStubAiDeps(twoAccountStubSeeds())
    let opened = 0
    const deps: AiRetrieveDeps = {
      ...base,
      rateLimiter: createAiRetrieveRateLimiter(),
      storeFor: async (accountId) => {
        opened += 1
        return base.storeFor(accountId)
      },
    }
    await expect(
      aiRetrieve(
        { scope: { accountIds: ['ghost'] }, op: 'triage-headers', query: 'invoice' },
        deps,
      ),
    ).rejects.toThrow('mail: unknown account in AI scope')
    expect(opened).toBe(0)
  })

  it('rejects empty, missing, and non-object scopes', () => {
    expect(() => normalizeAiScope({ accountIds: [] })).toThrow('mail:')
    expect(() => normalizeAiScope({})).toThrow('mail:')
    expect(() => normalizeAiScope(null)).toThrow('mail:')
    expect(() => normalizeAiScope('stub-a')).toThrow('mail:')
    expect(() => normalizeAiScope({ accountIds: 'stub-a' })).toThrow('mail:')
  })

  it('rejects oversize scopes at the schema and core layers', () => {
    const ids = Array.from({ length: 9 }, (_, i) => `a${i}`)
    expect(() => normalizeAiScope({ accountIds: ids })).toThrow('mail:')
    expect(
      mailAiRetrieveSchema.safeParse({
        scope: { accountIds: ids },
        op: 'triage-headers',
        query: 'x',
      }).success,
    ).toBe(false)
  })

  it('rejects malformed account ids over the channel', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, isolatedDeps(twoAccountStubSeeds()))
    for (const accountIds of [[''], [42], ['x'.repeat(257)], ['ok', '']]) {
      await expect(
        invoke(MAIL_CHANNELS.aiRetrieve, {
          scope: { accountIds },
          op: 'triage-headers',
          query: 'invoice',
        }),
      ).rejects.toThrow('mail:')
    }
  })

  it('dedupes scopes while preserving picker order', () => {
    expect(normalizeAiScope({ accountIds: ['b', 'a', 'b', 'a'] })).toEqual({
      accountIds: ['b', 'a'],
    })
  })

  it('treats injection-shaped query text as inert literals', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, isolatedDeps(twoAccountStubSeeds()))
    for (const query of ['"invoice" OR 1=1 --', 'inbox* MATCH *', '%', '_', '\\']) {
      const response = (await invoke(MAIL_CHANNELS.aiRetrieve, {
        scope: { accountIds: ['stub-a', 'stub-b'] },
        op: 'triage-headers',
        query,
      })) as AiRetrieveResponse
      expect(
        response.citations.every((cite) => ['stub-a', 'stub-b'].includes(cite.accountId)),
      ).toBe(true)
    }
  })

  it('flags stale picker selections order-insensitively', () => {
    expect(isScopeStale({ accountIds: ['a', 'b'] }, { accountIds: ['b', 'a'] })).toBe(false)
    expect(isScopeStale({ accountIds: ['a', 'b'] }, { accountIds: ['a'] })).toBe(true)
    expect(isScopeStale({ accountIds: ['a'] }, { accountIds: ['a', 'b'] })).toBe(true)
    expect(isScopeStale({ accountIds: ['a'] }, { accountIds: ['b'] })).toBe(true)
  })

  it('keeps cross-account confusion deterministic by scope order without a hint', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, isolatedDeps(sameUidTwoAccountSeeds()))
    const first = (await invoke(MAIL_CHANNELS.aiRetrieve, {
      scope: { accountIds: ['stub-a', 'stub-b'] },
      op: 'fetch-bodies',
      refs: [stubRef('inbox', 'dup1')],
    })) as AiRetrieveResponse
    expect(first.bodies).toHaveLength(1)
    expect(first.bodies?.[0].accountId).toBe('stub-a')
    expect(first.bodies?.[0].text).toContain('billing marker one')
    const flipped = (await invoke(MAIL_CHANNELS.aiRetrieve, {
      scope: { accountIds: ['stub-b', 'stub-a'] },
      op: 'fetch-bodies',
      refs: [stubRef('inbox', 'dup1')],
    })) as AiRetrieveResponse
    expect(flipped.bodies?.[0].accountId).toBe('stub-b')
    expect(flipped.bodies?.[0].text).toContain('clinic marker two')
  })
})

describe('H4 D-H1-C3 exact-match preference', () => {
  it('resolves a hint exactly even against scope order', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, isolatedDeps(sameUidTwoAccountSeeds()))
    const response = (await invoke(MAIL_CHANNELS.aiRetrieve, {
      scope: { accountIds: ['stub-a', 'stub-b'] },
      op: 'fetch-bodies',
      refs: [stubRefHinted('inbox', 'dup1', 'stub-b')],
    })) as AiRetrieveResponse
    expect(response.bodies).toHaveLength(1)
    expect(response.bodies?.[0].accountId).toBe('stub-b')
    expect(response.bodies?.[0].text).toContain('clinic marker two')
    expect(response.citations).toHaveLength(1)
    expect(response.citations[0]).toMatchObject({ uid: 'dup1', accountId: 'stub-b' })
  })

  it('prefers the hint in either direction', async () => {
    const deps = isolatedDeps(sameUidTwoAccountSeeds())
    const hinted = await aiRetrieve(
      {
        scope: { accountIds: ['stub-b', 'stub-a'] },
        op: 'fetch-bodies',
        refs: [stubRefHinted('inbox', 'dup1', 'stub-a')],
      },
      deps,
    )
    expect(hinted.bodies?.[0].accountId).toBe('stub-a')
  })

  it('fails loud on stale, out-of-scope, and unknown hints without leaking', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, isolatedDeps(twoAccountStubSeeds()))
    // Hint names a scope member that does not hold the ref.
    await expect(
      invoke(MAIL_CHANNELS.aiRetrieve, {
        scope: { accountIds: ['stub-a', 'stub-b'] },
        op: 'fetch-bodies',
        refs: [stubRefHinted('inbox', 'a1', 'stub-b')],
      }),
    ).rejects.toThrow('mail: stale account hint in message ref')
    // Hint names an account outside the scope.
    await expect(
      invoke(MAIL_CHANNELS.aiRetrieve, {
        scope: { accountIds: ['stub-a'] },
        op: 'fetch-bodies',
        refs: [stubRefHinted('inbox', 'b1', 'stub-b')],
      }),
    ).rejects.toThrow('mail: stale account hint in message ref')
    // Hint names an account nobody knows.
    await expect(
      invoke(MAIL_CHANNELS.aiRetrieve, {
        scope: { accountIds: ['stub-a'] },
        op: 'cite-resolve',
        refs: [stubRefHinted('inbox', 'a1', 'ghost')],
      }),
    ).rejects.toThrow('mail: stale account hint in message ref')
  })

  it('keeps stale-hint errors free of message content', async () => {
    const deps = isolatedDeps([
      {
        accountId: 'stub-a',
        messages: [
          {
            uid: 's1',
            folderId: 'inbox',
            from: ['billing@example.com'],
            subject: 'Secret ledger s1',
            snippet: 'Top secret ledger lines.',
            text: 'Ledger body with password=hunter2 inside.',
          },
        ],
      },
      ...twoAccountStubSeeds().filter((seed) => seed.accountId === 'stub-b'),
    ])
    let message = ''
    try {
      await aiRetrieve(
        {
          scope: { accountIds: ['stub-a', 'stub-b'] },
          op: 'fetch-bodies',
          refs: [stubRefHinted('inbox', 's1', 'stub-b')],
        },
        deps,
      )
    } catch (err) {
      message = err instanceof Error ? err.message : String(err)
    }
    expect(message).toContain('mail:')
    expect(message).not.toContain('hunter2')
    expect(message).not.toContain('Ledger')
    expect(containsPii(message)).toBe(false)
  })

  it('returns every in-scope match for cite-resolve without a hint', async () => {
    const deps = isolatedDeps(sameUidTwoAccountSeeds())
    const resolved = await aiRetrieve(
      {
        scope: { accountIds: ['stub-a', 'stub-b'] },
        op: 'cite-resolve',
        refs: [stubRef('inbox', 'dup1')],
      },
      deps,
    )
    expect(resolved.citations.map((cite) => cite.accountId)).toEqual(['stub-a', 'stub-b'])
    expect(resolved.envelopes).toHaveLength(2)
    expect(resolved.bodies).toBeUndefined()
  })

  it('returns the hinted match first plus other holders for hinted cite-resolve', async () => {
    const deps = isolatedDeps(sameUidTwoAccountSeeds())
    const resolved = await aiRetrieve(
      {
        scope: { accountIds: ['stub-a', 'stub-b'] },
        op: 'cite-resolve',
        refs: [stubRefHinted('inbox', 'dup1', 'stub-b')],
      },
      deps,
    )
    expect(resolved.citations.map((cite) => cite.accountId)).toEqual(['stub-b', 'stub-a'])
    expect(resolved.citations[0]).toMatchObject({
      uid: 'dup1',
      accountId: 'stub-b',
      msgIdHeader: '<dup1-b@example.net>',
    })
  })

  it('accepts hinted refs at the schema layer', () => {
    expect(
      mailAiRetrieveSchema.safeParse({
        scope: { accountIds: ['stub-a', 'stub-b'] },
        op: 'fetch-bodies',
        refs: [{ uid: 'dup1', folderId: 'inbox', threadId: 'thread-dup1', accountId: 'stub-b' }],
      }).success,
    ).toBe(true)
    expect(
      mailAiRetrieveSchema.safeParse({
        scope: { accountIds: ['a'] },
        op: 'fetch-bodies',
        refs: [{ uid: 'x', folderId: 'inbox', threadId: 't', accountId: '' }],
      }).success,
    ).toBe(false)
  })

  it('still omits refs held only outside the scope (no-leak regression)', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, isolatedDeps(sameUidTwoAccountSeeds()))
    const triage = (await invoke(MAIL_CHANNELS.aiRetrieve, {
      scope: { accountIds: ['stub-a'] },
      op: 'triage-headers',
      query: 'clinic marker two',
    })) as AiRetrieveResponse
    expect(triage.envelopes).toEqual([])
    expect(triage.citations).toEqual([])
  })
})

describe('H4 P-9 rate limit on mail:ai-retrieve', () => {
  it('throttles after a burst of 10 with a retry hint', async () => {
    const nowMs = 1_000_000
    const deps = isolatedDeps(twoAccountStubSeeds(), () => nowMs)
    for (let i = 0; i < 10; i += 1) {
      const ok = await aiRetrieve(
        { scope: { accountIds: ['stub-a'] }, op: 'triage-headers', query: 'invoice' },
        deps,
      )
      expect(ok.envelopes).toHaveLength(1)
    }
    let message = ''
    try {
      await aiRetrieve(
        { scope: { accountIds: ['stub-a'] }, op: 'triage-headers', query: 'invoice' },
        deps,
      )
    } catch (err) {
      message = err instanceof Error ? err.message : String(err)
    }
    expect(message).toMatch(/mail: rate-limited, retry in \d+s/)
    expect(containsPii(message)).toBe(false)
  })

  it('refills the bucket over time', async () => {
    let nowMs = 2_000_000
    const deps = isolatedDeps(twoAccountStubSeeds(), () => nowMs)
    for (let i = 0; i < 10; i += 1) {
      await aiRetrieve(
        { scope: { accountIds: ['stub-a'] }, op: 'triage-headers', query: 'invoice' },
        deps,
      )
    }
    await expect(
      aiRetrieve(
        { scope: { accountIds: ['stub-a'] }, op: 'triage-headers', query: 'invoice' },
        deps,
      ),
    ).rejects.toThrow('mail: rate-limited')
    nowMs += 2000
    await aiRetrieve(
      { scope: { accountIds: ['stub-a'] }, op: 'triage-headers', query: 'invoice' },
      deps,
    )
    await expect(
      aiRetrieve(
        { scope: { accountIds: ['stub-a'] }, op: 'triage-headers', query: 'invoice' },
        deps,
      ),
    ).rejects.toThrow('mail: rate-limited')
    nowMs += 120_000
    for (let i = 0; i < 10; i += 1) {
      await aiRetrieve(
        { scope: { accountIds: ['stub-a'] }, op: 'triage-headers', query: 'invoice' },
        deps,
      )
    }
  })

  it('isolates buckets per account', async () => {
    const deps = isolatedDeps(twoAccountStubSeeds())
    for (let i = 0; i < 10; i += 1) {
      await aiRetrieve(
        { scope: { accountIds: ['stub-a'] }, op: 'triage-headers', query: 'invoice' },
        deps,
      )
    }
    const other = await aiRetrieve(
      { scope: { accountIds: ['stub-b'] }, op: 'triage-headers', query: 'appointment' },
      deps,
    )
    expect(other.envelopes).toHaveLength(1)
    await expect(
      aiRetrieve(
        { scope: { accountIds: ['stub-a', 'stub-b'] }, op: 'triage-headers', query: 'invoice' },
        deps,
      ),
    ).rejects.toThrow('mail: rate-limited')
  })

  it('charges every scope member on multi-account calls', async () => {
    const deps = isolatedDeps([
      ...twoAccountStubSeeds(),
      {
        accountId: 'stub-c',
        messages: [
          {
            uid: 'c1',
            folderId: 'inbox',
            from: ['news@example.org'],
            subject: 'Digest c1',
            snippet: 'Weekly digest c1.',
            text: 'Digest body c1.',
          },
        ],
      },
    ])
    for (let i = 0; i < 10; i += 1) {
      await aiRetrieve(
        { scope: { accountIds: ['stub-a', 'stub-b'] }, op: 'triage-headers', query: 'invoice' },
        deps,
      )
    }
    await expect(
      aiRetrieve(
        { scope: { accountIds: ['stub-a'] }, op: 'triage-headers', query: 'invoice' },
        deps,
      ),
    ).rejects.toThrow('mail: rate-limited')
    const fresh = await aiRetrieve(
      { scope: { accountIds: ['stub-c'] }, op: 'triage-headers', query: 'digest' },
      deps,
    )
    expect(fresh.envelopes).toHaveLength(1)
  })

  it('does not charge rejected calls', async () => {
    const deps = isolatedDeps(twoAccountStubSeeds())
    await expect(
      aiRetrieve({ scope: { accountIds: ['ghost'] }, op: 'triage-headers', query: 'x' }, deps),
    ).rejects.toThrow('mail: unknown account')
    for (let i = 0; i < 10; i += 1) {
      await aiRetrieve(
        { scope: { accountIds: ['stub-a'] }, op: 'triage-headers', query: 'invoice' },
        deps,
      )
    }
    await expect(
      aiRetrieve(
        { scope: { accountIds: ['stub-a'] }, op: 'triage-headers', query: 'invoice' },
        deps,
      ),
    ).rejects.toThrow('mail: rate-limited')
  })

  it('enforces the shared bucket over the channel', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, createStubAiDeps(twoAccountStubSeeds()))
    for (let i = 0; i < 10; i += 1) {
      await invoke(MAIL_CHANNELS.aiRetrieve, {
        scope: { accountIds: ['stub-a'] },
        op: 'triage-headers',
        query: 'invoice',
      })
    }
    await expect(
      invoke(MAIL_CHANNELS.aiRetrieve, {
        scope: { accountIds: ['stub-a'] },
        op: 'triage-headers',
        query: 'invoice',
      }),
    ).rejects.toThrow(/mail: rate-limited, retry in \d+s/)
  })
})

describe('H4 retrieval and redaction throughput budgets', () => {
  function perfSeeds(size: number): ReturnType<typeof twoAccountStubSeeds> {
    const messages = Array.from({ length: size }, (_, i) => ({
      uid: `p${i}`,
      folderId: 'inbox',
      from: ['billing@example.com'],
      subject: `Invoice record ${i} alpha`,
      snippet: `Snippet for invoice record ${i} alpha.`,
      text: `Body for invoice record ${i} alpha with routine terms.`,
    }))
    return [{ accountId: 'perf', messages }]
  }

  function hostileBody(i: number): string {
    return [
      'Received: from mx1.example.com (mx1.example.com [192.0.2.10])',
      'DKIM-Signature: v=1; a=rsa-sha256; d=example.com; s=sel; b=abcd',
      `Record ${i}: bearer abcdefghij1234567890 password=hunter2 2001:db8::${i % 997}`,
      '-----BEGIN PRIVATE KEY-----',
      'MIIEvwIBADANBgkqhkiG9w0BAQEFAASC',
      '-----END PRIVATE KEY-----',
    ].join('\n')
  }

  it('P-7 triage-headers over 20k messages stays under 300 ms', async () => {
    const deps = isolatedDeps(perfSeeds(20_000))
    await aiRetrieve(
      { scope: { accountIds: ['perf'] }, op: 'triage-headers', query: 'alpha', limit: 50 },
      deps,
    )
    const samples: number[] = []
    for (let i = 0; i < 6; i += 1) {
      const start = performance.now()
      const response = await aiRetrieve(
        { scope: { accountIds: ['perf'] }, op: 'triage-headers', query: 'alpha', limit: 50 },
        deps,
      )
      samples.push(performance.now() - start)
      expect(response.envelopes).toHaveLength(50)
    }
    expect(Math.max(...samples)).toBeLessThan(300)
  })

  it('P-7 fetch-bodies for 50 hostile refs stays under 1 s with redaction', async () => {
    const messages = Array.from({ length: 50 }, (_, i) => ({
      uid: `h${i}`,
      folderId: 'inbox',
      from: ['billing@example.com'],
      subject: `Hostile record ${i}`,
      snippet: `Snippet hostile ${i}.`,
      text: hostileBody(i),
    }))
    const deps = isolatedDeps([{ accountId: 'perf', messages }])
    const refs = messages.map((message) => stubRef(message.folderId, message.uid))
    const start = performance.now()
    const response = await aiRetrieve(
      { scope: { accountIds: ['perf'] }, op: 'fetch-bodies', refs },
      deps,
    )
    expect(performance.now() - start).toBeLessThan(1000)
    expect(response.bodies).toHaveLength(50)
    for (const body of response.bodies ?? []) {
      expect(containsPii(body.text)).toBe(false)
    }
  })

  it('P-7 cite-resolve for 50 refs stays under 300 ms', async () => {
    const deps = isolatedDeps(perfSeeds(5000))
    const refs = Array.from({ length: 50 }, (_, i) => stubRef('inbox', `p${i}`))
    const start = performance.now()
    const response = await aiRetrieve(
      { scope: { accountIds: ['perf'] }, op: 'cite-resolve', refs },
      deps,
    )
    expect(performance.now() - start).toBeLessThan(300)
    expect(response.citations).toHaveLength(50)
  })

  it('P-8 redactForPrompt over a 200k hostile input stays under 250 ms', () => {
    const input = `${hostileBody(7)}\n`.repeat(9000).slice(0, 200_000)
    expect(input.length).toBe(200_000)
    const start = performance.now()
    const clean = redactForPrompt(input)
    expect(performance.now() - start).toBeLessThan(250)
    expect(containsPii(clean)).toBe(false)
  })

  it('P-8 redactForPrompt over a 1 MiB body stays under 1 s and marks the cap', () => {
    const input = `${hostileBody(11)}\n`.repeat(6000).slice(0, 1024 * 1024)
    expect(input.length).toBe(1024 * 1024)
    const start = performance.now()
    const clean = redactForPrompt(input)
    expect(performance.now() - start).toBeLessThan(1000)
    expect(clean.endsWith('[truncated]')).toBe(true)
    expect(clean.length).toBeLessThanOrEqual(PROMPT_TEXT_CAP_CHARS + '[truncated]'.length)
    expect(containsPii(clean)).toBe(false)
  })
})

describe('H4 P-6 store search quoting and LIKE fallback', () => {
  function sqliteStore(): {
    upsertMessageFull(input: Record<string, unknown>): void
    searchLocal(query: string, opts?: { limit?: number }): { ref: { uid: string } }[]
  } {
    const dir = mkdtempSync(join(tmpdir(), 'mail-h4-'))
    configureMailStoreDir(dir)
    const store = getMailStore('demo') as unknown as {
      upsertMessageFull(input: Record<string, unknown>): void
      searchLocal(query: string, opts?: { limit?: number }): { ref: { uid: string } }[]
    }
    if (typeof store.upsertMessageFull !== 'function' || typeof store.searchLocal !== 'function') {
      throw new Error('expected a SQLite mail store')
    }
    return store
  }

  it('keeps FTS MATCH quoting mandatory for hostile queries', () => {
    const store = sqliteStore()
    store.upsertMessageFull({
      folderId: 'inbox',
      uid: 'q1',
      threadId: 'thread-q1',
      from: ['billing@example.com'],
      to: ['owner@example.com'],
      subject: 'Plain invoice q1',
      date: '2026-10-01T09:00:00.000Z',
      bodyText: 'Routine body text.',
    })
    const start = performance.now()
    // All queries carry alnum tokens (quoted by buildFtsMatch), so MATCH
    // stays valid: user syntax never reaches the engine raw.
    for (const query of ['"invoice" OR 1=1', 'inbox* MATCH *', 'a AND b OR c']) {
      expect(() => store.searchLocal(query, { limit: 50 })).not.toThrow()
    }
    expect(performance.now() - start).toBeLessThan(1000)
  })

  it('escapes LIKE metacharacters so they match literally', () => {
    const store = sqliteStore()
    const base = {
      folderId: 'inbox',
      threadId: 'thread-like',
      from: ['billing@example.com'],
      to: ['owner@example.com'],
      date: '2026-10-01T09:00:00.000Z',
      bodyText: 'Routine body.',
    }
    store.upsertMessageFull({ ...base, uid: 'pct', subject: '100% real invoice' })
    store.upsertMessageFull({ ...base, uid: 'plain', subject: 'ordinary invoice' })
    store.upsertMessageFull({ ...base, uid: 'mid', subject: 'reference ZZQX99 filed' })
    // Percent stays literal: only the row containing it matches.
    expect(store.searchLocal('100%', { limit: 50 }).map((env) => env.ref.uid)).toEqual(['pct'])
    // Mid-token Latin forces the LIKE fallback (no FTS prefix hit).
    expect(store.searchLocal('QX99', { limit: 50 }).map((env) => env.ref.uid)).toEqual(['mid'])
    expect(store.searchLocal('ordinary', { limit: 50 }).map((env) => env.ref.uid)).toEqual([
      'plain',
    ])
  })

  it('serves mid-token CJK queries through the LIKE fallback in under 1 s', () => {
    const store = sqliteStore()
    store.upsertMessageFull({
      folderId: 'inbox',
      uid: 'cjk1',
      threadId: 'thread-cjk1',
      from: ['news@example.org'],
      to: ['owner@example.com'],
      subject: '週報まとめ 10月第2週',
      snippet: '今週の週報まとめです。',
      date: '2026-10-01T09:00:00.000Z',
      bodyText: '週報まとめの本文です。',
    })
    const start = performance.now()
    const hits = store.searchLocal('報ま', { limit: 50 })
    expect(performance.now() - start).toBeLessThan(1000)
    expect(hits.map((env) => env.ref.uid)).toContain('cjk1')
  })
})
