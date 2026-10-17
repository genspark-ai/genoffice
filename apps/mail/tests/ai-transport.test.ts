import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { beforeEach, describe, expect, it } from 'vitest'
import { MAIL_CHANNELS } from '../src/shared/ipc'
import type { AiRetrieveResponse } from '../src/shared/ipc'
import { mailAiRetrieveSchema } from '../src/shared/ipc-schemas'
import type { ComposeIpcBus } from '../src/main/compose-handlers'
import {
  collectStubStream,
  createStubAiDeps,
  isScopeStale,
  normalizeAiScope,
  registerAiRetrieveHandlers,
  resetAiRetrieveHandlers,
  resetSharedAiRetrieveRateLimiter,
  stubRef,
  twoAccountStubSeeds,
} from '../src/main/ai'
import { defaultAiRetrieveDeps } from '../src/main/ai'
import { resetMailDemoSeed } from '../src/main/mail-read-handlers'
import { configureMailStoreDir, getMailStore, resetMailStoreConfig } from '../src/main/mail-store'
import { closeAllMailStores } from '../src/main/store/store'
import {
  containsPii,
  redactForPrompt,
  sanitizeError,
  stripTransportHeaders,
} from '../src/main/sync/redact'

/**
 * Lane W AI transport tests: the single mail:ai-retrieve channel over a
 * fake bus, two-account no-leak fixtures, prompt redaction with hostile
 * mail fixtures, the stubbed streaming fixture, and schema validation
 * rejects. Synthetic example.com mail only — never real mailboxes.
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

beforeEach(() => {
  resetAiRetrieveHandlers()
  resetSharedAiRetrieveRateLimiter()
  resetMailDemoSeed()
  closeAllMailStores()
  resetMailStoreConfig()
})

describe('mail:ai-retrieve channel', () => {
  it('is the sole new mail:ai-* channel', () => {
    expect(MAIL_CHANNELS.aiRetrieve).toBe('mail:ai-retrieve')
  })

  it('triages one account without leaking the other', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, createStubAiDeps(twoAccountStubSeeds()))
    const response = (await invoke(MAIL_CHANNELS.aiRetrieve, {
      scope: { accountIds: ['stub-a'] },
      op: 'triage-headers',
      query: 'invoice',
    })) as AiRetrieveResponse
    expect(response.scope).toEqual({ accountIds: ['stub-a'] })
    expect(response.envelopes).toHaveLength(1)
    expect(response.envelopes[0].subject).toContain('INV-A1')
    expect(response.citations).toHaveLength(1)
    expect(response.citations[0].accountId).toBe('stub-a')
    expect(response.bodies).toBeUndefined()
  })

  it('returns nothing when the query only matches outside the scope', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, createStubAiDeps(twoAccountStubSeeds()))
    const response = (await invoke(MAIL_CHANNELS.aiRetrieve, {
      scope: { accountIds: ['stub-a'] },
      op: 'triage-headers',
      query: 'appointment',
    })) as AiRetrieveResponse
    expect(response.envelopes).toEqual([])
    expect(response.citations).toEqual([])
  })

  it('fetches bodies for the cited set only and redacts prompt secrets', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(
      bus,
      createStubAiDeps([
        {
          accountId: 'stub-a',
          messages: [
            {
              uid: 'a9',
              folderId: 'inbox',
              from: ['billing@example.com'],
              subject: 'Receipt',
              snippet: 'Your receipt is attached.',
              text: 'Receipt attached. bearer abcdefghij1234567890 from 10.1.2.3, thanks.',
            },
          ],
        },
        ...twoAccountStubSeeds().filter((seed) => seed.accountId === 'stub-b'),
      ]),
    )
    const response = (await invoke(MAIL_CHANNELS.aiRetrieve, {
      scope: { accountIds: ['stub-a'] },
      op: 'fetch-bodies',
      refs: [stubRef('inbox', 'a9')],
    })) as AiRetrieveResponse
    expect(response.bodies).toHaveLength(1)
    expect(response.bodies?.[0].accountId).toBe('stub-a')
    expect(response.bodies?.[0].text).not.toContain('bearer abcdefghij1234567890')
    expect(response.bodies?.[0].text).not.toContain('10.1.2.3')
    expect(containsPii(response.bodies?.[0].text ?? 'x')).toBe(false)
  })

  it('omits refs that resolve only outside the scope', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, createStubAiDeps(twoAccountStubSeeds()))
    const response = (await invoke(MAIL_CHANNELS.aiRetrieve, {
      scope: { accountIds: ['stub-a'] },
      op: 'fetch-bodies',
      refs: [stubRef('inbox', 'b1')],
    })) as AiRetrieveResponse
    expect(response.envelopes).toEqual([])
    expect(response.bodies).toEqual([])
    expect(response.citations).toEqual([])
  })

  it('resolves citations header-only without bodies', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, createStubAiDeps(twoAccountStubSeeds()))
    const response = (await invoke(MAIL_CHANNELS.aiRetrieve, {
      scope: { accountIds: ['stub-b'] },
      op: 'cite-resolve',
      refs: [stubRef('inbox', 'b1')],
    })) as AiRetrieveResponse
    expect(response.citations).toEqual([
      { uid: 'b1', folderId: 'inbox', msgIdHeader: '<b1@example.net>', accountId: 'stub-b' },
    ])
    expect(response.envelopes).toHaveLength(1)
    expect(response.bodies).toBeUndefined()
  })

  it('rejects unknown scope accounts before opening any store', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, createStubAiDeps(twoAccountStubSeeds()))
    await expect(
      invoke(MAIL_CHANNELS.aiRetrieve, {
        scope: { accountIds: ['nope'] },
        op: 'triage-headers',
        query: 'invoice',
      }),
    ).rejects.toThrow('mail:')
  })

  it('rejects malformed payloads with mail errors', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, createStubAiDeps(twoAccountStubSeeds()))
    await expect(
      invoke(MAIL_CHANNELS.aiRetrieve, { scope: { accountIds: [] }, op: 'triage-headers' }),
    ).rejects.toThrow('mail:')
    await expect(
      invoke(MAIL_CHANNELS.aiRetrieve, {
        scope: { accountIds: ['stub-a'] },
        op: 'triage-headers',
      }),
    ).rejects.toThrow('mail:')
    await expect(
      invoke(MAIL_CHANNELS.aiRetrieve, {
        scope: { accountIds: ['stub-a'] },
        op: 'fetch-bodies',
      }),
    ).rejects.toThrow('mail:')
    await expect(
      invoke(MAIL_CHANNELS.aiRetrieve, { scope: { accountIds: ['stub-a'] }, op: 'nope' }),
    ).rejects.toThrow('mail:')
  })

  it('serves the demo seed through the production default deps', async () => {
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, defaultAiRetrieveDeps())
    const response = (await invoke(MAIL_CHANNELS.aiRetrieve, {
      scope: { accountIds: ['demo'] },
      op: 'triage-headers',
      query: 'invoice',
    })) as AiRetrieveResponse
    expect(response.envelopes.map((env) => env.ref.uid)).toContain('m1')
    for (const env of response.envelopes) {
      expect(env).not.toHaveProperty('body')
      expect(env).not.toHaveProperty('bodyText')
    }
  })

  it('serves the real SQLite store when one is configured', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'mail-ai-'))
    configureMailStoreDir(dir)
    const store = getMailStore('demo')
    assertSqliteStore(store).upsertMessageFull({
      folderId: 'inbox',
      uid: 'sq1',
      threadId: 'thread-sq1',
      from: ['billing@example.com'],
      to: ['owner@example.com'],
      subject: 'SQLite invoice sq1',
      date: '2026-10-01T09:00:00.000Z',
      bodyText: 'Invoice sq1 body with password=hunter2 inside.',
    })
    const { bus, invoke } = makeBus()
    registerAiRetrieveHandlers(bus, defaultAiRetrieveDeps())
    const triage = (await invoke(MAIL_CHANNELS.aiRetrieve, {
      scope: { accountIds: ['demo'] },
      op: 'triage-headers',
      query: 'sq1',
    })) as AiRetrieveResponse
    expect(triage.envelopes.map((env) => env.ref.uid)).toContain('sq1')
    const bodies = (await invoke(MAIL_CHANNELS.aiRetrieve, {
      scope: { accountIds: ['demo'] },
      op: 'fetch-bodies',
      refs: [stubRef('inbox', 'sq1')],
    })) as AiRetrieveResponse
    expect(bodies.bodies).toHaveLength(1)
    expect(bodies.bodies?.[0].text).not.toContain('hunter2')
  })
})

function assertSqliteStore(store: unknown): {
  upsertMessageFull(input: {
    folderId: string
    uid: string
    threadId: string
    from: string[]
    to: string[]
    subject: string
    date: string
    bodyText: string
  }): void
} {
  const candidate = store as {
    upsertMessageFull?: (input: Record<string, unknown>) => void
  }
  if (typeof candidate.upsertMessageFull !== 'function') {
    throw new Error('expected a SQLite mail store')
  }
  return candidate as {
    upsertMessageFull(input: {
      folderId: string
      uid: string
      threadId: string
      from: string[]
      to: string[]
      subject: string
      date: string
      bodyText: string
    }): void
  }
}

describe('AI scope envelope', () => {
  it('normalizes scopes and rejects violations', () => {
    expect(normalizeAiScope({ accountIds: ['a', 'a', 'b'] })).toEqual({ accountIds: ['a', 'b'] })
    expect(() => normalizeAiScope({ accountIds: [] })).toThrow('mail:')
    expect(() =>
      normalizeAiScope({ accountIds: Array.from({ length: 9 }, (_, i) => `a${i}`) }),
    ).toThrow('mail:')
    expect(() => normalizeAiScope({ accountIds: [''] })).toThrow('mail:')
    expect(() => normalizeAiScope(null)).toThrow('mail:')
  })

  it('detects stale scopes order-insensitively', () => {
    expect(isScopeStale({ accountIds: ['a'] }, { accountIds: ['a'] })).toBe(false)
    expect(isScopeStale({ accountIds: ['a', 'b'] }, { accountIds: ['b', 'a'] })).toBe(false)
    expect(isScopeStale({ accountIds: ['a'] }, { accountIds: ['b'] })).toBe(true)
    expect(isScopeStale({ accountIds: ['a'] }, { accountIds: ['a', 'b'] })).toBe(true)
  })
})

describe('prompt redaction', () => {
  const hostile = [
    'Received: from mx1.example.com (mx1.example.com [192.0.2.10])',
    '  by mail.example.com with ESMTP id abc123;',
    '  Tue, 14 Oct 2026 10:00:00 +0000',
    'Authentication-Results: mail.example.com; dkim=pass header.i=@example.com',
    'DKIM-Signature: v=1; a=rsa-sha256; d=example.com; s=sel; b=abcd',
    'From: sender@example.com',
    'Subject: invoice INV-1042',
    '',
    'Please pay INV-1042. bearer abcdefghij1234567890, password=hunter2.',
    'Gateway 2001:db8::1 logged the request.',
    '-----BEGIN PRIVATE KEY-----',
    'MIIEvwIBADANBgkqhkiG9w0BAQEFAASCAmkwggJfAgEAAoIBAQDLONGKEYDATA',
    '-----END PRIVATE KEY-----',
  ].join('\n')

  it('strips transport internals, tokens, passwords, IPs, and key material', () => {
    const clean = redactForPrompt(hostile)
    expect(clean).not.toContain('Received:')
    expect(clean).not.toContain('Authentication-Results:')
    expect(clean).not.toContain('192.0.2.10')
    expect(clean).not.toContain('2001:db8::1')
    expect(clean).not.toContain('bearer abcdefghij1234567890')
    expect(clean).not.toContain('hunter2')
    expect(clean).not.toContain('BEGIN PRIVATE KEY')
    expect(clean).not.toContain('sender@example.com')
    expect(clean).toContain('INV-1042')
    expect(containsPii(clean)).toBe(false)
  })

  it('drops folded continuation lines of removed headers but keeps the body', () => {
    const stripped = stripTransportHeaders(hostile)
    expect(stripped).not.toMatch(/^received:/gim)
    expect(stripped).not.toContain('by mail.example.com')
    expect(stripped).toContain('Subject: invoice INV-1042')
    expect(stripped).toContain('Please pay INV-1042.')
  })

  it('keeps the error path PII-free and stable', () => {
    const clean = sanitizeError(new Error('sync failed for reader@example.com at 10.0.0.5'))
    expect(clean).not.toContain('reader@example.com')
    expect(clean).not.toContain('10.0.0.5')
    expect(containsPii(clean)).toBe(false)
    expect(containsPii('nothing sensitive here')).toBe(false)
  })
})

describe('stubbed provider streaming', () => {
  it('emits ordered deltas plus a terminal done chunk', () => {
    const chunks = collectStubStream(
      'req-1',
      'Triage digest for the inbox with several words to force chunking.',
    )
    expect(chunks.length).toBeGreaterThan(2)
    for (const chunk of chunks) expect(chunk.requestId).toBe('req-1')
    const deltas = chunks.filter((chunk) => chunk.type === 'delta')
    expect(deltas.length).toBeGreaterThan(1)
    expect(chunks[chunks.length - 1]).toEqual({ requestId: 'req-1', type: 'done' })
    expect(deltas.map((chunk) => chunk.text).join(' ')).toContain('Triage digest')
  })
})

describe('mailAiRetrieveSchema', () => {
  const ref = { uid: 'm1', folderId: 'inbox', threadId: 'thread-m1' }

  it('accepts all three ops with their required fields', () => {
    expect(
      mailAiRetrieveSchema.safeParse({
        scope: { accountIds: ['demo'] },
        op: 'triage-headers',
        query: 'invoice',
      }).success,
    ).toBe(true)
    expect(
      mailAiRetrieveSchema.safeParse({
        scope: { accountIds: ['demo'] },
        op: 'fetch-bodies',
        refs: [ref],
      }).success,
    ).toBe(true)
    expect(
      mailAiRetrieveSchema.safeParse({
        scope: { accountIds: ['demo'] },
        op: 'cite-resolve',
        refs: [ref],
      }).success,
    ).toBe(true)
  })

  it('rejects scope violations, missing op fields, and unknown keys', () => {
    expect(
      mailAiRetrieveSchema.safeParse({
        scope: { accountIds: [] },
        op: 'triage-headers',
        query: 'x',
      }).success,
    ).toBe(false)
    expect(
      mailAiRetrieveSchema.safeParse({
        scope: { accountIds: Array.from({ length: 9 }, (_, i) => `a${i}`) },
        op: 'triage-headers',
        query: 'x',
      }).success,
    ).toBe(false)
    expect(
      mailAiRetrieveSchema.safeParse({ scope: { accountIds: ['a'] }, op: 'triage-headers' })
        .success,
    ).toBe(false)
    expect(
      mailAiRetrieveSchema.safeParse({ scope: { accountIds: ['a'] }, op: 'fetch-bodies' }).success,
    ).toBe(false)
    expect(
      mailAiRetrieveSchema.safeParse({
        scope: { accountIds: ['a'] },
        op: 'triage-headers',
        query: 'x',
        limit: 500,
      }).success,
    ).toBe(false)
    expect(
      mailAiRetrieveSchema.safeParse({
        scope: { accountIds: ['a'] },
        op: 'triage-headers',
        query: 'x',
        bogus: 1,
      }).success,
    ).toBe(false)
  })
})
