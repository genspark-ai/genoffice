import { describe, expect, it } from 'vitest'
import { composeSkills, type AgentToolCall } from '@genoffice/agent-core'
import type { MailEnvelope } from '@genoffice/mail-protocols'
import {
  MAIL_AGENT_TOOLS,
  type AiRetrieveRequest,
  type AiRetrieveResponse,
  type MailAiBridge,
} from '../src/renderer/ai/mail-tools'
import { createTriageSkill } from '../src/renderer/ai/triage-skill'
import { createSummarizeSkill } from '../src/renderer/ai/summarize-skill'
import { createComposeAssistSkill } from '../src/renderer/ai/compose-assist-skill'
import { createMailSearchSkill } from '../src/renderer/ai/search-skill'

/**
 * K-lane skill tests over a spec-faithful fake `aiRetrieve` bridge
 * (two synthetic accounts, disjoint mail, scope enforced). The fake lives
 * in this file on purpose: K lane must not touch devFakes.ts (UI lane owns
 * it). All fixtures are synthetic; body markers never appear in snippets.
 */

interface SeedMessage {
  uid: string
  folderId: string
  from: string
  subject: string
  date: string
  seen?: boolean
  flagged?: boolean
  labels?: string[]
  snippet: string
  body: string
  msgIdHeader?: string
}

const ACCOUNT_A = 'acct-a'
const ACCOUNT_B = 'acct-b'

const SEEDS: Record<string, SeedMessage[]> = {
  [ACCOUNT_A]: [
    {
      uid: 'a-m1',
      folderId: 'inbox',
      from: 'Acme Billing <billing@example.com>',
      subject: 'Invoice INV-1042 ready for review',
      date: '2026-10-10T09:00:00.000Z',
      flagged: true,
      snippet: 'Invoice INV-1042 for October is attached for review.',
      body: 'BODY-ONLY-A-M1: full invoice terms, bank account 88112233, due Friday.',
      msgIdHeader: '<a-m1@example.com>',
    },
    {
      uid: 'a-m2',
      folderId: 'inbox',
      from: 'Weekly Deals <deals@example.com>',
      subject: 'Weekly newsletter: October deals digest',
      date: '2026-10-11T12:30:00.000Z',
      snippet: 'This week: storage promos. Unsubscribe any time.',
      body: 'BODY-ONLY-A-M2: promo codes and tracking pixels live here.',
    },
    {
      uid: 'a-m3',
      folderId: 'inbox',
      from: 'Team Chat <chat@example.com>',
      subject: 'Notes from Monday standup',
      date: '2026-10-12T08:15:00.000Z',
      seen: true,
      snippet: 'Shipped the new reader; remote images stay blocked.',
      body: 'BODY-ONLY-A-M3: standup transcript with action items.',
    },
  ],
  [ACCOUNT_B]: [
    {
      uid: 'b-m1',
      folderId: 'inbox',
      from: 'Payroll <payroll@example.net>',
      subject: 'Payslip September is available',
      date: '2026-10-10T10:00:00.000Z',
      snippet: 'Your September payslip is ready to download.',
      body: 'BODY-ONLY-B-M1: salary breakdown and tax withholding figures.',
      msgIdHeader: '<b-m1@example.net>',
    },
    {
      uid: 'b-m2',
      folderId: 'archive',
      from: 'Support <support@example.net>',
      subject: 'Ticket 4419 resolved',
      date: '2026-10-09T07:00:00.000Z',
      seen: true,
      snippet: 'Ticket 4419 was resolved yesterday.',
      body: 'BODY-ONLY-B-M2: internal support notes and agent roster.',
    },
  ],
}

function toEnvelope(seed: SeedMessage): MailEnvelope {
  return {
    ref: { uid: seed.uid, folderId: seed.folderId, threadId: `thread-${seed.uid}` },
    from: [seed.from],
    to: ['Reader <reader@example.org>'],
    cc: [],
    subject: seed.subject,
    date: seed.date,
    flags: {
      seen: seed.seen ?? false,
      flagged: seed.flagged ?? false,
      answered: false,
      draft: false,
    },
    labels: seed.labels ?? [],
    snippet: seed.snippet,
    hasAttach: false,
  }
}

function matches(text: string, query: string): boolean {
  const hay = text.toLowerCase()
  return query
    .toLowerCase()
    .split(/\s+/)
    .filter(Boolean)
    .every((token) => hay.includes(token))
}

/** Spec-faithful fake: rejects unknown accounts, filters by scope, echoes scope. */
class FakeAiBridge implements MailAiBridge {
  readonly calls: AiRetrieveRequest[] = []
  /** When set, the echo lies (stale-scope discard test). */
  echoOverride?: string[]

  async aiRetrieve(request: AiRetrieveRequest): Promise<AiRetrieveResponse> {
    this.calls.push(request)
    for (const id of request.scope.accountIds) {
      if (!SEEDS[id]) throw new Error(`unknown account: ${id}`)
    }
    const echo = this.echoOverride ?? request.scope.accountIds
    const inScope = request.scope.accountIds.flatMap((id) => SEEDS[id] ?? [])
    const ownerOf = (uid: string): string | undefined =>
      Object.entries(SEEDS).find(([, messages]) => messages.some((m) => m.uid === uid))?.[0]

    if (request.op === 'triage-headers') {
      const query = request.query ?? ''
      const hits = inScope.filter(
        (m) =>
          (!request.folderId || m.folderId === request.folderId) &&
          matches(`${m.subject} ${m.snippet} ${m.from}`, query),
      )
      const limited = hits.slice(0, request.limit ?? 50)
      return {
        scope: { accountIds: echo },
        envelopes: limited.map(toEnvelope),
        citations: limited.map((m) => ({ uid: m.uid, folderId: m.folderId })),
      }
    }

    const refs = request.refs ?? []
    const resolved = refs.filter((ref) => {
      const owner = ownerOf(ref.uid)
      return owner !== undefined && request.scope.accountIds.includes(owner)
    })
    if (request.op === 'fetch-bodies') {
      const byKey = new Map(inScope.map((m) => [`${m.folderId}/${m.uid}`, m]))
      const bodies = resolved.flatMap((ref) => {
        const seed = byKey.get(`${ref.folderId}/${ref.uid}`)
        return seed ? [{ ref, text: seed.body }] : []
      })
      return {
        scope: { accountIds: echo },
        envelopes: resolved.map((ref) => toEnvelope(byKey.get(`${ref.folderId}/${ref.uid}`)!)),
        bodies,
        citations: resolved.map((ref) => ({ uid: ref.uid, folderId: ref.folderId })),
      }
    }
    // cite-resolve: header-only, no bodies.
    return {
      scope: { accountIds: echo },
      envelopes: resolved.map((ref) => {
        const seed = inScope.find((m) => m.uid === ref.uid)!
        return toEnvelope(seed)
      }),
      citations: resolved.map((ref) => {
        const seed = inScope.find((m) => m.uid === ref.uid)!
        return {
          uid: ref.uid,
          folderId: ref.folderId,
          ...(seed.msgIdHeader ? { msgIdHeader: seed.msgIdHeader } : {}),
        }
      }),
    }
  }
}

function call(name: string, input: Record<string, unknown> = {}): AgentToolCall {
  return { id: `call-${name}`, name, input }
}

const scopeA = { accountIds: [ACCOUNT_A] }

describe('triage skill (headers + FTS only)', () => {
  it('ranks a folder into buckets without ever surfacing bodies', async () => {
    const bridge = new FakeAiBridge()
    const skill = createTriageSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    const result = await skill.executeTool(
      call('rank_folder', { query: 'invoice October', limit: 10 }),
    )
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('## priority')
    expect(result.output).toContain('[a-m1]')
    expect(result.output).not.toContain('BODY-ONLY-A-M1')
    expect(result.output).not.toContain('88112233')
    expect(bridge.calls[0]?.op).toBe('triage-headers')
    expect(bridge.calls[0]?.scope).toEqual(scopeA)
  })

  it('buckets newsletters and seen mail away from needs-reply', async () => {
    const bridge = new FakeAiBridge()
    const skill = createTriageSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    const bulk = await skill.executeTool(call('rank_folder', { query: 'newsletter' }))
    expect(bulk.output).toContain('## newsletter')
    expect(bulk.output).toContain('[a-m2]')
    const seen = await skill.executeTool(call('rank_folder', { query: 'standup' }))
    expect(seen.output).toContain('## noise')
    expect(seen.output).toContain('[a-m3]')
    expect(seen.output).toContain('## needs-reply (0)')
  })

  it('attaches the header snapshot to context, capped', () => {
    const snapshot: MailEnvelope[] = Array.from({ length: 70 }, (_, i) => ({
      ref: { uid: `a-x${i}`, folderId: 'inbox', threadId: `thread-a-x${i}` },
      from: ['Bulk <bulk@example.com>'],
      to: [],
      cc: [],
      subject: `Bulk message ${i}`,
      date: '2026-10-10T00:00:00.000Z',
      flags: { seen: true, flagged: false, answered: false, draft: false },
      labels: [],
      snippet: 'Bulk snippet.',
      hasAttach: false,
    }))
    const skill = createTriageSkill({ getScope: () => scopeA, getSnapshot: () => snapshot })
    const context = skill.buildContext?.() ?? ''
    expect(context).toContain('## Folder headers')
    expect(context).toContain('use rank_folder to narrow')
    expect(createTriageSkill({ getScope: () => scopeA }).buildContext?.()).toBe('')
  })

  it('apply_label proposes without mutating or calling the bridge', async () => {
    const bridge = new FakeAiBridge()
    const skill = createTriageSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    const result = await skill.executeTool(
      call('apply_label', { uids: ['a-m1'], folderId: 'inbox', label: 'follow-up' }),
    )
    expect(result.isError).toBeFalsy()
    expect(result.mutated).toBeFalsy()
    expect(result.output).toContain('NOT applied')
    expect(result.output).toContain('[a-m1]')
    expect(bridge.calls).toHaveLength(0)
  })
})

describe('scope guards', () => {
  it('rejects empty and oversized scopes', async () => {
    const bridge = new FakeAiBridge()
    const empty = createTriageSkill({
      getScope: () => ({ accountIds: [] }),
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    expect((await empty.executeTool(call('rank_folder', { query: 'x' }))).isError).toBe(true)
    const wide = createTriageSkill({
      getScope: () => ({ accountIds: Array.from({ length: 9 }, (_, i) => `acct-${i}`) }),
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    expect((await wide.executeTool(call('rank_folder', { query: 'x' }))).isError).toBe(true)
  })

  it('requires query for triage-headers and refs for body/cite ops', async () => {
    const bridge = new FakeAiBridge()
    const triage = createTriageSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    expect((await triage.executeTool(call('rank_folder', { query: '  ' }))).isError).toBe(true)
    const summarize = createSummarizeSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    expect(
      (await summarize.executeTool(call('summarize_thread', { folderId: 'inbox' }))).isError,
    ).toBe(true)
    const search = createMailSearchSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    expect((await search.executeTool(call('search_mail', {}))).isError).toBe(true)
  })

  it('discards stale-scope echoes instead of rendering them', async () => {
    const bridge = new FakeAiBridge()
    bridge.echoOverride = [ACCOUNT_B]
    const skill = createTriageSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    const result = await skill.executeTool(call('rank_folder', { query: 'invoice' }))
    expect(result.isError).toBe(true)
    expect(result.output).toContain('stale scope')
  })

  it('reports an unavailable bridge instead of failing silently', async () => {
    const skill = createTriageSkill({ getScope: () => scopeA })
    const result = await skill.executeTool(call('rank_folder', { query: 'invoice' }))
    expect(result.isError).toBe(true)
    expect(result.output).toContain('unavailable')
  })

  it('surfaces unknown-account rejection from the bridge', async () => {
    const bridge = new FakeAiBridge()
    const skill = createTriageSkill({
      getScope: () => ({ accountIds: ['ghost'] }),
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    const result = await skill.executeTool(call('rank_folder', { query: 'invoice' }))
    expect(result.isError).toBe(true)
    expect(result.output).toContain('unknown account')
  })
})

describe('two-account no-leak fixtures', () => {
  it('never surfaces B envelopes in an A-scoped search', async () => {
    const bridge = new FakeAiBridge()
    const skill = createMailSearchSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    const result = await skill.executeTool(
      call('search_mail', { query: 'available resolved payslip ticket' }),
    )
    expect(result.isError).toBeFalsy()
    expect(result.output).not.toContain('[b-m1]')
    expect(result.output).not.toContain('[b-m2]')
    expect(bridge.calls[0]?.scope).toEqual(scopeA)
  })

  it('never fetches B bodies under an A scope', async () => {
    const bridge = new FakeAiBridge()
    const skill = createSummarizeSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    const result = await skill.executeTool(
      call('summarize_thread', { uids: ['b-m1'], folderId: 'inbox' }),
    )
    expect(result.isError).toBeFalsy()
    expect(result.output).not.toContain('BODY-ONLY-B-M1')
    expect(result.output).not.toContain('[b-m1]')
  })

  it('never quotes B messages under an A scope', async () => {
    const bridge = new FakeAiBridge()
    const skill = createMailSearchSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    const result = await skill.executeTool(
      call('quote_message', { uids: ['b-m1'], folderId: 'inbox' }),
    )
    expect(result.isError).toBeFalsy()
    expect(result.output).not.toContain('BODY-ONLY-B-M1')
    expect(result.output).not.toContain('Payslip')
  })
})

describe('summarize skill (cited bullets)', () => {
  it('returns cited bodies and accepts a cited reply', async () => {
    const bridge = new FakeAiBridge()
    const skill = createSummarizeSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    const result = await skill.executeTool(
      call('summarize_thread', { uids: ['a-m1'], folderId: 'inbox' }),
    )
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('[a-m1]')
    expect(result.output).toContain('BODY-ONLY-A-M1')
    expect(
      skill.verifyResponse?.('- Invoice terms are due Friday [a-m1]', [
        { name: 'summarize_thread', ok: true },
      ]),
    ).toBeNull()
  })

  it('forces a corrective turn when claims lack citations', async () => {
    const bridge = new FakeAiBridge()
    const skill = createSummarizeSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    await skill.executeTool(call('summarize_thread', { uids: ['a-m1'], folderId: 'inbox' }))
    const correction = skill.verifyResponse?.('- Invoice terms are due Friday with no citation', [
      { name: 'summarize_thread', ok: true },
    ])
    expect(correction).toContain('[a-m1]')
    expect(
      skill.verifyResponse?.('Here is your summary, let me know if you want more.', [
        { name: 'summarize_thread', ok: true },
      ]),
    ).toBeNull()
  })

  it('catches up on a folder with cited top bodies', async () => {
    const bridge = new FakeAiBridge()
    const skill = createSummarizeSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    const result = await skill.executeTool(
      call('catch_up_folder', { query: 'October', folderId: 'inbox', limit: 10 }),
    )
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('Triage digest')
    expect(result.output).toContain('[a-m1]')
    expect(bridge.calls.map((c) => c.op)).toEqual(['triage-headers', 'fetch-bodies'])
  })
})

describe('compose-assist skill (editable text, never sends)', () => {
  it('assembles drafts from bullets without touching the bridge', async () => {
    const bridge = new FakeAiBridge()
    const skill = createComposeAssistSkill({
      getScope: () => scopeA,
      getDraft: () => 'Thanks for the invoice.',
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    expect(skill.buildContext?.()).toContain('Thanks for the invoice.')
    const result = await skill.executeTool(
      call('draft_from_bullets', { bullets: ['Confirm receipt', 'Ask for a W-8 form'] }),
    )
    expect(result.isError).toBeFalsy()
    expect(result.mutated).toBeFalsy()
    expect(result.output).toContain('- Confirm receipt')
    expect(result.output).toContain('Best regards,')
    expect(bridge.calls).toHaveLength(0)
  })

  it('validates tone rewrites and keeps the source text in scope', async () => {
    const bridge = new FakeAiBridge()
    const skill = createComposeAssistSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    expect(
      (await skill.executeTool(call('rewrite_tone', { text: 'hi', tone: 'royal' }))).isError,
    ).toBe(true)
    const result = await skill.executeTool(
      call('rewrite_tone', { text: 'Send it now.', tone: 'formal' }),
    )
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('formal')
    expect(result.output).toContain('Send it now.')
    expect(bridge.calls).toHaveLength(0)
  })

  it('quotes thread context for replies with citations and no send path', async () => {
    const bridge = new FakeAiBridge()
    const skill = createComposeAssistSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    const result = await skill.executeTool(
      call('reply_from_thread', { uids: ['a-m1'], folderId: 'inbox', focus: 'payment date' }),
    )
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('[a-m1]')
    expect(result.output).toContain('never send')
    expect(bridge.calls.map((c) => c.op)).toEqual(['fetch-bodies'])
  })
})

describe('search skill (ask-mail)', () => {
  it('searches headers-first and resolves header-only quote cards', async () => {
    const bridge = new FakeAiBridge()
    const skill = createMailSearchSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    const hits = await skill.executeTool(call('search_mail', { query: 'invoice' }))
    expect(hits.isError).toBeFalsy()
    expect(hits.output).toContain('[a-m1]')
    expect(hits.output).not.toContain('BODY-ONLY-A-M1')
    const quote = await skill.executeTool(
      call('quote_message', { uids: ['a-m1'], folderId: 'inbox' }),
    )
    expect(quote.isError).toBeFalsy()
    expect(quote.output).toContain('[a-m1]')
    expect(quote.output).toContain('<a-m1@example.com>')
    expect(quote.output).not.toContain('BODY-ONLY-A-M1')
    expect(bridge.calls.map((c) => c.op)).toEqual(['triage-headers', 'cite-resolve'])
  })

  it('requires citations on ask-mail answers', async () => {
    const bridge = new FakeAiBridge()
    const skill = createMailSearchSkill({
      getScope: () => scopeA,
      aiRetrieve: (r) => bridge.aiRetrieve(r),
    })
    await skill.executeTool(call('search_mail', { query: 'invoice' }))
    expect(
      skill.verifyResponse?.('The invoice is ready [a-m1]', [{ name: 'search_mail', ok: true }]),
    ).toBeNull()
    expect(
      skill.verifyResponse?.('- The invoice INV-1042 is ready for review with Friday terms', [
        { name: 'search_mail', ok: true },
      ]),
    ).toContain('citation')
  })
})

describe('skill composition (agent-core composeSkills)', () => {
  it('composes four skills with unique tools and chained guards', async () => {
    const bridge = new FakeAiBridge()
    const aiRetrieve = (r: AiRetrieveRequest) => bridge.aiRetrieve(r)
    const composed = composeSkills('mail-ai', 'Mail assistant.', [
      createTriageSkill({ getScope: () => scopeA, aiRetrieve }),
      createSummarizeSkill({ getScope: () => scopeA, aiRetrieve }),
      createComposeAssistSkill({ getScope: () => scopeA, aiRetrieve }),
      createMailSearchSkill({ getScope: () => scopeA, aiRetrieve }),
    ])
    expect(composed.id).toBe('mail-ai')
    expect(composed.tools.map((t) => t.name)).toEqual([
      'rank_folder',
      'apply_label',
      'summarize_thread',
      'catch_up_folder',
      'draft_from_bullets',
      'rewrite_tone',
      'reply_from_thread',
      'search_mail',
      'quote_message',
    ])
    expect(composed.systemPrompt).toContain('## Mail triage')
    expect(composed.systemPrompt).toContain('## Mail summarize')
    expect(composed.systemPrompt).toContain('## Mail compose assist')
    expect(composed.systemPrompt).toContain('## Mail search')
    const routed = await composed.executeTool(call('search_mail', { query: 'invoice' }))
    expect(routed.isError).toBeFalsy()
    expect(routed.output).toContain('[a-m1]')
    const correction = composed.verifyResponse?.('- Invoice ready, Friday terms, trust me', [
      { name: 'search_mail', ok: true },
    ])
    expect(correction).toContain('citation')
  })

  it('exposes the shared tool catalogue without duplicates', () => {
    const names = MAIL_AGENT_TOOLS.map((t) => t.name)
    expect(new Set(names).size).toBe(names.length)
    expect(names).toHaveLength(9)
  })
})
