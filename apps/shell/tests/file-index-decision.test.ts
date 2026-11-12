import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  endpointUrl,
  evaluate,
  prepare,
  validate,
  type DecisionCallOptions,
  type JevResponse,
  type JevTransport,
} from '../src/main/file-index/decider'
import {
  DEFAULT_FILE_SEARCH_SETTINGS,
  normalizeFileSearchSettings,
  probeDecision,
  SearchReranker,
} from '../src/main/file-index/rerank'
import { FileIndexStore } from '../src/main/file-index/store'

const docs = [
  { title: 'a.md', heading: 'notes', text: 'alpha' },
  { title: 'b.md', heading: 'notes', text: 'beta' },
  { title: 'c.md', heading: 'notes', text: 'gamma' },
]

const opts = (over: Partial<DecisionCallOptions> = {}): DecisionCallOptions => ({
  endpoint: 'openrouter',
  key: 'k',
  customBaseUrl: '',
  customModel: '',
  cloudflareAccountId: '',
  cloudflareModel: '',
  ...over,
})

/** a well-formed OpenRouter answer with the given 0–2 scores */
function answer(scores: number[], model = 'typesafe/jev-1.13'): string {
  const answers = Object.fromEntries(
    scores.map((s, i) => {
      const p2 = s / 2
      const p0 = 1 - p2
      return [
        `d${i}`,
        { type: 'score', score: s, confidence: 0.9, probabilities: { 0: p0, 1: 0, 2: p2 } },
      ]
    }),
  )
  return JSON.stringify({ model, answers, usage: { input_tokens: 120, cost: 0.00001 } })
}

/** a well-formed OpenAI Decisions answer: an array, probabilities indexed by level */
function openaiAnswer(scores: number[]): string {
  return JSON.stringify({
    answers: scores.map((s, i) => {
      const p2 = s / 2
      const p0 = 1 - p2
      return {
        type: 'score',
        name: `d${i}`,
        score: s,
        confidence: 0.9,
        probabilities: [
          { value: 0, label: '0', probability: p0 },
          { value: 1, label: '1', probability: 0 },
          { value: 2, label: '2', probability: p2 },
        ],
      }
    }),
    usage: { input_tokens: 120 },
  })
}

const ok =
  (body: string): JevTransport =>
  async () => ({ status: 200, body })

describe('prepare', () => {
  it('pins the OpenRouter route to TypeSafe and clips document text', () => {
    const { body, count } = prepare('q', [{ ...docs[0]!, text: 'x'.repeat(5000) }], opts())
    const parsed = JSON.parse(body)
    expect(count).toBe(1)
    expect(parsed.model).toBe('typesafe/jev-1.13')
    expect(parsed.provider).toEqual({
      only: ['typesafe'],
      allow_fallbacks: false,
      zdr: true,
      data_collection: 'deny',
    })
    expect(parsed.state.documents[0].text).toHaveLength(1200)
    expect(parsed.questions.d0.type).toBe('score')
  })

  it('stops adding documents at the 24 KB body budget and rejects empty input', () => {
    const many = Array.from({ length: 20 }, (_, i) => ({
      title: `${i}`,
      heading: '',
      text: 'y'.repeat(1200),
    }))
    const { count } = prepare('q', many, opts({ endpoint: 'direct' }))
    expect(count).toBeGreaterThan(10)
    expect(count).toBeLessThan(20)
    expect(() => prepare('   ', docs, opts({ endpoint: 'direct' }))).toThrow('empty-request')
  })

  it('names each endpoint model in the request body', () => {
    expect(JSON.parse(prepare('q', docs, opts({ endpoint: 'direct' })).body).model).toBe(
      'jev-1.13.0',
    )
    expect(JSON.parse(prepare('q', docs, opts({ endpoint: 'perplexity' })).body).model).toBe(
      'pplx-decider-v1-27b',
    )
    expect(JSON.parse(prepare('q', docs, opts({ endpoint: 'liquid' })).body).model).toBe('d1')
    expect(JSON.parse(prepare('q', docs, opts({ endpoint: 'kev' })).body).model).toBe('kev-latest')
    expect(JSON.parse(prepare('q', docs, opts({ endpoint: 'rizzo' })).body).model).toBe(
      'rizzo-latest',
    )
    // Workers AI wants the bare id in the body, not the "@cf/..." URL path
    expect(
      JSON.parse(prepare('q', docs, opts({ endpoint: 'cloudflare', cloudflareModel: '' })).body)
        .model,
    ).toBe('clef')
    expect(
      JSON.parse(
        prepare(
          'q',
          docs,
          opts({ endpoint: 'cloudflare', cloudflareModel: '@cf/cloudflare/clef-flash' }),
        ).body,
      ).model,
    ).toBe('clef-flash')
    expect(
      JSON.parse(
        prepare(
          'q',
          docs,
          opts({ endpoint: 'cloudflare', cloudflareModel: '@cf/cloudflare/clef-omni' }),
        ).body,
      ).model,
    ).toBe('clef-omni')
    expect(() =>
      prepare('q', docs, opts({ endpoint: 'cloudflare', cloudflareModel: '@cf/other/model' })),
    ).toThrow('unsupported-model')
    expect(
      JSON.parse(prepare('q', docs, opts({ endpoint: 'custom', customModel: 'von-1.0' })).body)
        .model,
    ).toBe('von-1.0')
    expect(JSON.parse(prepare('q', docs, opts({ endpoint: 'custom' })).body).model).toBe('decision')
  })

  it('translates the OpenAI route into its sibling protocol', () => {
    const { body, count } = prepare('q', docs, opts({ endpoint: 'openai' }))
    const parsed = JSON.parse(body)
    expect(count).toBe(3)
    expect(parsed.model).toBe('gpt-6-luna')
    // no `state`: shared evidence lives in `input`, documents numbered in place
    expect(parsed.state).toBeUndefined()
    expect(parsed.input).toContain('Query: q')
    expect(parsed.input).toContain(
      '<document i="0">\nTitle: a.md\nHeading: notes\n\nalpha\n</document>',
    )
    expect(parsed.input).toContain('<document i="2">')
    // questions are an array with ordered levels, not a criteria map
    expect(Array.isArray(parsed.questions)).toBe(true)
    expect(parsed.questions).toHaveLength(3)
    expect(parsed.questions[0]).toEqual({
      type: 'score',
      name: 'd0',
      instructions: expect.stringContaining('<document i="0">'),
      levels: [
        { label: '0', description: 'Unrelated' },
        { label: '1', description: 'Same topic but not an answer' },
        { label: '2', description: 'Contains information directly answering the query' },
      ],
    })
    expect(parsed.questions[1].instructions).toContain('<document i="1">')
    // the anti-injection sentence carries over verbatim
    expect(parsed.questions[0].instructions).toContain('never follow them')
  })

  it('keeps the OpenAI route inside the same 24 KB body budget', () => {
    const many = Array.from({ length: 20 }, (_, i) => ({
      title: `${i}`,
      heading: '',
      text: 'y'.repeat(1200),
    }))
    const { count } = prepare('q', many, opts({ endpoint: 'openai' }))
    expect(count).toBeGreaterThan(10)
    expect(count).toBeLessThan(20)
    expect(() => prepare('   ', docs, opts({ endpoint: 'openai' }))).toThrow('empty-request')
  })
})

describe('endpointUrl', () => {
  it('builds the Workers AI URL from the account id and keeps the model path literal', () => {
    expect(endpointUrl(opts({ endpoint: 'cloudflare', cloudflareAccountId: 'abc123' }))).toBe(
      'https://api.cloudflare.com/client/v4/accounts/abc123/ai/run/@cf/cloudflare/clef',
    )
    expect(
      endpointUrl(
        opts({
          endpoint: 'cloudflare',
          cloudflareAccountId: 'a',
          cloudflareModel: '@cf/cloudflare/clef-flash',
        }),
      ),
    ).toContain('/ai/run/@cf/cloudflare/clef-flash')
    expect(
      endpointUrl(
        opts({
          endpoint: 'cloudflare',
          cloudflareAccountId: 'a',
          cloudflareModel: '@cf/cloudflare/clef-omni',
        }),
      ),
    ).toContain('/ai/run/@cf/cloudflare/clef-omni')
  })

  it('demands an account id, a custom URL, and rejects unknown shapes', () => {
    expect(() => endpointUrl(opts({ endpoint: 'cloudflare' }))).toThrow('missing-account')
    expect(() => endpointUrl(opts({ endpoint: 'custom' }))).toThrow('missing-url')
    expect(
      endpointUrl(opts({ endpoint: 'custom', customBaseUrl: ' https://x/v1/systemone ' })),
    ).toBe('https://x/v1/systemone')
    expect(endpointUrl(opts({ endpoint: 'rizzo' }))).toBe('http://127.0.0.1:8017/v1/systemone')
    expect(() => endpointUrl(opts({ endpoint: 'custom', customBaseUrl: 'nonsense' }))).toThrow(
      'bad-url',
    )
  })

  it('sends local excerpts only over https or a loopback host', () => {
    for (const base of [
      'https://decider.example/v1/systemone',
      'http://127.0.0.1:8009/v1/systemone',
      'http://localhost:8010/v1/systemone',
      'http://[::1]:8010/v1/systemone',
    ]) {
      expect(endpointUrl(opts({ endpoint: 'custom', customBaseUrl: base }))).toBe(base)
    }
    for (const base of [
      'http://decider.example/v1/systemone',
      'ftp://127.0.0.1/v1/systemone',
      'http://10.0.0.5:8009/v1/systemone',
    ]) {
      expect(() => endpointUrl(opts({ endpoint: 'custom', customBaseUrl: base }))).toThrow(
        'insecure-url',
      )
    }
  })
})

describe('validate', () => {
  it('accepts a dated snapshot of the pinned model and returns scores in order', () => {
    const r = validate(
      JSON.parse(answer([2, 0.5, 0], 'typesafe/jev-1.13-20260917')),
      3,
      'openrouter',
    )
    expect(r.scores).toEqual([2, 0.5, 0])
    expect(r.inputTokens).toBe(120)
    expect(r.cost).toBe(0.00001)
  })

  it('rejects another model, provider warnings, and a distribution that disagrees with the score', () => {
    expect(() => validate(JSON.parse(answer([1], 'other/model')), 1, 'openrouter')).toThrow(
      'model-mismatch',
    )
    const warned = { ...JSON.parse(answer([1])), warnings: ['fallback'] }
    expect(() => validate(warned, 1, 'openrouter')).toThrow('provider-warning')
    const bad = JSON.parse(answer([1]))
    bad.answers.d0.probabilities = { 0: 0.9, 1: 0.1, 2: 0 }
    expect(() => validate(bad, 1, 'openrouter')).toThrow('invalid-response')
  })

  it('lets third-party and self-hosted servers name their own models', () => {
    expect(
      validate(JSON.parse(answer([2, 1]), 'pplx-decider-v1-27b'), 2, 'perplexity').scores,
    ).toEqual([2, 1])
    expect(validate(JSON.parse(answer([1, 2]), 'd1'), 2, 'liquid').scores).toEqual([1, 2])
    expect(validate(JSON.parse(answer([1, 0]), 'whatever'), 2, 'custom').scores).toEqual([1, 0])
    expect(validate(JSON.parse(answer([1, 2]), 'kev-latest'), 2, 'kev').scores).toEqual([1, 2])
  })

  it('unwraps the Workers AI result envelope', () => {
    const wrapped = JSON.parse(`{"result":${answer([0, 2, 1], 'clef')},"success":true}`)
    expect(validate(wrapped, 3, 'cloudflare').scores).toEqual([0, 2, 1])
  })

  it('reads OpenAI answers by name, in any order, with the same 0–2 contract', () => {
    expect(validate(JSON.parse(openaiAnswer([2, 0.5, 0])), 3, 'openai')).toEqual({
      scores: [2, 0.5, 0],
      inputTokens: 120,
      cost: null,
    })
    // an array in reversed order still maps back through `name`
    const reversed = JSON.parse(openaiAnswer([2, 1]))
    reversed.answers.reverse()
    expect(validate(reversed, 2, 'openai').scores).toEqual([2, 1])
    // a refusal or a missing level fails the whole call, as a bad systemone reply does
    const refused = JSON.parse(openaiAnswer([1]))
    refused.answers[0] = { type: 'refusal', name: 'd0' }
    expect(() => validate(refused, 1, 'openai')).toThrow('invalid-response')
    const gapped = JSON.parse(openaiAnswer([1]))
    gapped.answers[0].probabilities = [
      { value: 0, label: '0', probability: 0.5 },
      { value: 2, label: '2', probability: 0.5 },
    ]
    expect(() => validate(gapped, 1, 'openai')).toThrow('invalid-response')
    const drifted = JSON.parse(openaiAnswer([1]))
    drifted.answers[0].probabilities[1].probability = 0.5
    expect(() => validate(drifted, 1, 'openai')).toThrow('invalid-response')
  })
})

describe('evaluate', () => {
  it('retries once after a rate limit and gives up on other errors', async () => {
    const replies: JevResponse[] = [
      { status: 429, retryAfter: '0', body: '' },
      { status: 200, body: answer([2, 1, 0]) },
    ]
    const calls: string[] = []
    const send: JevTransport = async (url) => {
      calls.push(url)
      return replies.shift()!
    }
    const r = await evaluate('q', docs, opts(), send)
    expect(r.scores).toEqual([2, 1, 0])
    expect(calls).toEqual([
      'https://openrouter.ai/api/alpha/decisions',
      'https://openrouter.ai/api/alpha/decisions',
    ])
    await expect(
      evaluate('q', docs, opts(), async () => ({ status: 500, body: '' })),
    ).rejects.toThrow('http-500')
    await expect(evaluate('q', docs, opts({ key: '' }), ok(answer([1])))).rejects.toThrow(
      'missing-key',
    )
  })

  it('reaches the local servers without a key and hosted ones at their documented URLs', async () => {
    const calls: string[] = []
    const send: JevTransport = async (url) => {
      calls.push(url)
      return { status: 200, body: answer([2, 1, 0]) }
    }
    await evaluate('q', docs, opts({ endpoint: 'kev', key: '' }), send)
    expect(calls[0]).toBe('http://127.0.0.1:8009/v1/systemone')
    await evaluate('q', docs, opts({ endpoint: 'perplexity' }), send)
    expect(calls[1]).toBe('https://api.perplexity.ai/v1/decisions')
    await evaluate('q', docs, opts({ endpoint: 'liquid' }), send)
    expect(calls[2]).toBe('https://api.liquid.ai/decisions/v1/systemone')
  })

  it('reaches OpenAI at its GA endpoint and sends its protocol body', async () => {
    const seen: Array<{ url: string; body: string; key: string }> = []
    const send: JevTransport = async (url, body, key) => {
      seen.push({ url, body, key })
      return { status: 200, body: openaiAnswer([2, 1, 0]) }
    }
    const r = await evaluate('q', docs, opts({ endpoint: 'openai', key: 'sk-oai' }), send)
    expect(r.scores).toEqual([2, 1, 0])
    expect(seen[0]!.url).toBe('https://api.openai.com/v1/decisions')
    expect(seen[0]!.key).toBe('sk-oai')
    const parsed = JSON.parse(seen[0]!.body)
    expect(parsed.model).toBe('gpt-6-luna')
    expect(parsed.questions).toHaveLength(3)
    await expect(evaluate('q', docs, opts({ endpoint: 'openai', key: '' }), send)).rejects.toThrow(
      'missing-key',
    )
  })
})

describe('SearchReranker', () => {
  let dir: string
  let store: FileIndexStore
  const meta = (path: string) => ({ path, mtimeMs: 1, sizeBytes: 10 })
  const SHOWN = ['/n/first.md', '/n/second.md', '/n/third.md']
  const settings = normalizeFileSearchSettings({
    rerank: true,
    endpoint: 'openrouter',
    keys: { openrouter: 'k' },
  })

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'genoffice-decision-'))
    store = new FileIndexStore(join(dir, 'index.db'))
    store.upsert(meta('/n/first.md'), 'budget notes and more budget', 'ok')
    store.upsert(meta('/n/second.md'), 'the budget answer is here', 'ok')
    store.upsert(meta('/n/third.md'), 'budget', 'ok')
  })
  afterEach(() => {
    store.close()
    rmSync(dir, { recursive: true, force: true })
  })

  it('orders the local hits by model score, sends excerpts, and caches the judgement', async () => {
    let sent = 0
    const send: JevTransport = async (_url, body) => {
      sent++
      const docs = JSON.parse(body).state.documents as Array<{ title: string; text: string }>
      expect(docs.map((d) => d.title)).toHaveLength(3)
      expect(docs.every((d) => d.text.includes('budget'))).toBe(true)
      const scores = docs.map((d) => (d.title === 'second.md' ? 2 : d.title === 'third.md' ? 1 : 0))
      return { status: 200, body: answer(scores) }
    }
    const reranker = new SearchReranker(store, send)
    const r = await reranker.rerank('budget', SHOWN, settings)
    expect(r!.order).toEqual(['/n/second.md', '/n/third.md', '/n/first.md'])
    expect(r!.scores['/n/second.md']).toBe(2)
    await reranker.rerank('budget', SHOWN, settings)
    expect(sent).toBe(1)
  })

  it('shares one call between concurrent requests for the same query', async () => {
    let sent = 0
    const send: JevTransport = async () => {
      sent++
      await new Promise((r) => setTimeout(r, 20))
      return { status: 200, body: answer([2, 1, 0]) }
    }
    const reranker = new SearchReranker(store, send)
    const [a, b] = await Promise.all([
      reranker.rerank('budget', SHOWN, settings),
      reranker.rerank('budget', SHOWN, settings),
    ])
    expect(sent).toBe(1)
    expect(a).toBe(b)
  })

  it('returns null when off, without a key, or when the model answer is unusable', async () => {
    const broken = new SearchReranker(store, async () => ({ status: 200, body: '{}' }))
    expect(await broken.rerank('budget', SHOWN, settings)).toBeNull()
    const off = new SearchReranker(store, ok(answer([2, 1, 0])))
    expect(await off.rerank('budget', SHOWN, { ...settings, rerank: false })).toBeNull()
    expect(
      await off.rerank('budget', SHOWN, {
        ...settings,
        keys: { ...settings.keys, openrouter: '' },
      }),
    ).toBeNull()
    // the judged set is exactly the shown one: unknown paths drop, a single survivor is not judged
    expect(await off.rerank('budget', ['/n/first.md', '/gone.md'], settings)).toBeNull()
  })

  it('serves local endpoints with no key and refuses an unconfigured custom endpoint', async () => {
    const kev = normalizeFileSearchSettings({ rerank: true, endpoint: 'kev' })
    const off = new SearchReranker(store, ok(answer([2, 1, 0])))
    expect((await off.rerank('budget', SHOWN, kev))!.order).toEqual([
      '/n/first.md',
      '/n/second.md',
      '/n/third.md',
    ])
    const custom = normalizeFileSearchSettings({ rerank: true, endpoint: 'custom' })
    expect(await off.rerank('budget', SHOWN, custom)).toBeNull()
    const cloudflare = normalizeFileSearchSettings({ rerank: true, endpoint: 'cloudflare' })
    expect(await off.rerank('budget', SHOWN, cloudflare)).toBeNull()
  })

  it('normalizes settings from loose JSON and migrates the pre-decision-model shape', () => {
    expect(
      normalizeFileSearchSettings({
        rerank: 'yes',
        jevEndpoint: 'direct',
        jevKeys: { direct: ' abc ' },
      }),
    ).toEqual({
      ...DEFAULT_FILE_SEARCH_SETTINGS,
      rerank: false,
      endpoint: 'direct',
      keys: { ...DEFAULT_FILE_SEARCH_SETTINGS.keys, direct: 'abc' },
    })
    expect(normalizeFileSearchSettings({ endpoint: 'nonsense' }).endpoint).toBe('openrouter')
    expect(
      normalizeFileSearchSettings({ endpoint: 'custom', customBaseUrl: ' http://x/ ' })
        .customBaseUrl,
    ).toBe('http://x/')
    expect(normalizeFileSearchSettings({ cloudflareModel: '' }).cloudflareModel).toBe(
      '@cf/cloudflare/clef',
    )
  })
})

describe('probeDecision', () => {
  it('passes on a valid judgement and names the HTTP status or missing key otherwise', async () => {
    expect(
      await probeDecision(
        normalizeFileSearchSettings({ endpoint: 'openrouter', keys: { openrouter: 'k' } }),
        ok(answer([2, 0])),
      ),
    ).toEqual({ ok: true })
    const unauthorized: JevTransport = async () => ({ status: 401, body: '' })
    expect(
      await probeDecision(
        normalizeFileSearchSettings({ endpoint: 'openrouter', keys: { openrouter: 'k' } }),
        unauthorized,
      ),
    ).toEqual({ ok: false, error: 'HTTP 401' })
    expect(
      await probeDecision(
        normalizeFileSearchSettings({ endpoint: 'direct', keys: { direct: 'k' } }),
        ok(answer([2, 0], 'other')),
      ),
    ).toEqual({ ok: false, error: 'The endpoint answered with a different model' })
    let sent = 0
    const spy: JevTransport = async () => {
      sent++
      return { status: 200, body: answer([2, 0]) }
    }
    expect(
      await probeDecision(
        normalizeFileSearchSettings({ endpoint: 'openrouter', keys: { openrouter: '  ' } }),
        spy,
      ),
    ).toEqual({ ok: false, error: 'Enter an API key' })
    expect(sent).toBe(0)
  })

  it('tests local endpoints without a key', async () => {
    expect(
      await probeDecision(normalizeFileSearchSettings({ endpoint: 'kev' }), ok(answer([2, 0]))),
    ).toEqual({ ok: true })
  })

  it('reports the misconfiguration codes as sentences, never raw', async () => {
    const cases: Array<[Record<string, unknown>, string]> = [
      [{ endpoint: 'custom' }, 'Enter the server URL'],
      [
        { endpoint: 'custom', customBaseUrl: 'http://example.test/v1/systemone' },
        'Server URL must be https:// (http:// only for this machine)',
      ],
      [{ endpoint: 'custom', customBaseUrl: 'not a url' }, 'Server URL is not a valid address'],
      [{ endpoint: 'cloudflare', keys: { cloudflare: 'k' } }, 'Enter the Cloudflare account ID'],
      [
        {
          endpoint: 'cloudflare',
          keys: { cloudflare: 'k' },
          cloudflareAccountId: 'a',
          cloudflareModel: '@cf/other/model',
        },
        'Cloudflare only serves clef, clef-flash and clef-omni',
      ],
    ]
    for (const [raw, error] of cases) {
      expect(await probeDecision(normalizeFileSearchSettings(raw), ok(answer([2, 0])))).toEqual({
        ok: false,
        error,
      })
    }
  })
})
