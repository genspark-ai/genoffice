import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  evaluate,
  prepare,
  validate,
  type JevResponse,
  type JevTransport,
} from '../src/main/file-index/jev'
import {
  normalizeFileSearchSettings,
  probeJev,
  SearchReranker,
} from '../src/main/file-index/rerank'
import { FileIndexStore } from '../src/main/file-index/store'

const docs = [
  { title: 'a.md', heading: 'notes', text: 'alpha' },
  { title: 'b.md', heading: 'notes', text: 'beta' },
  { title: 'c.md', heading: 'notes', text: 'gamma' },
]

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

const ok =
  (body: string): JevTransport =>
  async () => ({ status: 200, body })

describe('prepare', () => {
  it('pins the OpenRouter route to TypeSafe and clips document text', () => {
    const { body, count } = prepare('q', [{ ...docs[0]!, text: 'x'.repeat(5000) }], 'openrouter')
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
    const { count } = prepare('q', many, 'direct')
    expect(count).toBeGreaterThan(10)
    expect(count).toBeLessThan(20)
    expect(() => prepare('   ', docs, 'direct')).toThrow('empty-request')
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
    const r = await evaluate('q', docs, 'openrouter', 'k', send)
    expect(r.scores).toEqual([2, 1, 0])
    expect(calls).toEqual([
      'https://openrouter.ai/api/alpha/decisions',
      'https://openrouter.ai/api/alpha/decisions',
    ])
    await expect(
      evaluate('q', docs, 'openrouter', 'k', async () => ({ status: 500, body: '' })),
    ).rejects.toThrow('http-500')
    await expect(evaluate('q', docs, 'openrouter', '', ok(answer([1])))).rejects.toThrow(
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
    jevEndpoint: 'openrouter',
    jevKeys: { openrouter: 'k' },
  })

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'genoffice-jev-'))
    store = new FileIndexStore(join(dir, 'index.db'))
    store.upsert(meta('/n/first.md'), 'budget notes and more budget', 'ok')
    store.upsert(meta('/n/second.md'), 'the budget answer is here', 'ok')
    store.upsert(meta('/n/third.md'), 'budget', 'ok')
  })
  afterEach(() => {
    store.close()
    rmSync(dir, { recursive: true, force: true })
  })

  it('orders the local hits by Jev score, sends excerpts, and caches the judgement', async () => {
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
        jevKeys: { openrouter: '', direct: '' },
      }),
    ).toBeNull()
    // the judged set is exactly the shown one: unknown paths drop, a single survivor is not judged
    expect(await off.rerank('budget', ['/n/first.md', '/gone.md'], settings)).toBeNull()
  })

  it('normalizes settings from loose JSON', () => {
    expect(
      normalizeFileSearchSettings({
        rerank: 'yes',
        jevEndpoint: 'direct',
        jevKeys: { direct: ' abc ' },
      }),
    ).toEqual({
      rerank: false,
      jevEndpoint: 'direct',
      jevKeys: { openrouter: '', direct: 'abc' },
    })
  })
})

describe('probeJev', () => {
  it('passes on a valid judgement and names the HTTP status or missing key otherwise', async () => {
    expect(await probeJev('openrouter', 'k', ok(answer([2, 0])))).toEqual({ ok: true })
    const unauthorized: JevTransport = async () => ({ status: 401, body: '' })
    expect(await probeJev('openrouter', 'k', unauthorized)).toEqual({
      ok: false,
      error: 'HTTP 401',
    })
    expect(await probeJev('direct', 'k', ok(answer([2, 0])))).toEqual({
      ok: false,
      error: 'model-mismatch',
    })
    let sent = 0
    const spy: JevTransport = async () => {
      sent++
      return { status: 200, body: answer([2, 0]) }
    }
    expect(await probeJev('openrouter', '  ', spy)).toEqual({
      ok: false,
      error: 'Enter an API key',
    })
    expect(sent).toBe(0)
  })
})
