import { beforeEach, describe, expect, it, vi } from 'vitest'

// mock the LLM client seam before importing the module under test
// (W9: translation-core now talks to `callLlm` from `../src/llm-client`
// rather than `chatForProvider` directly. This lets us swap the underlying
// SDK without touching tests.)
vi.mock('../src/llm-client', async () => {
  const actual = await vi.importActual<typeof import('../src/llm-client')>('../src/llm-client')
  return {
    ...actual,
    callLlm: vi.fn(),
  }
})

import { callLlm } from '../src/llm-client'

import { KnowledgeBase } from '../src/knowledge-base'
import { TranslationMemory } from '../src/memory'
import { sharedMemory, translateBatch, translateBatchStream, translateOne } from '../src/provider'

const mockedCall = vi.mocked(callLlm)

describe('translateOne', () => {
  beforeEach(() => {
    mockedCall.mockReset()
    sharedMemory.clear()
  })

  it('rejects empty instruction', async () => {
    const r = await translateOne(
      { instruction: '', targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(r.ok).toBe(false)
    expect(r.error).toMatch(/instruction/)
    expect(mockedCall).not.toHaveBeenCalled()
  })

  it('rejects empty target lang', async () => {
    const r = await translateOne(
      { instruction: 'hi', targetLang: '' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(r.ok).toBe(false)
    expect(r.error).toMatch(/targetLang/)
  })

  it('rejects missing API key (non-codex / non-genspark)', async () => {
    const r = await translateOne(
      { instruction: 'hi', targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: '', model: 'm' } },
    )
    expect(r.ok).toBe(false)
    expect(r.error).toMatch(/API key/)
  })

  it('rejects missing model (non-codex)', async () => {
    const r = await translateOne(
      { instruction: 'hi', targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: 'k', model: '' } },
    )
    expect(r.ok).toBe(false)
    expect(r.error).toMatch(/model/i)
  })

  it('calls chatForProvider and returns the translation', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: '你好' })
    const r = await translateOne(
      { instruction: 'Hello', targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'claude-sonnet-5' } },
    )
    expect(r).toEqual({
      ok: true,
      translated: '你好',
      planId: expect.stringMatching(/^translate-/),
      sourceLang: 'auto',
      targetLang: 'zh-CN',
      preserveFormat: true,
      status: 'translated',
    })
    expect(mockedCall).toHaveBeenCalledOnce()
  })

  it('strips <think> blocks before returning', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: '<think>internal</think>你好' })
    const r = await translateOne(
      { instruction: 'StripMe', targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(r.translated).toBe('你好')
  })

  it('flags an empty provider response', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: '<think>only thoughts</think>' })
    const r = await translateOne(
      { instruction: 'EmptyMe', targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(r.ok).toBe(false)
    expect(r.error).toMatch(/final text/)
  })

  it('reports provider errors', async () => {
    mockedCall.mockResolvedValue({ ok: false, error: 'HTTP 500' })
    const r = await translateOne(
      { instruction: 'ErrorMe', targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(r.ok).toBe(false)
    expect(r.error).toBe('HTTP 500')
  })

  it('serves a memory hit without calling the provider', async () => {
    const mem = new TranslationMemory()
    mem.save({
      sourceLang: 'auto',
      targetLang: 'zh-CN',
      sourceText: 'Hello',
      translatedText: '你好',
    })
    const r = await translateOne(
      { instruction: 'Hello', targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' }, memory: mem },
    )
    expect(r.status).toBe('memory-hit')
    expect(r.translated).toBe('你好')
    expect(mockedCall).not.toHaveBeenCalled()
  })

  it('saves successful translations back into the memory', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: '你好' })
    const mem = new TranslationMemory()
    await translateOne(
      { instruction: 'Hello', targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' }, memory: mem },
    )
    expect(mem.lookup('auto', 'zh-CN', 'Hello')?.translatedText).toBe('你好')
  })

  it('memoryEnabled=false bypasses the shared TM and calls the provider', async () => {
    const mem = new TranslationMemory()
    mem.save({
      sourceLang: 'auto',
      targetLang: 'zh-CN',
      sourceText: 'Hello',
      translatedText: '缓存命中（应被忽略）',
    })
    mockedCall.mockResolvedValue({ ok: true, content: '你好' })
    const r = await translateOne(
      { instruction: 'Hello', targetLang: 'zh-CN', memoryEnabled: false },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' }, memory: mem },
    )
    expect(r.status).toBe('translated')
    expect(r.translated).toBe('你好')
    expect(mockedCall).toHaveBeenCalledTimes(1)
  })

  it('memoryEnabled=false also prevents saving the new translation into TM', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: '你好' })
    const mem = new TranslationMemory()
    await translateOne(
      { instruction: 'Hello', targetLang: 'zh-CN', memoryEnabled: false },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' }, memory: mem },
    )
    expect(mem.lookup('auto', 'zh-CN', 'Hello')).toBeNull()
  })

  it('glossaryCategory is threaded through to the provider call metadata', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: '你好' })
    await translateOne(
      { instruction: 'Hello', targetLang: 'zh-CN', glossaryCategory: 'legal' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(mockedCall).toHaveBeenCalledTimes(1)
    const call = mockedCall.mock.calls[0]
    // glossaryCategory flows through to the chat request — at least one arg references it
    const flat = JSON.stringify(call)
    expect(flat).toContain('legal')
  })
})

describe('translateBatch', () => {
  beforeEach(() => {
    mockedCall.mockReset()
    sharedMemory.clear()
  })

  it('rejects an empty units array', async () => {
    const r = await translateBatch(
      { units: [], targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(r.ok).toBe(false)
    expect(r.error).toMatch(/units/)
  })

  it('returns per-unit results with status flags', async () => {
    mockedCall.mockImplementation(async (opts) => {
      if (opts.userPrompt.includes('Hello')) return { ok: true, content: '你好' }
      if (opts.userPrompt.includes('World')) return { ok: false, error: 'boom' }
      return { ok: true, content: '？' }
    })
    const r = await translateBatch(
      {
        units: [
          { unitId: 'u1', kind: 'paragraph', sourceText: 'Hello', order: 0 },
          { unitId: 'u2', kind: 'paragraph', sourceText: 'World', order: 1 },
        ],
        targetLang: 'zh-CN',
      },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(r.ok).toBe(false)
    expect(r.units?.[0].status).toBe('translated')
    expect(r.units?.[0].translatedText).toBe('你好')
    expect(r.units?.[1].status).toBe('failed')
    expect(r.units?.[1].errorMessage).toBe('boom')
    expect(r.error).toBe('boom')
  })

  it('marks all units translated when every call succeeds', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: '好' })
    const r = await translateBatch(
      {
        units: [
          { unitId: 'u1', kind: 'paragraph', sourceText: 'Alpha', order: 0 },
          { unitId: 'u2', kind: 'paragraph', sourceText: 'Beta', order: 1 },
        ],
        targetLang: 'zh-CN',
      },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(r.ok).toBe(true)
    expect(r.quality?.overallScore).toBeGreaterThan(0)
  })

  it('serves memory hits for known units', async () => {
    const mem = new TranslationMemory()
    mem.save({
      sourceLang: 'auto',
      targetLang: 'zh-CN',
      sourceText: 'Hello',
      translatedText: '你好',
    })
    mockedCall.mockResolvedValue({ ok: true, content: '世界' })
    const r = await translateBatch(
      {
        units: [
          { unitId: 'u1', kind: 'paragraph', sourceText: 'Hello', order: 0 },
          { unitId: 'u2', kind: 'paragraph', sourceText: 'World', order: 1 },
        ],
        targetLang: 'zh-CN',
      },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' }, memory: mem },
    )
    expect(r.units?.[0].status).toBe('memory-hit')
    expect(r.units?.[1].status).toBe('translated')
    // only the second unit hit the provider
    expect(mockedCall).toHaveBeenCalledTimes(1)
  })
})


describe('translateBatchStream', () => {
  beforeEach(() => {
    mockedCall.mockReset()
    sharedMemory.clear()
  })

  it('rejects empty units array', async () => {
    const r = await translateBatchStream(
      { units: [], targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(r.ok).toBe(false)
    expect(r.error).toMatch(/units/)
    expect(mockedCall).not.toHaveBeenCalled()
  })

  it('rejects empty target lang', async () => {
    const r = await translateBatchStream(
      { units: [{ unitId: 'u1', kind: 'paragraph', sourceText: 'hi', order: 0 }], targetLang: '' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(r.ok).toBe(false)
    expect(r.error).toMatch(/targetLang/)
  })

  it('fires onUnit for every translated unit with stable order payload', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: '<source_text>hi</source_text>译' })
    const units = [
      { unitId: 'u1', kind: 'paragraph' as const, sourceText: 'hello', order: 0 },
      { unitId: 'u2', kind: 'paragraph' as const, sourceText: 'world', order: 1 },
      { unitId: 'u3', kind: 'paragraph' as const, sourceText: 'foo', order: 2 },
    ]
    const events: Array<{ unitId: string; index: number; total: number }> = []
    const r = await translateBatchStream(
      { units, targetLang: 'zh-CN', qualityCheck: false },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
      {
        concurrency: 2,
        onUnit: (e) => {
          events.push({ unitId: e.result.unitId, index: e.index, total: e.total })
        },
      },
    )
    expect(r.ok).toBe(true)
    expect(r.units).toHaveLength(3)
    expect(events).toHaveLength(3)
    expect(events.map((e) => e.total).every((t) => t === 3)).toBe(true)
    const seenUnitIds = new Set(events.map((e) => e.unitId))
    expect(seenUnitIds).toEqual(new Set(['u1', 'u2', 'u3']))
  })

  it('serves memory hits without calling the provider and emits unit event', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: '<source_text>fresh</source_text>新' })
    sharedMemory.save({ sourceLang: 'en-US', targetLang: 'zh-CN', sourceText: 'cached', translatedText: '已缓存' })
    const events: string[] = []
    const r = await translateBatchStream(
      {
        units: [
          { unitId: 'u1', kind: 'paragraph' as const, sourceText: 'cached', order: 0 },
          { unitId: 'u2', kind: 'paragraph' as const, sourceText: 'fresh', order: 1 },
        ],
        sourceLang: 'en-US',
        targetLang: 'zh-CN',
        qualityCheck: false,
      },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
      { onUnit: (e) => { events.push(`${e.result.unitId}:${e.result.status}`) } },
    )
    expect(r.ok).toBe(true)
    expect(events).toContain('u1:memory-hit')
    expect(events).toContain('u2:translated')
  })

  it('skips provider calls when memoryEnabled is false and skips memory write', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: '<source_text>hi</source_text>译' })
    const r = await translateBatchStream(
      {
        units: [{ unitId: 'u1', kind: 'paragraph' as const, sourceText: 'hi', order: 0 }],
        targetLang: 'zh-CN',
        memoryEnabled: false,
        qualityCheck: false,
      },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(r.ok).toBe(true)
    expect(sharedMemory.size()).toBe(0)
  })

  it('returns the same final shape as translateBatch for the same input', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: '<source_text>hi</source_text>译' })
    const units = [
      { unitId: 'u1', kind: 'paragraph' as const, sourceText: 'a', order: 0 },
      { unitId: 'u2', kind: 'paragraph' as const, sourceText: 'b', order: 1 },
    ]
    const batch = await translateBatch(
      { units, targetLang: 'zh-CN', qualityCheck: false },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    mockedCall.mockClear()
    const stream = await translateBatchStream(
      { units, targetLang: 'zh-CN', qualityCheck: false },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(stream.ok).toBe(batch.ok)
    expect(stream.units?.map((u) => u.unitId)).toEqual(batch.units?.map((u) => u.unitId))
    expect(stream.units?.map((u) => u.translatedText)).toEqual(batch.units?.map((u) => u.translatedText))
    expect(stream.quality?.overallScore).toBe(batch.quality?.overallScore)
  })
})

// ---------------------------------------------------------------------------
// Terminology wiring: `matchedTerms` + the host-supplied dictionary.
//
// The Settings pane renders a "KB · N" badge off `matchedTerms` and can layer
// the generated `--dictionary` onto snippet translations. Both were previously
// never populated by translateOne, so these pin the contract.
// ---------------------------------------------------------------------------
describe('translateOne — terminology provenance', () => {
  const config = { apiKey: 'k', model: 'm' }
  const kb = () =>
    new KnowledgeBase({
      seed: {
        'trade.translation.term': [
          {
            id: 't1',
            scope: 'company',
            priority: 50,
            sourceTerm: 'fabric weight',
            targetTerm: '克重',
          },
          {
            id: 't2',
            scope: 'company',
            priority: 50,
            sourceTerm: 'cotton',
            targetTerm: '棉',
          },
        ],
        'trade.translation.brand': [
          { id: 'b1', scope: 'global', priority: 50, word: 'GSM', policy: 'neverTranslate' },
        ],
      },
    })

  beforeEach(() => {
    mockedCall.mockReset()
    sharedMemory.clear()
  })

  it('reports the KB terms the source text actually touched', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: '面料克重为220。' })
    const r = await translateOne(
      { instruction: 'The fabric weight is 220.', targetLang: 'zh-CN' },
      { provider: 'anthropic', config, knowledgeBase: kb() },
    )
    expect(r.ok).toBe(true)
    expect(r.matchedTerms).toEqual(['fabric weight'])
  })

  it('omits matchedTerms when no KB term is present', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: '你好' })
    const r = await translateOne(
      { instruction: 'Hello there', targetLang: 'zh-CN' },
      { provider: 'anthropic', config, knowledgeBase: kb() },
    )
    expect(r.ok).toBe(true)
    expect(r.matchedTerms).toBeUndefined()
  })

  it('enforces a KB term the model left in the source language', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: 'The fabric weight is heavy.' })
    const r = await translateOne(
      { instruction: 'The fabric weight is heavy.', targetLang: 'zh-CN' },
      { provider: 'anthropic', config, knowledgeBase: kb() },
    )
    expect(r.translated).toBe('The 克重 is heavy.')
    expect(r.matchedTerms).toEqual(['fabric weight'])
  })

  it('still counts terms on a memory hit', async () => {
    const memory = new TranslationMemory()
    memory.save({
      sourceLang: 'en-US',
      targetLang: 'zh-CN',
      sourceText: 'cotton shirt',
      translatedText: '棉衬衫',
    })
    const r = await translateOne(
      { instruction: 'cotton shirt', sourceLang: 'en-US', targetLang: 'zh-CN' },
      { provider: 'anthropic', config, knowledgeBase: kb(), memory },
    )
    expect(r.status).toBe('memory-hit')
    expect(r.matchedTerms).toEqual(['cotton'])
    expect(mockedCall).not.toHaveBeenCalled()
  })

  it('honours a host dictionary: prompt, enforcement and matchedTerms', async () => {
    // The model left the term in English; enforcement must substitute it.
    mockedCall.mockResolvedValue({ ok: true, content: 'Order 100 units of Oxford cloth.' })
    const r = await translateOne(
      { instruction: 'Order 100 units of Oxford cloth.', targetLang: 'zh-CN' },
      {
        provider: 'anthropic',
        config,
        dictionary: [{ source: 'Oxford cloth', target: '牛津布' }],
      },
    )
    expect(r.translated).toBe('Order 100 units of 牛津布.')
    expect(r.matchedTerms).toEqual(['Oxford cloth'])
    const prompt = mockedCall.mock.calls[0]?.[0]?.systemPrompt ?? ''
    expect(prompt).toContain('Oxford cloth => 牛津布')
  })

  it('does not inject dictionary terms the source text lacks', async () => {
    mockedCall.mockResolvedValue({ ok: true, content: '你好' })
    await translateOne(
      { instruction: 'Hello', targetLang: 'zh-CN' },
      {
        provider: 'anthropic',
        config,
        dictionary: [
          { source: 'Oxford cloth', target: '牛津布' },
          { source: 'poplin', target: '府绸' },
        ],
      },
    )
    const prompt = mockedCall.mock.calls[0]?.[0]?.systemPrompt ?? ''
    expect(prompt).not.toContain('牛津布')
    expect(prompt).not.toContain('府绸')
  })
})
