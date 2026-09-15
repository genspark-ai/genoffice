import { beforeEach, describe, expect, it, vi } from 'vitest'

// mock the provider entrypoint before importing the module under test
vi.mock('@genoffice/ai-provider', async () => {
  const actual =
    await vi.importActual<typeof import('@genoffice/ai-provider')>('@genoffice/ai-provider')
  return {
    ...actual,
    chatForProvider: vi.fn(),
  }
})

import { chatForProvider } from '@genoffice/ai-provider'

import { TranslationMemory } from '../src/memory'
import { sharedMemory, translateBatch, translateOne } from '../src/provider'

const mockedChat = vi.mocked(chatForProvider)

describe('translateOne', () => {
  beforeEach(() => {
    mockedChat.mockReset()
    sharedMemory.clear()
  })

  it('rejects empty instruction', async () => {
    const r = await translateOne(
      { instruction: '', targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(r.ok).toBe(false)
    expect(r.error).toMatch(/instruction/)
    expect(mockedChat).not.toHaveBeenCalled()
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
    mockedChat.mockResolvedValue({ ok: true, content: '你好' })
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
    expect(mockedChat).toHaveBeenCalledOnce()
  })

  it('strips <think> blocks before returning', async () => {
    mockedChat.mockResolvedValue({ ok: true, content: '<think>internal</think>你好' })
    const r = await translateOne(
      { instruction: 'StripMe', targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(r.translated).toBe('你好')
  })

  it('flags an empty provider response', async () => {
    mockedChat.mockResolvedValue({ ok: true, content: '<think>only thoughts</think>' })
    const r = await translateOne(
      { instruction: 'EmptyMe', targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' } },
    )
    expect(r.ok).toBe(false)
    expect(r.error).toMatch(/final text/)
  })

  it('reports provider errors', async () => {
    mockedChat.mockResolvedValue({ ok: false, error: 'HTTP 500' })
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
    expect(mockedChat).not.toHaveBeenCalled()
  })

  it('saves successful translations back into the memory', async () => {
    mockedChat.mockResolvedValue({ ok: true, content: '你好' })
    const mem = new TranslationMemory()
    await translateOne(
      { instruction: 'Hello', targetLang: 'zh-CN' },
      { provider: 'anthropic', config: { apiKey: 'k', model: 'm' }, memory: mem },
    )
    expect(mem.lookup('auto', 'zh-CN', 'Hello')?.translatedText).toBe('你好')
  })
})

describe('translateBatch', () => {
  beforeEach(() => {
    mockedChat.mockReset()
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
    mockedChat.mockImplementation(async (_provider, _cfg, _sys, user) => {
      if (user.includes('Hello')) return { ok: true, content: '你好' }
      if (user.includes('World')) return { ok: false, error: 'boom' }
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
    mockedChat.mockResolvedValue({ ok: true, content: '好' })
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
    mockedChat.mockResolvedValue({ ok: true, content: '世界' })
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
    expect(mockedChat).toHaveBeenCalledTimes(1)
  })
})
