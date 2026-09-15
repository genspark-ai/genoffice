/**
 * Tests for the LLM client seam (W9 deliverable).
 *
 * Verifies:
 *   - callLlm routes through the active caller (default = aiProviderCaller)
 *   - setLlmCaller replaces the active caller (testbed-friendly)
 *   - piAiCaller throws "not implemented" until the migration lands
 *   - aiProviderCaller wraps @genoffice/ai-provider's chatForProvider faithfully
 *     (overloaded flag passes through, errors propagate)
 */

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

// Mock the underlying ai-provider so we don't hit the network
vi.mock('@genoffice/ai-provider', async () => {
  const actual = await vi.importActual<typeof import('@genoffice/ai-provider')>('@genoffice/ai-provider')
  return {
    ...actual,
    chatForProvider: vi.fn(),
    isAiOverloadedError: vi.fn((e: unknown) =>
      typeof e === 'string' && e.includes('overloaded'),
    ),
  }
})

import { chatForProvider, isAiOverloadedError } from '@genoffice/ai-provider'
import {
  aiProviderCaller,
  callLlm,
  callLlmWith,
  getLlmCaller,
  piAiCaller,
  setLlmCaller,
} from '../src/llm-client'

const mockedChat = vi.mocked(chatForProvider)
const mockedOverloaded = vi.mocked(isAiOverloadedError)

describe('llm-client seam', () => {
  beforeEach(() => {
    mockedChat.mockReset()
    mockedOverloaded.mockReset()
    // Reset the active caller to the default after each test
    setLlmCaller(aiProviderCaller)
  })

  afterEach(() => {
    setLlmCaller(aiProviderCaller)
  })

  it('default caller is aiProviderCaller', () => {
    expect(getLlmCaller()).toBe(aiProviderCaller)
  })

  it('callLlm routes to the active caller', async () => {
    mockedChat.mockResolvedValue({ ok: true, content: '你好' })
    const r = await callLlm({
      provider: 'anthropic',
      config: { apiKey: 'k', model: 'm' },
      systemPrompt: 'SYS',
      userPrompt: 'USER',
    })
    expect(r).toEqual({ ok: true, content: '你好' })
    expect(mockedChat).toHaveBeenCalledWith('anthropic', { apiKey: 'k', model: 'm' }, 'SYS', 'USER')
  })

  it('aiProviderCaller maps ok=true content to LlmCallResult', async () => {
    mockedChat.mockResolvedValue({ ok: true, content: 'hola' })
    const r = await aiProviderCaller({
      provider: 'gemini',
      config: { apiKey: 'k', model: 'g-1' },
      systemPrompt: 's',
      userPrompt: 'u',
    })
    expect(r.ok).toBe(true)
    expect(r.content).toBe('hola')
  })

  it('aiProviderCaller surfaces the overloaded flag', async () => {
    mockedChat.mockResolvedValue({ ok: false, error: 'overloaded: try later' })
    mockedOverloaded.mockReturnValue(true)
    const r = await aiProviderCaller({
      provider: 'openai',
      config: { apiKey: 'k', model: 'm' },
      systemPrompt: 's',
      userPrompt: 'u',
    })
    expect(r.ok).toBe(false)
    expect(r.overloaded).toBe(true)
    expect(r.error).toBe('overloaded: try later')
  })

  it('aiProviderCaller wraps thrown errors and maps overloaded', async () => {
    mockedChat.mockRejectedValue(new Error('Network down'))
    mockedOverloaded.mockReturnValue(false)
    const r = await aiProviderCaller({
      provider: 'openai',
      config: { apiKey: 'k', model: 'm' },
      systemPrompt: 's',
      userPrompt: 'u',
    })
    expect(r.ok).toBe(false)
    expect(r.error).toBe('Network down')
    expect(r.overloaded).toBe(false)
  })

  it('setLlmCaller swaps the active caller', async () => {
    const stubCaller = vi.fn(async () => ({ ok: true, content: 'STUB' }))
    setLlmCaller(stubCaller)

    const r = await callLlm({
      provider: 'anthropic',
      config: { apiKey: 'k', model: 'm' },
      systemPrompt: 's',
      userPrompt: 'u',
    })
    expect(r).toEqual({ ok: true, content: 'STUB' })
    expect(stubCaller).toHaveBeenCalledOnce()
    expect(mockedChat).not.toHaveBeenCalled()
  })

  it('callLlmWith bypasses the global active caller', async () => {
    const stubCaller = vi.fn(async () => ({ ok: true, content: 'BYPASS' }))
    const r = await callLlmWith(stubCaller, {
      provider: 'anthropic',
      config: { apiKey: 'k', model: 'm' },
      systemPrompt: 's',
      userPrompt: 'u',
    })
    expect(r).toEqual({ ok: true, content: 'BYPASS' })
  })

  it('piAiCaller throws not-implemented until the migration lands', async () => {
    await expect(
      piAiCaller({
        provider: 'anthropic',
        config: { apiKey: 'k', model: 'm' },
        systemPrompt: 's',
        userPrompt: 'u',
      }),
    ).rejects.toThrow(/piAiCaller is not yet implemented/)
  })
})
