/**
 * Stream-level integration for the declarative custom endpoint: the declared
 * protocol decides the wire shape of the request (path, auth header, body
 * naming), the declared headers reach the wire, and the model declaration
 * caps the output budget. These are the behaviors that make a self-hosted
 * gateway usable without any code change.
 */
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentToolCall } from '@genoffice/agent-core'
import { streamForProvider } from '../src/stream'
import { chatForProvider } from '../src/chat'
import type { AiProviderConfig, AiSettings } from '../src/types'
import { jsonResponse, okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

function collector() {
  const deltas: string[] = []
  const toolCalls: AgentToolCall[] = []
  return {
    cb: {
      signal: new AbortController().signal,
      onDelta: (text: string) => deltas.push(text),
      onToolCall: (call: AgentToolCall) => toolCalls.push(call),
    },
    deltas,
    toolCalls,
  }
}

const BASE: AiProviderConfig = {
  apiKey: 'k',
  model: 'step-3.7-flash',
  baseUrl: 'https://gw.example.com/v1',
}

const sse = (lines: string[]) => okResponse(sseStream(lines))

describe('streamForProvider: custom endpoint protocol dispatch', () => {
  it('speaks OpenAI chat-completions by default', async () => {
    const fetchMock = vi.fn().mockResolvedValue(sse(['data: [DONE]']))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider('custom', BASE, 'sys', [], [], 100, cb).catch(() => {})
    expect(fetchMock).toHaveBeenCalledWith(
      'https://gw.example.com/v1/chat/completions',
      expect.anything(),
    )
  })

  it('speaks the Anthropic Messages route when declared', async () => {
    const fetchMock = vi.fn().mockResolvedValue(sse(['data: [DONE]']))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'custom',
      { ...BASE, profile: { api: 'anthropic' } },
      'sys',
      [],
      [],
      100,
      cb,
    ).catch(() => {})
    expect(fetchMock).toHaveBeenCalledWith('https://gw.example.com/v1/messages', expect.anything())
    const body = JSON.parse(fetchMock.mock.calls[0]![1].body as string)
    expect(body).toMatchObject({ model: 'step-3.7-flash', max_tokens: 100, stream: true })
  })

  it('speaks the Gemini generateContent route when declared', async () => {
    const fetchMock = vi.fn().mockResolvedValue(sse(['data: [DONE]']))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'custom',
      { ...BASE, profile: { api: 'gemini' } },
      'sys',
      [],
      [],
      100,
      cb,
    ).catch(() => {})
    expect(fetchMock).toHaveBeenCalledWith(
      'https://gw.example.com/v1/models/step-3.7-flash:streamGenerateContent?alt=sse',
      expect.anything(),
    )
  })

  it('sends declared headers on the wire', async () => {
    const fetchMock = vi.fn().mockResolvedValue(sse(['data: [DONE]']))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'custom',
      {
        ...BASE,
        profile: { api: 'anthropic', headers: { 'X-Tenant': 'office', 'X-Route': 'eu' } },
      },
      'sys',
      [],
      [],
      100,
      cb,
    ).catch(() => {})
    const headers = fetchMock.mock.calls[0]![1].headers as Record<string, string>
    expect(headers['X-Tenant']).toBe('office')
    expect(headers['X-Route']).toBe('eu')
    // the protocol's own auth header still rides along
    expect(headers['x-api-key']).toBe('k')
  })

  it('honors a per-model output ceiling', async () => {
    const fetchMock = vi.fn().mockResolvedValue(sse(['data: [DONE]']))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'custom',
      {
        ...BASE,
        profile: { api: 'anthropic', modelProfiles: [{ id: 'step-3.7-flash', maxTokens: 8192 }] },
      },
      'sys',
      [],
      [],
      32768,
      cb,
    ).catch(() => {})
    const body = JSON.parse(fetchMock.mock.calls[0]![1].body as string)
    expect(body.max_tokens).toBe(8192)
  })

  it('sends the renamed output field when declared', async () => {
    const fetchMock = vi.fn().mockResolvedValue(sse(['data: [DONE]']))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'custom',
      { ...BASE, profile: { compat: { useMaxCompletionTokens: true } } },
      'sys',
      [],
      [],
      100,
      cb,
    ).catch(() => {})
    const body = JSON.parse(fetchMock.mock.calls[0]![1].body as string)
    expect(body).toMatchObject({ max_completion_tokens: 100 })
    expect(body.max_tokens).toBeUndefined()
  })

  it('drops the temperature field when the endpoint declares it', async () => {
    const fetchMock = vi.fn().mockResolvedValue(sse(['data: [DONE]']))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'custom',
      { ...BASE, profile: { compat: { omitTemperature: true } } },
      'sys',
      [],
      [],
      100,
      cb,
    ).catch(() => {})
    const body = JSON.parse(fetchMock.mock.calls[0]![1].body as string)
    expect(body.temperature).toBeUndefined()
  })

  it('routes a mirror base URL to a gateway under a path prefix', async () => {
    const fetchMock = vi.fn().mockResolvedValue(sse(['data: [DONE]']))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'custom',
      { ...BASE, baseUrl: 'https://mirror.internal/api/anthropic', profile: { api: 'anthropic' } },
      'sys',
      [],
      [],
      100,
      cb,
    ).catch(() => {})
    expect(fetchMock).toHaveBeenCalledWith(
      'https://mirror.internal/api/anthropic/v1/messages',
      expect.anything(),
    )
  })
})

describe('chatForProvider: custom endpoint protocol dispatch', () => {
  it('speaks the declared protocol for the one-shot test path', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ content: [{ type: 'text', text: 'OK' }] }))
    vi.stubGlobal('fetch', fetchMock)
    const r = await chatForProvider(
      'custom',
      { ...BASE, profile: { api: 'anthropic' } },
      'sys',
      'ping',
    )
    expect(r.ok).toBe(true)
    expect(fetchMock).toHaveBeenCalledWith(
      'https://gw.example.com/v1/messages',
      expect.objectContaining({ headers: expect.objectContaining({ 'x-api-key': 'k' }) }),
    )
  })

  it('reports a missing Base URL as a failed reply, not a rejection', async () => {
    const r = await chatForProvider('custom', { apiKey: 'k', model: 'm' }, 'sys', 'ping')
    expect(r.ok).toBe(false)
    expect(r.error).toContain('Base URL')
  })

  it('carries declared headers into the one-shot test path', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ choices: [{ message: { content: 'OK' } }] }))
    vi.stubGlobal('fetch', fetchMock)
    await chatForProvider(
      'custom',
      { ...BASE, profile: { headers: { 'X-Tenant': 'office' } } },
      'sys',
      'ping',
    )
    const headers = fetchMock.mock.calls[0]![1].headers as Record<string, string>
    expect(headers['X-Tenant']).toBe('office')
  })
})

describe('custom profile end to end through settings', () => {
  it('a settings payload routes through the slot mirror', async () => {
    const settings: AiSettings = {
      provider: 'custom',
      providers: {
        custom: {
          apiKey: 'k',
          model: 'kimi-k3',
          baseUrl: 'https://gw.example.com/v1',
          profile: {
            api: 'anthropic',
            headers: { 'X-Tenant': 'office' },
            compat: { omitTemperature: true },
          },
        },
      },
    } as unknown as AiSettings
    const fetchMock = vi.fn().mockResolvedValue(sse(['data: [DONE]']))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    const config = settings.providers.custom
    await streamForProvider('custom', config, 'sys', [], [], 100, cb).catch(() => {})
    expect(fetchMock).toHaveBeenCalledWith(
      'https://gw.example.com/v1/messages',
      expect.objectContaining({
        headers: expect.objectContaining({ 'X-Tenant': 'office', 'x-api-key': 'k' }),
      }),
    )
  })
})
