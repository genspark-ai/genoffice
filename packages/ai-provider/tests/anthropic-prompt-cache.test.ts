import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentMessage } from '@genoffice/agent-core'
import { streamAnthropic } from '../src/protocols/anthropic'
import { GENSPARK_LLM_BASE_URLS } from '../src/providers'
import { okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

const turn = () =>
  sseStream([
    'data: {"type":"content_block_delta","delta":{"type":"text_delta","text":"hi"}}',
    'data: {"type":"message_delta","delta":{"stop_reason":"end_turn"}}',
  ])

const cb = () => ({
  signal: new AbortController().signal,
  onDelta: vi.fn(),
  onToolCall: vi.fn(),
})

const history: AgentMessage[] = [
  { role: 'user', text: 'first ask' },
  { role: 'assistant', text: 'ok', toolCalls: [{ id: 't1', name: 'read', input: {} }] },
  { role: 'tool', results: [{ id: 't1', name: 'read', output: 'data' }] },
]

async function requestBody(baseUrl?: string): Promise<Record<string, unknown>> {
  const fetchMock = vi.fn().mockResolvedValue(okResponse(turn()))
  vi.stubGlobal('fetch', fetchMock)
  await streamAnthropic({ apiKey: 'k', model: 'm' }, 'sys', history, [], 100, cb(), baseUrl)
  return JSON.parse(fetchMock.mock.calls[0]![1].body as string) as Record<string, unknown>
}

describe('anthropic prompt-cache breakpoints', () => {
  it('marks system and the last two messages on the official endpoint', async () => {
    const body = await requestBody()
    expect(body['system']).toEqual([
      { type: 'text', text: 'sys', cache_control: { type: 'ephemeral' } },
    ])
    const messages = body['messages'] as Array<{ content: unknown }>
    // earlier messages stay untouched (plain-string user text keeps its shape)
    expect(messages[0]).toEqual({ role: 'user', content: 'first ask' })
    // assistant turn: marker on its final block (the tool_use)
    expect(messages[1]!.content).toEqual([
      { type: 'text', text: 'ok' },
      {
        type: 'tool_use',
        id: 't1',
        name: 'read',
        input: {},
        cache_control: { type: 'ephemeral' },
      },
    ])
    // tool results: marker on the final tool_result block
    expect(messages[2]!.content).toEqual([
      {
        type: 'tool_result',
        tool_use_id: 't1',
        content: 'data',
        cache_control: { type: 'ephemeral' },
      },
    ])
  })

  it('marks requests to the Genspark proxy', async () => {
    const body = await requestBody(GENSPARK_LLM_BASE_URLS.anthropic)
    expect(JSON.stringify(body)).toContain('cache_control')
  })

  it('marks the official endpoint when the stored URL carries a trailing slash', async () => {
    const body = await requestBody('https://api.anthropic.com/')
    expect(JSON.stringify(body)).toContain('cache_control')
  })

  it('sends no cache_control to a user-hosted Messages endpoint', async () => {
    const body = await requestBody('http://localhost:11434')
    expect(body['system']).toBe('sys')
    expect(JSON.stringify(body)).not.toContain('cache_control')
  })

  it('upgrades a plain-string final user message to a marked text block', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(turn()))
    vi.stubGlobal('fetch', fetchMock)
    const single: AgentMessage[] = [{ role: 'user', text: 'only ask' }]
    await streamAnthropic({ apiKey: 'k', model: 'm' }, 'sys', single, [], 100, cb())
    const body = JSON.parse(fetchMock.mock.calls[0]![1].body as string) as {
      messages: Array<{ content: unknown }>
    }
    expect(body.messages[0]!.content).toEqual([
      { type: 'text', text: 'only ask', cache_control: { type: 'ephemeral' } },
    ])
  })
})
