import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentMessage, AgentToolCall } from '@genoffice/agent-core'
import { streamForProvider } from '../src/stream'
import { okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

function collector() {
  const deltas: string[] = []
  const toolCalls: AgentToolCall[] = []
  const stopReasons: string[] = []
  return {
    deltas,
    toolCalls,
    stopReasons,
    cb: {
      signal: new AbortController().signal,
      onDelta: (text: string) => deltas.push(text),
      onToolCall: (call: AgentToolCall) => toolCalls.push(call),
      onStopReason: (reason: string) => stopReasons.push(reason),
    },
  }
}

const okTurn = () =>
  okResponse(
    sseStream([
      'data: {"choices":[{"delta":{"content":"hi"},"finish_reason":"stop"}]}',
      'data: [DONE]',
    ]),
  )

describe('streamForProvider: openai-compatible isError', () => {
  it('prefixes failed tool results with an error marker so the model can retry', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okTurn())
    vi.stubGlobal('fetch', fetchMock)
    const history: AgentMessage[] = [
      { role: 'user', text: 'do the thing' },
      {
        role: 'assistant',
        text: '',
        toolCalls: [{ id: 't1', name: 'do_thing', input: { a: 1 } }],
      },
      {
        role: 'tool',
        results: [{ id: 't1', name: 'do_thing', output: 'boom', isError: true }],
      },
    ]
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'm' },
      'sys',
      history,
      [],
      100,
      collector().cb,
    )
    const body = JSON.parse(fetchMock.mock.calls[0]![1].body as string) as {
      messages: Array<{ role: string; content?: string }>
    }
    const toolMsg = body.messages.find((m) => m.role === 'tool')
    expect(toolMsg?.content).toBe('Error: boom')
  })

  it('leaves successful tool results unprefixed', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okTurn())
    vi.stubGlobal('fetch', fetchMock)
    const history: AgentMessage[] = [
      { role: 'user', text: 'do the thing' },
      {
        role: 'assistant',
        text: '',
        toolCalls: [{ id: 't1', name: 'do_thing', input: { a: 1 } }],
      },
      { role: 'tool', results: [{ id: 't1', name: 'do_thing', output: 'all good' }] },
    ]
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'm' },
      'sys',
      history,
      [],
      100,
      collector().cb,
    )
    const body = JSON.parse(fetchMock.mock.calls[0]![1].body as string) as {
      messages: Array<{ role: string; content?: string }>
    }
    const toolMsg = body.messages.find((m) => m.role === 'tool')
    expect(toolMsg?.content).toBe('all good')
  })
})
