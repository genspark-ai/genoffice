import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentToolCall } from '@genoffice/agent-core'
import { streamForProvider } from '../src/stream'
import { okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

function collector() {
  const toolCalls: AgentToolCall[] = []
  return {
    toolCalls,
    cb: {
      signal: new AbortController().signal,
      onDelta: vi.fn(),
      onToolCall: (call: AgentToolCall) => toolCalls.push(call),
      onStopReason: vi.fn(),
    },
  }
}

describe('streamForProvider: tool-name parity', () => {
  it('anthropic: drops a tool_use block with an empty name', async () => {
    const body = sseStream([
      'data: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"t1","name":""}}',
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"input_json_delta","partial_json":"{}"}}',
      'data: {"type":"content_block_stop","index":0}',
      'data: {"type":"message_delta","delta":{"stop_reason":"tool_use"}}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { toolCalls, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(toolCalls).toEqual([])
  })

  it('openai: drops a tool call whose name fragments to empty', async () => {
    const body = sseStream([
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"c1","function":{"name":""}}]}}]}',
      'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}',
      'data: [DONE]',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { toolCalls, cb } = collector()
    await streamForProvider('openai', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(toolCalls).toEqual([])
  })

  it('anthropic: still emits a named tool call', async () => {
    const body = sseStream([
      'data: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"t1","name":"read"}}',
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"input_json_delta","partial_json":"{}"}}',
      'data: {"type":"content_block_stop","index":0}',
      'data: {"type":"message_delta","delta":{"stop_reason":"tool_use"}}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { toolCalls, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(toolCalls).toEqual([{ id: 't1', name: 'read', input: {} }])
  })

  it('openai: still emits a named tool call', async () => {
    const body = sseStream([
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"c1","function":{"name":"read"}}]}}]}',
      'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}',
      'data: [DONE]',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { toolCalls, cb } = collector()
    await streamForProvider('openai', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(toolCalls).toEqual([{ id: 'c1', name: 'read', input: {} }])
  })
})
