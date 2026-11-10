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

describe('streamForProvider: anthropic block index', () => {
  it('attributes index-less delta/stop events to the last started block (sequential parallel tools)', async () => {
    // A gateway that omits the optional block index on delta/stop events;
    // tools are started and completed one at a time (sequential)
    const body = sseStream([
      'data: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"t1","name":"read"}}',
      'data: {"type":"content_block_delta","delta":{"type":"input_json_delta","partial_json":"{\\"path\\":\\"a\\"}"}}',
      'data: {"type":"content_block_stop"}',
      'data: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"t2","name":"write"}}',
      'data: {"type":"content_block_delta","delta":{"type":"input_json_delta","partial_json":"{\\"path\\":\\"b\\"}"}}',
      'data: {"type":"content_block_stop"}',
      'data: {"type":"message_delta","delta":{"stop_reason":"tool_use"}}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { toolCalls, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(toolCalls).toEqual([
      { id: 't1', name: 'read', input: { path: 'a' } },
      { id: 't2', name: 'write', input: { path: 'b' } },
    ])
  })

  it('still honors an explicit index on delta/stop events', async () => {
    const body = sseStream([
      'data: {"type":"content_block_start","index":0,"content_block":{"type":"tool_use","id":"t1","name":"read"}}',
      'data: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"t2","name":"write"}}',
      'data: {"type":"content_block_delta","index":1,"delta":{"type":"input_json_delta","partial_json":"{\\"path\\":\\"b\\"}"}}',
      'data: {"type":"content_block_stop","index":1}',
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"input_json_delta","partial_json":"{\\"path\\":\\"a\\"}"}}',
      'data: {"type":"content_block_stop","index":0}',
      'data: {"type":"message_delta","delta":{"stop_reason":"tool_use"}}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { toolCalls, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(toolCalls).toEqual([
      { id: 't2', name: 'write', input: { path: 'b' } },
      { id: 't1', name: 'read', input: { path: 'a' } },
    ])
  })
})
