import { afterEach, describe, expect, it, vi } from 'vitest'
import { streamForProvider } from '../src/stream'
import { okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

const TOOL = [{ name: 'read_file', description: 'read a file', inputSchema: { type: 'object' } }]

const sse = () =>
  sseStream([
    'data: {"choices":[{"delta":{"content":"ok"},"finish_reason":"stop"}]}',
    'data: [DONE]',
  ])

function collector() {
  return {
    cb: {
      signal: new AbortController().signal,
      onDelta: vi.fn(),
      onToolCall: vi.fn(),
      onStopReason: vi.fn(),
    },
  }
}

describe('streamForProvider: reasoning_effort beside function tools', () => {
  it('pins effort none for a gpt-5.6 tool turn', async () => {
    const fetchMock = vi.fn().mockImplementation(async () => okResponse(sse()))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-5.6-terra' },
      'sys',
      [],
      TOOL,
      100,
      cb,
    )
    const body = JSON.parse(fetchMock.mock.calls[0][1].body as string)
    expect(body.tools).toHaveLength(1)
    expect(body.reasoning_effort).toBe('none')
  })

  it('leaves other models and tool-less turns alone', async () => {
    const fetchMock = vi.fn().mockImplementation(async () => okResponse(sse()))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider('openai', { apiKey: 'k', model: 'gpt-5.5' }, 'sys', [], TOOL, 100, cb)
    expect(JSON.parse(fetchMock.mock.calls[0][1].body as string).reasoning_effort).toBeUndefined()
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-5.6-terra' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(JSON.parse(fetchMock.mock.calls[1][1].body as string).reasoning_effort).toBeUndefined()
  })
})
