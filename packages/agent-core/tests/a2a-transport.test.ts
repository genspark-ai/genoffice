import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest'
import { createA2ATransport, type A2ATransportOptions } from '../src/a2a-transport'
import type { AgentStreamCallbacks, AgentStreamRequest } from '../src/types'

// ---- helpers ----

function makeRequest(overrides?: Partial<AgentStreamRequest>): AgentStreamRequest {
  return {
    system: 'You are a helpful assistant.',
    messages: [{ role: 'user', text: 'Hello' }],
    tools: [],
    ...overrides,
  }
}

function makeCb(): AgentStreamCallbacks & { calls: string[] } {
  const calls: string[] = []
  return {
    calls,
    onDelta: vi.fn((text: string) => calls.push(`delta:${text}`)),
    onToolCall: vi.fn(),
    onDone: vi.fn(() => calls.push('done')),
    onError: vi.fn((err: string) => calls.push(`error:${err}`)),
  }
}

/** Build a minimal SSE body from an array of JSON-RPC result payloads */
function sseBody(events: object[]): ReadableStream<Uint8Array> {
  const enc = new TextEncoder()
  const chunks = events.map((e) => enc.encode(`data: ${JSON.stringify(e)}\n\n`))
  let i = 0
  return new ReadableStream({
    pull(controller) {
      if (i < chunks.length) {
        controller.enqueue(chunks[i++]!)
      } else {
        controller.close()
      }
    },
  })
}

function taskEvent(id: string, result: object): { jsonrpc: '2.0'; id: string; result: object } {
  return { jsonrpc: '2.0', id, result }
}

function statusEvent(id: string, state: string, final = false) {
  return taskEvent(id, { id, status: { state, timestamp: '2026-01-01T00:00:00Z' }, final })
}

function artifactEvent(id: string, text: string, lastChunk = false) {
  return taskEvent(id, {
    id,
    artifact: { name: 'response', parts: [{ type: 'text', text }], lastChunk },
  })
}

const flush = () => new Promise<void>((r) => setTimeout(r, 10))

// ---- tests ----

describe('createA2ATransport', () => {
  const fetchMock = vi.fn<typeof fetch>()

  beforeEach(() => {
    vi.stubGlobal('fetch', fetchMock)
    vi.stubGlobal('crypto', { randomUUID: () => 'test-uuid-1234' })
    fetchMock.mockReset()
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  function makeTransport(opts?: Partial<A2ATransportOptions>) {
    return createA2ATransport({ agentUrl: 'http://agent.local:9100', ...opts })
  }

  it('sends a POST with correct JSON-RPC body and headers', async () => {
    fetchMock.mockResolvedValueOnce(
      new Response(sseBody([statusEvent('test-uuid-1234', 'completed')]), {
        status: 200,
        headers: { 'Content-Type': 'text/event-stream' },
      }),
    )

    const transport = makeTransport()
    const cb = makeCb()
    transport.stream(makeRequest(), cb)
    await flush()

    expect(fetchMock).toHaveBeenCalledOnce()
    const [url, init] = fetchMock.mock.calls[0]!
    expect(url).toBe('http://agent.local:9100')
    expect((init?.headers as Record<string, string>)['Content-Type']).toBe('application/json')
    expect((init?.headers as Record<string, string>)['Accept']).toBe('text/event-stream')

    const body = JSON.parse(init?.body as string)
    expect(body.method).toBe('tasks/sendSubscribe')
    expect(body.params.id).toBe('test-uuid-1234')
    expect(body.params.message.role).toBe('user')
    expect(body.params.message.parts[0]).toMatchObject({
      type: 'text',
      text: expect.stringContaining('You are a helpful assistant.'),
    })
  })

  it('includes Authorization header when apiKey is set', async () => {
    fetchMock.mockResolvedValueOnce(
      new Response(sseBody([statusEvent('test-uuid-1234', 'completed')]), {
        status: 200,
        headers: { 'Content-Type': 'text/event-stream' },
      }),
    )

    const transport = makeTransport({ apiKey: 'sk-test-key' })
    transport.stream(makeRequest(), makeCb())
    await flush()

    const [, init] = fetchMock.mock.calls[0]!
    expect((init?.headers as Record<string, string>)['Authorization']).toBe('Bearer sk-test-key')
  })

  it('includes sessionId in params when provided', async () => {
    fetchMock.mockResolvedValueOnce(
      new Response(sseBody([statusEvent('test-uuid-1234', 'completed')]), {
        status: 200,
        headers: { 'Content-Type': 'text/event-stream' },
      }),
    )

    const transport = makeTransport({ sessionId: 'session-abc' })
    transport.stream(makeRequest(), makeCb())
    await flush()

    const [, init] = fetchMock.mock.calls[0]!
    const body = JSON.parse(init?.body as string)
    expect(body.params.sessionId).toBe('session-abc')
  })

  it('emits onDelta for artifact text parts and onDone on lastChunk', async () => {
    fetchMock.mockResolvedValueOnce(
      new Response(
        sseBody([
          artifactEvent('test-uuid-1234', 'Hello '),
          artifactEvent('test-uuid-1234', 'world', true),
        ]),
        { status: 200, headers: { 'Content-Type': 'text/event-stream' } },
      ),
    )

    const cb = makeCb()
    makeTransport().stream(makeRequest(), cb)
    await flush()

    expect(cb.calls).toEqual(['delta:Hello ', 'delta:world', 'done'])
    expect(cb.onDone).toHaveBeenCalledOnce()
    expect(cb.onError).not.toHaveBeenCalled()
  })

  it('emits onDone on status completed with no artifact', async () => {
    fetchMock.mockResolvedValueOnce(
      new Response(sseBody([statusEvent('test-uuid-1234', 'completed')]), {
        status: 200,
        headers: { 'Content-Type': 'text/event-stream' },
      }),
    )

    const cb = makeCb()
    makeTransport().stream(makeRequest(), cb)
    await flush()

    expect(cb.onDone).toHaveBeenCalledOnce()
    expect(cb.onError).not.toHaveBeenCalled()
  })

  it('emits onError when task state is failed', async () => {
    const failedEvent = taskEvent('test-uuid-1234', {
      id: 'test-uuid-1234',
      status: {
        state: 'failed',
        message: { role: 'agent', parts: [{ type: 'text', text: 'out of credits' }] },
      },
    })
    fetchMock.mockResolvedValueOnce(
      new Response(sseBody([failedEvent]), {
        status: 200,
        headers: { 'Content-Type': 'text/event-stream' },
      }),
    )

    const cb = makeCb()
    makeTransport().stream(makeRequest(), cb)
    await flush()

    expect(cb.calls).toContain('error:out of credits')
    expect(cb.onDone).not.toHaveBeenCalled()
  })

  it('emits onError on JSON-RPC error object', async () => {
    const rpcError = {
      jsonrpc: '2.0',
      id: 'test-uuid-1234',
      error: { code: -32600, message: 'Invalid Request' },
    }
    fetchMock.mockResolvedValueOnce(
      new Response(sseBody([rpcError]), {
        status: 200,
        headers: { 'Content-Type': 'text/event-stream' },
      }),
    )

    const cb = makeCb()
    makeTransport().stream(makeRequest(), cb)
    await flush()

    expect(cb.calls).toContain('error:Invalid Request')
  })

  it('emits onError on non-200 HTTP response', async () => {
    fetchMock.mockResolvedValueOnce(new Response('Unauthorized', { status: 401 }))

    const cb = makeCb()
    makeTransport().stream(makeRequest(), cb)
    await flush()

    expect(cb.calls).toContain('error:A2A agent returned HTTP 401')
  })

  it('emits onError on fetch network failure', async () => {
    fetchMock.mockRejectedValueOnce(new Error('ECONNREFUSED'))

    const cb = makeCb()
    makeTransport().stream(makeRequest(), cb)
    await flush()

    expect(cb.calls).toContain('error:ECONNREFUSED')
  })

  it('cancel() aborts the in-flight request', async () => {
    let capturedSignal: AbortSignal | undefined
    fetchMock.mockImplementation((_url, init) => {
      capturedSignal = init?.signal as AbortSignal
      // never resolves — simulates a stalled agent
      return new Promise(() => undefined)
    })

    const cb = makeCb()
    const handle = makeTransport().stream(makeRequest(), cb)
    handle.cancel()
    await flush()

    expect(capturedSignal?.aborted).toBe(true)
    expect(cb.onDone).not.toHaveBeenCalled()
    expect(cb.onError).not.toHaveBeenCalled()
  })

  it('encodes multi-turn history into the A2A message parts', async () => {
    fetchMock.mockResolvedValueOnce(
      new Response(sseBody([statusEvent('test-uuid-1234', 'completed')]), {
        status: 200,
        headers: { 'Content-Type': 'text/event-stream' },
      }),
    )

    const request = makeRequest({
      messages: [
        { role: 'user', text: 'First question' },
        { role: 'assistant', text: 'First answer' },
        { role: 'user', text: 'Follow-up' },
      ],
    })

    makeTransport().stream(request, makeCb())
    await flush()

    const [, init] = fetchMock.mock.calls[0]!
    const body = JSON.parse(init?.body as string)
    const parts: Array<{ type: string; text?: string; data?: unknown }> = body.params.message.parts

    // system + 3 message parts
    expect(parts.length).toBe(4)
    expect(parts[0]).toMatchObject({
      type: 'text',
      text: expect.stringContaining('You are a helpful assistant.'),
    })
    expect(parts[1]).toMatchObject({ type: 'text', text: 'First question' })
    expect(parts[2]).toMatchObject({
      type: 'data',
      data: { role: 'assistant', text: 'First answer' },
    })
    expect(parts[3]).toMatchObject({ type: 'text', text: 'Follow-up' })
  })

  it('strips trailing slash from agentUrl', async () => {
    fetchMock.mockResolvedValueOnce(
      new Response(sseBody([statusEvent('test-uuid-1234', 'completed')]), {
        status: 200,
        headers: { 'Content-Type': 'text/event-stream' },
      }),
    )

    makeA2ATransport({ agentUrl: 'http://agent.local:9100/' }).stream(makeRequest(), makeCb())
    await flush()

    const [url] = fetchMock.mock.calls[0]!
    expect(url).toBe('http://agent.local:9100')
  })
})

// local alias so the test above can reference a clean factory after global setup
function makeA2ATransport(opts: A2ATransportOptions) {
  return createA2ATransport(opts)
}
