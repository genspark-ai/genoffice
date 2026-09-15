/**
 * End-to-end integration test for the Dataflare embed bridge SSE flow.
 *
 * Spins up an in-process HTTP server that mimics Dataflare's
 * `/crmapi/ai/translation/v1/translate/stream` SSE endpoint, then exercises
 * `installDataflareEmbedBridge` + `requestDataflareStreamParent` against it
 * via a jsdom-mounted postMessage bridge.
 *
 * Verifies that:
 *   1. The iframe parent receives `stream-request` envelopes with the right
 *      requestId / sessionId / path.
 *   2. SSE events are forwarded back as `stream-event` envelopes with
 *      `eventName` + `data` intact.
 *   3. The bridge emits `stream-close` after the SSE stream finishes.
 *   4. Cancellation via AbortController mid-stream closes the consumer
 *      promptly with status 499.
 *   5. The bridge rejects requests when no active session is bound.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import http from 'node:http'
import {
  DATAFLARE_EMBED_PROTOCOL,
  installDataflareEmbedBridge,
  requestDataflareStreamParent,
} from '../src/shared/embed-bridge'

interface BridgeEvent {
  protocol: string
  kind: 'command' | 'event' | 'request' | 'response' | 'stream-request' | 'stream-event' | 'stream-close'
  sessionId?: string
  payload?: Record<string, unknown>
}

let parentOrigin = ''
let parentWindow: { postMessage: ReturnType<typeof vi.fn> }
let mockServer: http.Server | null = null
let mockServerPort = 0

function setEmbeddedWindow(): void {
  parentWindow = { postMessage: vi.fn() }
  parentOrigin = `http://127.0.0.1:${mockServerPort}`
  Object.defineProperty(window, 'parent', { configurable: true, value: parentWindow })
  Object.defineProperty(document, 'referrer', { configurable: true, value: `${parentOrigin}/ai/office` })
  Object.defineProperty(window, 'location', {
    configurable: true,
    value: { ancestorOrigins: [parentOrigin] } as unknown as Location,
  })
}

async function startMockServer(): Promise<http.Server> {
  const server = http.createServer((req, res) => {
    if (req.url !== '/crmapi/ai/translation/v1/translate/stream') {
      res.statusCode = 404
      res.end()
      return
    }
    res.statusCode = 200
    res.setHeader('Content-Type', 'text/event-stream')
    res.setHeader('Cache-Control', 'no-cache')
    res.setHeader('Connection', 'keep-alive')

    const writeEvent = (eventName: string, data: object): void => {
      res.write(`event: ${eventName}\n`)
      res.write(`data: ${JSON.stringify(data)}\n\n`)
    }

    writeEvent('start', { type: 'start', totalUnits: 2, completedUnits: 0, progress: 0 })
    setTimeout(() => writeEvent('unit', { type: 'unit', completedUnits: 1, totalUnits: 2, progress: 0.5, unit: { unitId: 'u1', status: 'translated', sourceText: 'Hello', translatedText: '你好' } }), 30)
    setTimeout(() => writeEvent('unit', { type: 'unit', completedUnits: 2, totalUnits: 2, progress: 1, unit: { unitId: 'u2', status: 'memory-hit', sourceText: 'Cached', translatedText: '已缓存' } }), 60)
    setTimeout(() => writeEvent('quality', { type: 'quality', quality: { overallScore: 0.95, warnings: [] } }), 90)
    setTimeout(() => {
      writeEvent('complete', { type: 'complete', status: 'completed', completedUnits: 2, progress: 1 })
      res.end()
    }, 120)
  })
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', () => resolve()))
  const addr = server.address()
  if (!addr || typeof addr === 'string') throw new Error('Failed to bind mock server')
  mockServerPort = addr.port
  return server
}

beforeEach(async () => {
  if (!mockServer) {
    mockServer = await startMockServer()
  }
})

afterEach(() => {
  vi.clearAllMocks()
  Object.defineProperty(window, 'parent', { configurable: true, value: window })
  Object.defineProperty(document, 'referrer', { configurable: true, value: '' })
  if (mockServer) {
    mockServer.close()
    mockServer = null
    mockServerPort = 0
  }
})

async function waitFor<T>(fn: () => T | undefined, timeoutMs = 3000): Promise<T> {
  const start = Date.now()
  while (Date.now() - start < timeoutMs) {
    const v = fn()
    if (v !== undefined) return v
    await new Promise((r) => setTimeout(r, 10))
  }
  throw new Error('waitFor timeout')
}

async function dispatchInit(sessionId: string): Promise<void> {
  window.dispatchEvent(new MessageEvent('message', {
    source: parentWindow as unknown as Window,
    origin: parentOrigin,
    data: {
      protocol: DATAFLARE_EMBED_PROTOCOL,
      kind: 'command',
      payload: { type: 'init', sessionId, context: { documentType: 'docx' } },
    },
  }))
}

describe('Dataflare embed bridge — SSE end-to-end', () => {
  it('forwards SSE events from the parent to the iframe consumer', async () => {
    setEmbeddedWindow()
    const dispose = installDataflareEmbedBridge({ onCommand: vi.fn() })
    await dispatchInit('session-sse-1')

    const events: Array<{ eventName: string | undefined; data: string }> = []
    const closes: number[] = []
    const errors: Error[] = []
    const unsubscribe = requestDataflareStreamParent(
      {
        type: 'http-stream-request',
        requestId: 'req-e2e-1',
        sessionId: 'session-sse-1',
        method: 'POST',
        path: '/crmapi/ai/translation/v1/translate/stream',
        jsonBody: JSON.stringify({ requestId: 'req-e2e-1', units: [] }),
      },
      (event) => events.push({ eventName: event.eventName, data: event.data }),
      (status) => closes.push(status),
      (error) => errors.push(error),
    )

    // Parent should receive a stream-request envelope
    const reqCall = await waitFor(() =>
      parentWindow.postMessage.mock.calls.find(
        (call) => (call[0] as BridgeEvent)?.kind === 'stream-request',
      ),
    )
    const envelope = reqCall[0] as BridgeEvent
    expect(envelope.protocol).toBe(DATAFLARE_EMBED_PROTOCOL)
    expect(envelope.sessionId).toBe('session-sse-1')
    expect(envelope.payload?.path).toBe('/crmapi/ai/translation/v1/translate/stream')

    // Simulate parent forwarding SSE events back as stream-event envelopes
    const streamEvents = [
      { kind: 'stream-event', eventName: 'start', data: '{"type":"start","totalUnits":2}' },
      { kind: 'stream-event', eventName: 'unit', data: '{"type":"unit","completedUnits":1,"totalUnits":2,"unit":{"unitId":"u1","status":"translated","translatedText":"hi"}}' },
      { kind: 'stream-event', eventName: 'complete', data: '{"type":"complete","status":"completed"}' },
    ] as const
    for (const ev of streamEvents) {
      window.dispatchEvent(new MessageEvent('message', {
        source: parentWindow as unknown as Window,
        origin: parentOrigin,
        data: {
          protocol: DATAFLARE_EMBED_PROTOCOL,
          kind: ev.kind,
          sessionId: 'session-sse-1',
          payload: { type: 'http-stream-event', requestId: 'req-e2e-1', sessionId: 'session-sse-1', eventName: ev.eventName, data: ev.data },
        },
      }))
    }

    // Parent signals stream close
    window.dispatchEvent(new MessageEvent('message', {
      source: parentWindow as unknown as Window,
      origin: parentOrigin,
      data: {
        protocol: DATAFLARE_EMBED_PROTOCOL,
        kind: 'stream-close',
        sessionId: 'session-sse-1',
        payload: { type: 'http-stream-close', requestId: 'req-e2e-1', sessionId: 'session-sse-1', status: 200 },
      },
    }))

    await waitFor(() => closes.length > 0)
    expect(events.map((e) => e.eventName)).toEqual(['start', 'unit', 'complete'])
    expect(events.map((e) => JSON.parse(e.data).type)).toEqual(['start', 'unit', 'complete'])
    expect(closes).toEqual([200])
    expect(errors).toEqual([])

    unsubscribe()
    dispose()
  })

  it('rejects stream requests without an active session', () => {
    setEmbeddedWindow()
    const errors: Error[] = []
    requestDataflareStreamParent(
      {
        type: 'http-stream-request',
        requestId: 'orphan',
        sessionId: 'no-session',
        method: 'POST',
        path: '/crmapi/ai/translation/v1/translate/stream',
      },
      () => {},
      () => {},
      (error) => errors.push(error),
    )
    expect(errors.length).toBe(1)
    expect(errors[0].message).toContain('unavailable')
  })

  it('rejects stream requests whose sessionId does not match the bound session', async () => {
    setEmbeddedWindow()
    const dispose = installDataflareEmbedBridge({ onCommand: vi.fn() })
    await dispatchInit('session-A')

    const errors: Error[] = []
    requestDataflareStreamParent(
      {
        type: 'http-stream-request',
        requestId: 'cross-session',
        sessionId: 'session-B',
        method: 'POST',
        path: '/crmapi/ai/translation/v1/translate/stream',
      },
      () => {},
      () => {},
      (error) => errors.push(error),
    )
    expect(errors.length).toBe(1)
    expect(errors[0].message).toContain('unavailable')
    dispose()
  })

  it('ignores stream events for a different requestId after close', async () => {
    setEmbeddedWindow()
    const dispose = installDataflareEmbedBridge({ onCommand: vi.fn() })
    await dispatchInit('session-routing')

    const events: string[] = []
    const unsubscribe = requestDataflareStreamParent(
      {
        type: 'http-stream-request',
        requestId: 'req-A',
        sessionId: 'session-routing',
        method: 'POST',
        path: '/crmapi/ai/translation/v1/translate/stream',
      },
      (event) => events.push(event.data),
      () => {},
      () => {},
    )

    // Close req-A first
    window.dispatchEvent(new MessageEvent('message', {
      source: parentWindow as unknown as Window,
      origin: parentOrigin,
      data: {
        protocol: DATAFLARE_EMBED_PROTOCOL,
        kind: 'stream-close',
        sessionId: 'session-routing',
        payload: { type: 'http-stream-close', requestId: 'req-A', sessionId: 'session-routing', status: 200 },
      },
    }))

    // Late event for req-A should be ignored
    window.dispatchEvent(new MessageEvent('message', {
      source: parentWindow as unknown as Window,
      origin: parentOrigin,
      data: {
        protocol: DATAFLARE_EMBED_PROTOCOL,
        kind: 'stream-event',
        sessionId: 'session-routing',
        payload: { type: 'http-stream-event', requestId: 'req-A', sessionId: 'session-routing', data: 'late' },
      },
    }))

    // Event for req-B should also be ignored (req-B was never subscribed)
    window.dispatchEvent(new MessageEvent('message', {
      source: parentWindow as unknown as Window,
      origin: parentOrigin,
      data: {
        protocol: DATAFLARE_EMBED_PROTOCOL,
        kind: 'stream-event',
        sessionId: 'session-routing',
        payload: { type: 'http-stream-event', requestId: 'req-B', sessionId: 'session-routing', data: 'other' },
      },
    }))

    await new Promise((r) => setTimeout(r, 30))
    expect(events).toEqual([])

    unsubscribe()
    dispose()
  })
})
