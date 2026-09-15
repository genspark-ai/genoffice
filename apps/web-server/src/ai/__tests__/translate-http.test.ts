/**
 * HTTP integration tests for `/api/ai/translate`, `/api/ai/translate/stream`,
 * `/api/ai/translate/stream/cancel`. Spawns the real GenOffice web-server on
 * a random port and exercises the HTTP surface end-to-end.
 *
 * These tests deliberately cover only paths that do NOT hit a real AI provider:
 *   - empty units array
 *   - missing provider config
 *   - bad JSON body
 *   - cancel requestId miss
 * Happy-path translations are covered at the package level by
 * `@genoffice/translation-core` `tests/provider.test.ts` (translateBatchStream
 * has 5 dedicated tests including SSE-equivalent onUnit event ordering).
 *
 * Run with: `tsx --test src/ai/__tests__/translate-http.test.ts`
 */
import test from 'node:test'
import assert from 'node:assert/strict'

import { createServer, type Server, type IncomingMessage, type ServerResponse } from 'node:http'

import {
  handleTranslateBatchHttp,
  handleTranslateStreamCancelHttp,
  handleTranslateStreamHttp,
} from '../translate-http'

function makeServer(): Promise<{ server: Server; port: number; close: () => void }> {
  return new Promise((resolve) => {
    const server = createServer((req: IncomingMessage, res: ServerResponse) => {
      if (req.url === '/api/ai/translate' && req.method === 'POST') {
        void handleTranslateBatchHttp(req, res)
        return
      }
      if (req.url === '/api/ai/translate/stream' && req.method === 'POST') {
        void handleTranslateStreamHttp(req, res)
        return
      }
      if (req.url === '/api/ai/translate/stream/cancel' && req.method === 'POST') {
        void handleTranslateStreamCancelHttp(req, res)
        return
      }
      res.writeHead(404).end()
    })
    server.listen(0, '127.0.0.1', () => {
      const addr = server.address()
      if (addr && typeof addr === 'object') {
        resolve({
          server,
          port: addr.port,
          close: () => server.close(),
        })
      }
    })
  })
}

async function readSseEvents(response: Response): Promise<Array<{ name: string; data: unknown }>> {
  if (!response.body) throw new Error('no response body')
  const reader = response.body.getReader()
  const decoder = new TextDecoder()
  let buffer = ''
  const events: Array<{ name: string; data: unknown }> = []
  while (true) {
    const { value, done } = await reader.read()
    if (done) break
    buffer += decoder.decode(value, { stream: true })
    let sep: number
    while ((sep = buffer.indexOf('\n\n')) >= 0) {
      const frame = buffer.slice(0, sep)
      buffer = buffer.slice(sep + 2)
      const lines = frame.split('\n')
      let name = 'message'
      let data = ''
      for (const line of lines) {
        if (line.startsWith('event:')) name = line.slice(6).trim()
        else if (line.startsWith('data:')) data += line.slice(5).trim()
      }
      if (data) {
        try {
          events.push({ name, data: JSON.parse(data) })
        } catch {
          events.push({ name, data })
        }
      }
    }
  }
  return events
}

test('POST /api/ai/translate — empty units → 400 with PROVIDER_NOT_CONFIGURED', async () => {
  const { port, close } = await makeServer()
  try {
    const response = await fetch(`http://127.0.0.1:${port}/api/ai/translate`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        sourceLanguage: 'en-US',
        targetLanguage: 'zh-CN',
        // Force a provider that the server has no key for.
        settings: { provider: 'openai', providers: {} },
        units: [{ unitId: 'u1', sourceText: 'hello', order: 0 }],
      }),
    })
    assert.equal(response.status, 400)
    const body = (await response.json()) as { error?: { code?: string; message?: string } }
    assert.equal(body.error?.code, 'PROVIDER_NOT_CONFIGURED')
    assert.match(body.error?.message ?? '', /not configured/)
  } finally {
    close()
  }
})

test('POST /api/ai/translate — malformed JSON → 500 with wrapped error', async () => {
  const { port, close } = await makeServer()
  try {
    const response = await fetch(`http://127.0.0.1:${port}/api/ai/translate`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: '{not valid json',
    })
    assert.equal(response.status, 500)
    const body = (await response.json()) as { error?: { message?: string } }
    assert.ok(body.error?.message)
  } finally {
    close()
  }
})

test('POST /api/ai/translate/stream — empty units → SSE error event with requestId', async () => {
  const { port, close } = await makeServer()
  try {
    const response = await fetch(`http://127.0.0.1:${port}/api/ai/translate/stream`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        requestId: 'test-empty-units',
        sourceLanguage: 'en-US',
        targetLanguage: 'zh-CN',
        units: [],
      }),
    })
    assert.equal(response.status, 200)
    assert.equal(response.headers.get('content-type'), 'text/event-stream')
    const events = await readSseEvents(response)
    const errorEvent = events.find((e) => e.name === 'error')
    assert.ok(errorEvent, 'must emit an error event')
    const payload = errorEvent!.data as { type: string; requestId: string; message: string }
    assert.equal(payload.type, 'error')
    assert.equal(payload.requestId, 'test-empty-units')
    assert.match(payload.message, /non-empty/)
  } finally {
    close()
  }
})

test('POST /api/ai/translate/stream — missing provider → SSE error event', async () => {
  const { port, close } = await makeServer()
  try {
    const response = await fetch(`http://127.0.0.1:${port}/api/ai/translate/stream`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        requestId: 'test-no-provider',
        sourceLanguage: 'en-US',
        targetLanguage: 'zh-CN',
        settings: { provider: 'openai', providers: {} },
        units: [{ unitId: 'u1', sourceText: 'hello', order: 0 }],
      }),
    })
    const events = await readSseEvents(response)
    const errorEvent = events.find((e) => e.name === 'error')
    assert.ok(errorEvent, 'must emit an error event')
    const payload = errorEvent!.data as { type: string; requestId: string; message: string }
    assert.equal(payload.requestId, 'test-no-provider')
    assert.match(payload.message, /not configured/)
  } finally {
    close()
  }
})

test('POST /api/ai/translate/stream — bad JSON → 400', async () => {
  const { port, close } = await makeServer()
  try {
    const response = await fetch(`http://127.0.0.1:${port}/api/ai/translate/stream`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: '{malformed',
    })
    assert.equal(response.status, 400)
    const body = (await response.json()) as { error?: { message?: string } }
    assert.match(body.error?.message ?? '', /JSON/)
  } finally {
    close()
  }
})

test('POST /api/ai/translate/stream/cancel — missing requestId → 400', async () => {
  const { port, close } = await makeServer()
  try {
    const response = await fetch(`http://127.0.0.1:${port}/api/ai/translate/stream/cancel`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({}),
    })
    assert.equal(response.status, 400)
    const body = (await response.json()) as { error?: { message?: string } }
    assert.match(body.error?.message ?? '', /requestId/)
  } finally {
    close()
  }
})

test('POST /api/ai/translate/stream/cancel — unknown requestId → 200 with aborted:false', async () => {
  const { port, close } = await makeServer()
  try {
    const response = await fetch(`http://127.0.0.1:${port}/api/ai/translate/stream/cancel`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ requestId: 'does-not-exist' }),
    })
    assert.equal(response.status, 200)
    const body = (await response.json()) as { ok: boolean; aborted: boolean }
    assert.equal(body.ok, true)
    assert.equal(body.aborted, false)
  } finally {
    close()
  }
})
