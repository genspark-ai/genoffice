/**
 * Regression test for POST /api/ai/pi-prompt.
 *
 * The endpoint is the bridge that lets the GenOffice UI drive the embedded pi
 * AgentSession through SSE. Two paths to verify:
 *
 *   1. Bad inputs are rejected with HTTP 400 and a structured body.
 *   2. A valid prompt reaches the pi session — the server streams at least
 *      one event back before the agent fails because no LLM key is configured
 *      on CI. The point is the wiring, not the LLM call.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { request } from 'node:http'

const PORT = 18910 + Math.floor(Math.random() * 100)
const TMP_DATA = mkdtempSync(join(tmpdir(), 'genoffice-pi-prompt-'))
const BUNDLE = join(__dirname, '..', 'dist', 'bundle', 'index.js')

let server: ChildProcessWithoutNullStreams | null = null

interface SseEvent {
  type: string
  requestId?: string
  message?: string
  error?: string
}

function postJson(
  pathname: string,
  body: unknown,
  timeoutMs = 30_000,
): Promise<{ status: number; body: string }> {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(body)
    const req = request(
      {
        hostname: '127.0.0.1',
        port: PORT,
        path: pathname,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Content-Length': Buffer.byteLength(data),
        },
      },
      (res) => {
        let chunks = ''
        res.on('data', (c) => (chunks += c))
        res.on('end', () => resolve({ status: res.statusCode ?? 0, body: chunks }))
        res.on('error', reject)
      },
    )
    req.setTimeout(timeoutMs, () => req.destroy(new Error(`timeout ${pathname}`)))
    req.on('error', reject)
    req.write(data)
    req.end()
  })
}

function getSseEvents(pathname: string, body: unknown, maxMs = 30_000): Promise<SseEvent[]> {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(body)
    const req = request(
      {
        hostname: '127.0.0.1',
        port: PORT,
        path: pathname,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Content-Length': Buffer.byteLength(data),
          Accept: 'text/event-stream',
        },
      },
      (res) => {
        if ((res.statusCode ?? 0) !== 200) {
          let chunks = ''
          res.on('data', (c) => (chunks += c))
          res.on('end', () => reject(new Error(`non-200: ${res.statusCode} ${chunks}`)))
          return
        }
        const events: SseEvent[] = []
        let buffer = ''
        const timer = setTimeout(() => {
          req.destroy()
          resolve(events)
        }, maxMs)
        res.on('data', (chunk) => {
          buffer += chunk.toString('utf-8')
          let idx
          while ((idx = buffer.indexOf('\n\n')) !== -1) {
            const frame = buffer.slice(0, idx)
            buffer = buffer.slice(idx + 2)
            const dataLines = frame
              .split('\n')
              .filter((l) => l.startsWith('data:'))
              .map((l) => l.slice(5).trim())
            for (const line of dataLines) {
              try {
                events.push(JSON.parse(line) as SseEvent)
              } catch {
                /* ignore non-JSON frames */
              }
            }
          }
        })
        res.on('end', () => {
          clearTimeout(timer)
          resolve(events)
        })
        res.on('error', (err) => {
          clearTimeout(timer)
          reject(err)
        })
      },
    )
    req.setTimeout(maxMs + 5_000, () => req.destroy(new Error('timeout')))
    req.on('error', reject)
    req.write(data)
    req.end()
  })
}

beforeAll(async () => {
  server = spawn('node', [BUNDLE], {
    env: { ...process.env, DATA_DIR: TMP_DATA, PORT: String(PORT) },
    stdio: ['ignore', 'pipe', 'pipe'],
  })
  // Wait for the banner to appear.
  await new Promise<void>((resolve) => {
    const onChunk = (chunk: Buffer) => {
      const text = chunk.toString('utf-8')
      if (text.includes('GenOffice Web Server')) {
        server?.stdout.off('data', onChunk)
        resolve()
      }
    }
    server!.stdout.on('data', onChunk)
  })
}, 30_000)

afterAll(() => {
  if (server) {
    server.kill('SIGTERM')
    server = null
  }
  try { rmSync(TMP_DATA, { recursive: true, force: true }) } catch { /* ignore */ }
})

describe('POST /api/ai/pi-prompt', () => {
  it('returns 400 with a structured body for empty text', async () => {
    const res = await postJson('/api/ai/pi-prompt', { text: '' })
    expect(res.status).toBe(400)
    const parsed = JSON.parse(res.body) as { ok: boolean; error: string }
    expect(parsed.ok).toBe(false)
    expect(parsed.error).toMatch(/empty text/)
  })

  it('returns 400 with a structured body for invalid JSON', async () => {
    const res = await new Promise<{ status: number; body: string }>((resolve, reject) => {
      const req = request(
        {
          hostname: '127.0.0.1',
          port: PORT,
          path: '/api/ai/pi-prompt',
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
        },
        (response) => {
          let chunks = ''
          response.on('data', (c) => (chunks += c))
          response.on('end', () =>
            resolve({ status: response.statusCode ?? 0, body: chunks }),
          )
        },
      )
      req.on('error', reject)
      req.write('not json')
      req.end()
    })
    expect(res.status).toBe(400)
    expect(JSON.parse(res.body).error).toMatch(/invalid JSON/i)
  })

  it('streams a `start` event and at least one agent-side event before terminating', async () => {
    const events = await getSseEvents(
      '/api/ai/pi-prompt',
      { text: 'Use the web_search skill to find out who created the Linux kernel' },
      25_000,
    )
    // We expect at least the `start` event from the server. With no LLM key
    // configured, the agent errors out fast, so we also expect an `error`
    // event whose message mentions the missing API key — the *bridge*
    // worked, the LLM call did not.
    const types = events.map((e) => e.type)
    expect(types).toContain('start')
    const errorEvent = events.find((e) => e.type === 'error')
    expect(errorEvent, `events: ${JSON.stringify(types)}`).toBeTruthy()
    expect(String(errorEvent?.message ?? '')).toMatch(/API key|login|provider/i)
  })
})
