import { readFile, rm, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import http from 'node:http'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { applyMcpSettings, configureMcpRuntime, stopMcp } from '../../src/main/mcp/app-mcp'
import type { McpSettings } from '../../src/main/mcp/app-mcp'

/**
 * Regression (runbook §2/§10): flipping "background generation" while the
 * server is running must restart it, because the flip changes the exposed
 * tool list — create_docx appears/disappears without touching the port.
 * The old code compared next.background against the already-overwritten
 * currentSettings, so the comparison was always false and the server never
 * rebuilt until a port change or app restart.
 */

let dir: string
let logPath: string
let port: number

beforeEach(async () => {
  dir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-apply-'))
  logPath = join(dir, 'mcp-log.txt')
  port = await freePort()
  configureMcpRuntime({
    version: 'test',
    defaultSaveDir: () => dir,
    openPath: () => false,
    logFilePath: logPath,
  })
})

afterEach(async () => {
  await stopMcp()
  await rm(dir, { recursive: true, force: true })
})

async function freePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const server = http.createServer()
    server.on('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const address = server.address()
      const p = typeof address === 'object' && address ? address.port : 0
      server.close(() => resolve(p))
    })
  })
}

/** tools/list over real HTTP, fresh session per call */
async function listTools(): Promise<string[]> {
  const handshake = await post('/mcp', {
    jsonrpc: '2.0',
    id: 1,
    method: 'initialize',
    params: {
      protocolVersion: '2024-11-05',
      capabilities: {},
      clientInfo: { name: 'apply-test', version: '1' },
    },
  })
  const sid = handshake.headers['mcp-session-id'] ?? null
  await post('/mcp', { jsonrpc: '2.0', method: 'notifications/initialized' }, sid)
  const list = await post('/mcp', { jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} }, sid)
  const body = parseBody(list.text) as { result?: { tools?: Array<{ name: string }> } }
  return body.result?.tools?.map((t) => t.name) ?? []
}

/** the transport may answer as JSON or as an SSE event stream */
function parseBody(text: string): unknown {
  const trimmed = text.trim()
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) return JSON.parse(trimmed)
  const events: unknown[] = []
  for (const line of trimmed.split('\n')) {
    if (line.startsWith('data: ')) {
      try {
        events.push(JSON.parse(line.slice(6)))
      } catch {
        /* comment line */
      }
    }
  }
  return events[0]
}

function post(
  path: string,
  payload: unknown,
  sid: string | null = null,
): Promise<{ status: number; headers: http.IncomingHttpHeaders; text: string }> {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(payload)
    const headers: Record<string, string | number> = {
      'Content-Type': 'application/json',
      Accept: 'application/json, text/event-stream',
      'Content-Length': Buffer.byteLength(data),
    }
    if (sid) headers['Mcp-Session-Id'] = sid
    const req = http.request(
      // agent:false: the restart under test destroys old sockets, so a pooled
      // keep-alive connection would ECONNRESET mid-suite
      { hostname: '127.0.0.1', port, path, method: 'POST', headers, agent: false },
      (res) => {
        let text = ''
        res.on('data', (c) => (text += c))
        res.on('end', () => resolve({ status: res.statusCode ?? 0, headers: res.headers, text }))
      },
    )
    req.on('error', reject)
    req.write(data)
    req.end()
  })
}

function settings(patch: Partial<McpSettings>): McpSettings {
  return { enabled: true, port, background: false, logging: true, ...patch }
}

describe('applyMcpSettings restarts on a background flip', () => {
  it('exposes and hides create_docx live, without a port change', async () => {
    // baseline: background off (the default) — no headless tool
    await applyMcpSettings(settings({ background: false }))
    expect(await listTools()).not.toContain('create_docx')

    // flip ON: server must rebuild and expose create_docx on the same port
    await applyMcpSettings(settings({ background: true }))
    expect(await listTools()).toContain('create_docx')

    // flip OFF again: it disappears without restarting the app
    await applyMcpSettings(settings({ background: false }))
    expect(await listTools()).not.toContain('create_docx')

    // the restarts went through the logger (stop + a second listen)
    const log = await readFile(logPath, 'utf8')
    expect(log).toContain('[mcp] stopped')
    expect(log.match(/\[mcp\] listening on/g)?.length).toBe(3)
  })

  it('reports the new background/logging state from mcpStatus', async () => {
    const status = await applyMcpSettings(settings({ background: true, logging: true }))
    expect(status.running).toBe(true)
    expect(status.enabled).toBe(true)
    expect(status.background).toBe(true)
    expect(status.logging).toBe(true)
    expect(status.url).toContain(`:${port}`)
  })
})
