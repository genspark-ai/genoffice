import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js'
import { startHttp } from '../src/mcp/http'

const TOKEN = 't-session-cap'

describe('mcp http session cap', () => {
  let handle: Awaited<ReturnType<typeof startHttp>>
  const headers = { authorization: `Bearer ${TOKEN}` }
  const clients: Client[] = []

  beforeAll(async () => {
    handle = await startHttp({
      port: 0,
      host: '127.0.0.1',
      token: TOKEN,
      cwd: mkdtempSync(join(tmpdir(), 'mcp-cap-')),
      env: { ...process.env, GENOFFICE_AUDIT_LOG: 'off', GENOFFICE_ALLOWED_ROOTS: '' },
      log: () => {},
    })
  })

  afterAll(async () => {
    for (const c of clients) await c.close().catch(() => undefined)
    await handle.close()
  })

  const connect = async (): Promise<Client> => {
    const c = new Client({ name: 'cap-test', version: '0' })
    await c.connect(
      new StreamableHTTPClientTransport(new URL(`${handle.url}/mcp`), {
        requestInit: { headers },
      }),
    )
    clients.push(c)
    return c
  }

  it('evicts the oldest session once the cap is reached (no unbounded growth)', async () => {
    const first = await connect()
    // fill to the cap and one past it; each connect is one session
    for (let i = 0; i < 100; i += 1) await connect()
    const health = (await (await fetch(`${handle.url}/health`, { headers })).json()) as {
      sessions: number
    }
    expect(health.sessions).toBe(100)
    // the oldest session was evicted server-side: further calls on it fail
    await expect(first.listTools()).rejects.toThrow()
  })
})
