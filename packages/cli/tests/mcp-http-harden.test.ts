import { readdirSync, rmSync, type Dirent } from 'node:fs'
import { createServer, type Server } from 'node:http'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { defaultRegistry } from '../src/cli'
import { fetchToFile } from '../src/mcp/files'
import { startHttp, type HttpHandle } from '../src/mcp/http'
import { tempDir } from './helpers'

const TOKEN = 'harden-token'

describe('mcp http hardening', () => {
  let handle: HttpHandle
  const headers = { authorization: `Bearer ${TOKEN}` }

  beforeAll(async () => {
    handle = await startHttp({
      port: 0,
      host: '127.0.0.1',
      token: TOKEN,
      cwd: tempDir(),
      env: { ...process.env, GENOFFICE_AUDIT_LOG: 'off', GENOFFICE_ALLOWED_ROOTS: '' },
      log: () => {},
      registry: defaultRegistry(),
    })
  })

  afterAll(async () => {
    await handle.close()
  })

  it('fails closed on an empty --token value instead of serving unauthenticated', async () => {
    const { mcpCommand } = await import('../src/commands/mcp')
    const { parseArgs } = await import('../src/args')
    await expect(
      mcpCommand.run(
        parseArgs(['mcp', 'serve', '--http', '8080', '--token', '']),
        {
          cwd: process.cwd(),
          env: {},
          log: () => {},
        } as Parameters<typeof mcpCommand.run>[1],
      ),
    ).rejects.toThrow(/--token needs a non-empty value/)
  })

  it('ignores poisoned forwarded host/proto by default', async () => {
    const res = await fetch(`${handle.url}/files/report.txt`, {
      method: 'PUT',
      headers: {
        ...headers,
        'x-forwarded-host': 'evil.example.com',
        'x-forwarded-proto': 'https',
      },
      body: 'hello',
    })
    expect(res.status).toBe(201)
    const body = (await res.json()) as { url: string }
    expect(body.url.startsWith(handle.url)).toBe(true)
    expect(body.url).not.toContain('evil.example.com')
  })

  it('honors forwarded headers when the proxy opt-in env is set', async () => {
    const trusted = await startHttp({
      port: 0,
      host: '127.0.0.1',
      token: TOKEN,
      cwd: tempDir(),
      env: {
        ...process.env,
        GENOFFICE_AUDIT_LOG: 'off',
        GENOFFICE_ALLOWED_ROOTS: '',
        GENOFFICE_TRUST_PROXY_HEADERS: '1',
      },
      log: () => {},
      registry: defaultRegistry(),
    })
    try {
      const res = await fetch(`${trusted.url}/files/report.txt`, {
        method: 'PUT',
        headers: {
          ...headers,
          'x-forwarded-host': 'proxy.example.com',
          'x-forwarded-proto': 'https',
        },
        body: 'hello',
      })
      expect(res.status).toBe(201)
      const body = (await res.json()) as { url: string }
      expect(body.url).toContain('proxy.example.com')
    } finally {
      await trusted.close()
    }
  })

  it('falls back to a sanitized name for a %ZZ upload instead of 500', async () => {
    const res = await fetch(`${handle.url}/files/%ZZ`, {
      method: 'PUT',
      headers,
      body: 'hello',
    })
    expect(res.status).toBe(201)
    const body = (await res.json()) as { name: string; size: number }
    expect(body.size).toBe(5)
    expect(body.name).not.toContain('%')
    expect(body.name.length).toBeGreaterThan(0)
  })

  it('falls back instead of throwing for a %ZZ remote name', async () => {
    const server: Server = createServer((req, res) => {
      if (req.url === '/redirect-bad-name') {
        res.writeHead(200, {
          'content-type': 'text/plain',
          'content-disposition': `attachment; filename*=UTF-8''%ZZ`,
        })
        res.end('hello')
        return
      }
      res.writeHead(200, { 'content-type': 'text/plain' })
      res.end('hello')
    })
    await new Promise<void>((r) => server.listen(0, '127.0.0.1', r))
    const port = (server.address() as { port: number }).port
    try {
      const dir = tempDir()
      const fromHeader = await fetchToFile(`http://127.0.0.1:${port}/redirect-bad-name`, dir)
      // '_ZZ' has no extension, so the text/plain content type contributes '.txt'
      expect(basename(fromHeader)).toBe('_ZZ.txt')
      const fromPath = await fetchToFile(`http://127.0.0.1:${port}/%ZZ`, dir)
      expect(basename(fromPath)).toBe('_ZZ.txt')
    } finally {
      server.close()
    }
  })

  it('keeps redirects http(s)-only', async () => {
    const server: Server = createServer((req, res) => {
      res.writeHead(302, { location: 'ftp://127.0.0.1/evil.bin' })
      res.end()
    })
    await new Promise<void>((r) => server.listen(0, '127.0.0.1', r))
    const port = (server.address() as { port: number }).port
    try {
      await expect(fetchToFile(`http://127.0.0.1:${port}/start`, tempDir())).rejects.toThrow(
        'only http(s)',
      )
    } finally {
      server.close()
    }
  })

  it('returns 404 (not 500) for a missing or raced download', async () => {
    const missing = await fetch(`${handle.url}/files/0123456789abcdef/gone.bin`, { headers })
    expect(missing.status).toBe(404)

    // Simulate the race: upload, delete the stored bytes behind the
    // server's back, then the download must still be a 404, not a 500.
    const up = await fetch(`${handle.url}/files/race.bin`, {
      method: 'PUT',
      headers,
      body: 'race-bytes',
    })
    expect(up.status).toBe(201)
    const { url } = (await up.json()) as { url: string }
    const first = await fetch(url, { headers })
    expect(first.status).toBe(200)
    await first.arrayBuffer()
    removeStoredCopy('race.bin')
    const second = await fetch(url, { headers })
    expect(second.status).toBe(404)
  })
})

/** Delete this server's stored copy found under the tmp FileStore roots. */
function removeStoredCopy(name: string): void {
  const prefix = `genoffice-mcp-http-${process.pid}-`
  for (const entry of readdirSync(tmpdir())) {
    if (!entry.startsWith(prefix)) continue
    removeMatching(join(tmpdir(), entry), name)
  }
}

function removeMatching(dir: string, name: string): void {
  let entries: Dirent[]
  try {
    entries = readdirSync(dir, { withFileTypes: true })
  } catch {
    return
  }
  for (const entry of entries) {
    const full = join(dir, entry.name)
    try {
      if (entry.isDirectory()) removeMatching(full, name)
      else if (entry.name === name) rmSync(full, { force: true })
    } catch {
      // best effort cleanup for the race simulation
    }
  }
}
