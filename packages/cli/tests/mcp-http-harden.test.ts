import { readdirSync, rmSync, type Dirent } from 'node:fs'
import { createServer, type Server } from 'node:http'
import { connect } from 'node:net'
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
      mcpCommand.run(parseArgs(['--http', '8080', '--token', '']), {
        cwd: process.cwd(),
        env: {},
        log: () => {},
        warn: () => {},
      }),
    ).rejects.toThrow(/--token needs a non-empty value/)
  })

  it('fails closed on an empty --host value instead of binding every interface', async () => {
    const { mcpCommand } = await import('../src/commands/mcp')
    const { parseArgs } = await import('../src/args')
    await expect(
      mcpCommand.run(parseArgs(['--http', '8080', '--host', '']), {
        cwd: process.cwd(),
        env: {},
        log: () => {},
        warn: () => {},
      }),
    ).rejects.toThrow(/--host needs a non-empty value/)
  })

  it('treats a programmatically empty host as the loopback default', async () => {
    const { startHttp } = await import('../src/mcp/http')
    const handle = await startHttp({
      cwd: process.cwd(),
      env: {},
      log: () => {},

      port: 0,
      host: '',
    })
    try {
      // the empty string fell back to the loopback default, not every interface
      expect(handle.url.startsWith('http://127.0.0.1:')).toBe(true)
    } finally {
      await handle.close()
    }
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

describe('mcp http loopback Host guard', () => {
  let handle: HttpHandle
  const logs: string[] = []

  beforeAll(async () => {
    // no token, so the loopback Host guard is the thing under test
    handle = await startHttp({
      port: 0,
      host: '127.0.0.1',
      cwd: tempDir(),
      env: { ...process.env, GENOFFICE_AUDIT_LOG: 'off', GENOFFICE_ALLOWED_ROOTS: '' },
      log: (line) => logs.push(line),
      registry: defaultRegistry(),
    })
  })

  afterAll(async () => {
    await handle.close()
  })

  it('lets a normal loopback Host through and still refuses a rebound name', async () => {
    const allowed = await rawRequest(
      handle.port,
      `GET /nope HTTP/1.1\r\nHost: 127.0.0.1:${handle.port}\r\nConnection: close\r\n\r\n`,
    )
    expect(allowed).toMatch(/^HTTP\/1\.1 404 /)

    const rebound = await rawRequest(
      handle.port,
      'GET /nope HTTP/1.1\r\nHost: rebound.example.com\r\nConnection: close\r\n\r\n',
    )
    expect(rebound).toMatch(/^HTTP\/1\.1 403 /)
    expect(rebound).toContain('host not allowed')
  })

  it('refuses a request with no Host header instead of 500ing', async () => {
    // An HTTP/1.0 client may omit Host entirely; the guard must still answer
    // 403, not let a URL parse error escape as a 500.
    const res = await rawRequest(handle.port, 'GET /nope HTTP/1.0\r\n\r\n')
    expect(res).toMatch(/^HTTP\/1\.1 403 /)
    expect(res).toContain('host not allowed')
  })

  it('refuses a Host header the URL parser cannot read instead of 500ing', async () => {
    const res = await rawRequest(
      handle.port,
      'GET /nope HTTP/1.1\r\nHost: a b\r\nConnection: close\r\n\r\n',
    )
    expect(res).toMatch(/^HTTP\/1\.1 403 /)
    expect(res).toContain('host not allowed')
  })

  it('refuses the other authorities new URL() rejects, none of them 500', async () => {
    // A non-numeric port, a bare '::1', an unterminated bracket and an
    // IPv4-mapped form all throw inside new URL(); the guard must absorb them.
    for (const authority of ['127.0.0.1:abc', '::1', '[::1', '::ffff:127.0.0.1']) {
      const res = await rawRequest(
        handle.port,
        `GET /nope HTTP/1.1\r\nHost: ${authority}\r\nConnection: close\r\n\r\n`,
      )
      expect(res, authority).toMatch(/^HTTP\/1\.1 403 /)
      expect(res, authority).toContain('host not allowed')
    }
  })

  it('never reports a rejected Host as a server error', () => {
    expect(logs.some((line) => /invalid url|TypeError/i.test(line))).toBe(false)
  })
})

/** Writes a raw request so a test can control the Host header byte for byte. */
function rawRequest(port: number, request: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    const socket = connect(port, '127.0.0.1', () => socket.write(request))
    socket.on('data', (chunk: Buffer) => chunks.push(chunk))
    socket.on('close', () => resolve(Buffer.concat(chunks).toString('utf8')))
    socket.on('error', reject)
  })
}

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
