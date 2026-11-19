import { mkdtempSync, readFileSync, statSync } from 'node:fs'
import { connect } from 'node:net'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { ControlReply, ControlRequest } from '@genoffice/cli/control-protocol'
import { controlHandler, type ControlHost } from '../src/main/control-handlers'
import { parseEnvelope, startControlServer, type ControlServer } from '../src/main/control-server'

const servers: ControlServer[] = []
afterEach(() => {
  while (servers.length) servers.pop()!.close()
})

function roundTrip(endpoint: string, line: string): Promise<{ reply: string; closed: boolean }> {
  return new Promise((resolve, reject) => {
    const socket = connect(endpoint)
    let reply = ''
    socket.setEncoding('utf8')
    socket.on('error', reject)
    socket.on('connect', () => socket.write(line))
    socket.on('data', (chunk: string) => (reply += chunk))
    socket.on('close', () => resolve({ reply, closed: true }))
  })
}

describe('control server', () => {
  it('publishes a token-protected endpoint and answers one request per connection', async () => {
    const userData = mkdtempSync(join(tmpdir(), 'genoffice-control-'))
    const seen: ControlRequest[] = []
    const server = await startControlServer(userData, async (req) => {
      seen.push(req)
      return { ok: true, result: { echoed: req.path } }
    })
    servers.push(server)
    const published = JSON.parse(readFileSync(join(userData, 'control.json'), 'utf8'))
    expect(published).toMatchObject({
      protocol: 1,
      pid: process.pid,
      endpoint: server.endpoint.endpoint,
    })
    expect(published.token).toHaveLength(48)
    if (process.platform !== 'win32') {
      expect(statSync(join(userData, 'control.json')).mode & 0o777).toBe(0o600)
    }

    const good = await roundTrip(
      server.endpoint.endpoint,
      JSON.stringify({ token: published.token, request: { cmd: 'selection', path: '/a.docx' } }) +
        '\n',
    )
    expect(JSON.parse(good.reply)).toEqual({ ok: true, result: { echoed: '/a.docx' } })
    expect(seen).toEqual([{ cmd: 'selection', path: '/a.docx' }])

    const wrongToken = await roundTrip(
      server.endpoint.endpoint,
      JSON.stringify({ token: 'nope', request: { cmd: 'selection', path: '/a.docx' } }) + '\n',
    )
    expect(wrongToken.reply).toBe('')
    expect(seen).toHaveLength(1)

    const garbage = await roundTrip(server.endpoint.endpoint, 'not json\n')
    expect(garbage.reply).toBe('')
    expect(seen).toHaveLength(1)
  })

  it('rejects envelopes that are not a known command', () => {
    expect(parseEnvelope('{"token":"t","request":{"cmd":"rm","path":"/x"}}')).toBeNull()
    expect(parseEnvelope('{"token":"t","request":{"cmd":"open"}}')).toBeNull()
    expect(parseEnvelope('{"request":{"cmd":"open","path":"/x"}}')).toBeNull()
    expect(parseEnvelope('{"token":"t","request":{"cmd":"open","path":"/x"}}')).toEqual({
      token: 't',
      request: { cmd: 'open', path: '/x' },
    })
  })
})

describe('control handler', () => {
  const fakeWebContents = (replies: unknown[]) => {
    const executeJavaScript = vi.fn(async () => replies.shift())
    return {
      wc: { isDestroyed: () => false, isLoading: () => false, executeJavaScript } as never,
      executeJavaScript,
    }
  }
  const host = (overrides: Partial<ControlHost> & { wc?: never }): ControlHost => ({
    reveal: vi.fn(),
    openDocument: vi.fn(() => true),
    activateTab: vi.fn(),
    findTab: () => undefined,
    readyTimeoutMs: 1_000,
    ...overrides,
  })

  it('activates the tab that already shows the file, waits for the renderer and relays the target', async () => {
    const { wc, executeJavaScript } = fakeWebContents([
      { status: 'not_ready' },
      { status: 'ok', result: { slide: 2 } },
    ])
    const h = host({ findTab: () => ({ id: 't1', kind: 'slides', webContents: wc }) })
    const reply = await controlHandler(h)({
      cmd: 'open',
      path: __filename,
      target: { kind: 'slide', slide: 2 },
    })
    expect(reply).toEqual({ ok: true, result: { slide: 2 } })
    expect(h.reveal).toHaveBeenCalled()
    expect(h.activateTab).toHaveBeenCalledWith('t1')
    expect(h.openDocument).not.toHaveBeenCalled()
    expect(executeJavaScript).toHaveBeenCalledTimes(2)
    expect(String(executeJavaScript.mock.calls[0]![0])).toContain('"cmd":"goto"')
  })

  it('opens a file no tab shows yet, then finds the new tab', async () => {
    const { wc } = fakeWebContents([{ status: 'ok', result: { block: 1 } }])
    let opened = false
    const h = host({
      openDocument: vi.fn(() => {
        opened = true
        return true
      }),
      findTab: () => (opened ? { id: 't9', kind: 'docs', webContents: wc } : undefined),
    })
    const reply = await controlHandler(h)({
      cmd: 'open',
      path: __filename,
      target: { kind: 'block', block: 1 },
    })
    expect(reply).toEqual({ ok: true, result: { block: 1 } })
    expect(h.openDocument).toHaveBeenCalledWith(__filename)
    expect(h.activateTab).not.toHaveBeenCalled()
  })

  it('refuses a target kind the tab cannot take and reports a missing tab', async () => {
    const { wc } = fakeWebContents([])
    const mismatch = await controlHandler(
      host({ findTab: () => ({ id: 't2', kind: 'docs', webContents: wc }) }),
    )({
      cmd: 'open',
      path: __filename,
      target: { kind: 'page', page: 1 },
    })
    expect(mismatch).toMatchObject({
      ok: false,
      error: { reason: 'unsupported', detail: { supported: ['block'] } },
    })

    const notOpen: ControlReply = await controlHandler(host({}))({
      cmd: 'selection',
      path: __filename,
    })
    expect(notOpen).toMatchObject({ ok: false, error: { reason: 'file_not_open_in_gui' } })

    const missing = await controlHandler(host({}))({
      cmd: 'open',
      path: join(tmpdir(), 'nope.docx'),
    })
    expect(missing).toMatchObject({ ok: false, error: { reason: 'file_not_found' } })
  })

  it('passes renderer errors through and gives up when the document never loads', async () => {
    const { wc } = fakeWebContents([
      {
        status: 'error',
        error: { reason: 'out_of_range', message: 'slide 9', detail: { valid_range: '0-3' } },
      },
    ])
    const h = host({ findTab: () => ({ id: 't1', kind: 'slides', webContents: wc }) })
    expect(await controlHandler(h)({ cmd: 'selection', path: __filename })).toEqual({
      ok: false,
      error: { reason: 'out_of_range', message: 'slide 9', detail: { valid_range: '0-3' } },
    })

    const never = fakeWebContents([])
    const slow = host({
      findTab: () => ({ id: 't3', kind: 'pdf', webContents: never.wc }),
      readyTimeoutMs: 300,
    })
    expect(await controlHandler(slow)({ cmd: 'selection', path: __filename })).toMatchObject({
      ok: false,
      error: { reason: 'app_unavailable' },
    })
  })
})
