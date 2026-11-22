import http from 'node:http'
import net from 'node:net'
import tls from 'node:tls'
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest'
import {
  egressDecision,
  externalOpenNeedsConfirm,
  isLoopbackHost,
} from '../src/main/local-only/allowlist'
import {
  BLACK_HOLE_PROXY,
  OFFLINE_ERROR_CODE,
  OfflineBlockedError,
  connectTarget,
  installElectronEgressGuard,
  installNodeEgressGuard,
  loopbackListenArgs,
  pinChildProxyEnv,
  type EgressEvent,
  type ElectronEgressDeps,
} from '../src/main/local-only/egress-guard'

/**
 * Local-only fork egress guard (src/main/local-only/). Public destinations
 * must fail before any DNS lookup or connection, so these tests never touch
 * the network: a blocked attempt errors locally, an allowed one hits a
 * loopback server started here.
 */

describe('allowlist', () => {
  it('accepts loopback hosts only', () => {
    for (const host of ['localhost', 'LOCALHOST', '127.0.0.1', '127.8.9.10', '::1', '[::1]']) {
      expect(isLoopbackHost(host), host).toBe(true)
    }
    for (const host of [
      'example.com',
      '192.168.1.10',
      '10.0.0.1',
      '0.0.0.0',
      'intranet',
      'localhost.evil.com',
    ]) {
      expect(isLoopbackHost(host), host).toBe(false)
    }
  })

  it('blocks network schemes to public hosts and allows local schemes', () => {
    expect(egressDecision('https://api.openai.com/v1/chat')).toEqual({
      allowed: false,
      host: 'api.openai.com',
    })
    expect(egressDecision('wss://example.com/socket').allowed).toBe(false)
    expect(egressDecision('http://localhost:5173/index.html').allowed).toBe(true)
    expect(egressDecision('ws://127.0.0.1:8080').allowed).toBe(true)
    expect(egressDecision('file:///Users/me/report.docx').allowed).toBe(true)
    expect(egressDecision('data:text/plain,hi').allowed).toBe(true)
    expect(egressDecision('genoffice-docx-media://x/y.png').allowed).toBe(true)
  })

  it('asks before opening public web links only', () => {
    expect(externalOpenNeedsConfirm('https://github.com/genspark-ai/genoffice')).toBe(true)
    expect(externalOpenNeedsConfirm('http://localhost:3000')).toBe(false)
    expect(externalOpenNeedsConfirm('mailto:someone@example.com')).toBe(false)
    expect(externalOpenNeedsConfirm('file:///tmp')).toBe(false)
  })
})

describe('argument normalization', () => {
  it('reads every Socket.connect overload', () => {
    expect(connectTarget([443, 'example.com'])).toEqual({ host: 'example.com' })
    expect(connectTarget(['443', 'example.com'])).toEqual({ host: 'example.com' })
    expect(connectTarget([8080])).toEqual({ host: 'localhost' })
    expect(connectTarget([{ port: 1, host: '10.0.0.2' }])).toEqual({ host: '10.0.0.2' })
    expect(connectTarget([[{ port: 1, host: 'a.b' }, null]])).toEqual({ host: 'a.b' })
    expect(connectTarget(['/tmp/app.sock'])).toEqual({ path: '/tmp/app.sock' })
    expect(connectTarget([{ path: '\\\\.\\pipe\\x' }])).toEqual({ path: '\\\\.\\pipe\\x' })
  })

  it('binds listen() to loopback and refuses public hosts', () => {
    const cb = () => {}
    expect(loopbackListenArgs([3000])).toEqual([3000, '127.0.0.1'])
    expect(loopbackListenArgs([3000, cb])).toEqual([3000, '127.0.0.1', cb])
    expect(loopbackListenArgs([cb])).toEqual([0, '127.0.0.1', cb])
    expect(loopbackListenArgs([])).toEqual([0, '127.0.0.1'])
    expect(loopbackListenArgs([3000, '::1'])).toEqual([3000, '::1'])
    expect(loopbackListenArgs([{ port: 1 }])).toEqual([{ port: 1, host: '127.0.0.1' }])
    expect(loopbackListenArgs(['/tmp/app.sock'])).toEqual(['/tmp/app.sock'])
    expect(() => loopbackListenArgs([3000, '0.0.0.0'])).toThrow(OfflineBlockedError)
    expect(() => loopbackListenArgs([{ port: 1, host: '192.168.0.2' }])).toThrow(
      OfflineBlockedError,
    )
  })

  it('points child-process proxy variables at the black hole', () => {
    const env: NodeJS.ProcessEnv = { HTTPS_PROXY: 'http://corp-proxy:8080' }
    pinChildProxyEnv(env)
    expect(env.HTTPS_PROXY).toBe(BLACK_HOLE_PROXY)
    expect(env.https_proxy).toBe(BLACK_HOLE_PROXY)
    expect(env.ALL_PROXY).toBe(BLACK_HOLE_PROXY)
    expect(env.NO_PROXY).toContain('127.0.0.1')
  })
})

describe('Node guard', () => {
  const events: EgressEvent[] = []
  let uninstall: () => void
  let server: http.Server
  let port = 0

  beforeAll(async () => {
    uninstall = installNodeEgressGuard((e) => events.push(e))
    server = http.createServer((_req, res) => res.end('local ok'))
    // no host given: the guard binds it to loopback
    await new Promise<void>((resolve) => server.listen(0, resolve))
    port = (server.address() as net.AddressInfo).port
  })

  afterAll(async () => {
    await new Promise<void>((resolve) => server.close(() => resolve()))
    uninstall()
  })

  it('binds an unspecified listen host to loopback', () => {
    expect((server.address() as net.AddressInfo).address).toBe('127.0.0.1')
  })

  it('lets loopback fetch through', async () => {
    const res = await fetch(`http://127.0.0.1:${port}/`)
    expect(await res.text()).toBe('local ok')
  })

  it('rejects fetch to a public host', async () => {
    await expect(fetch('https://api.anthropic.com/v1/messages')).rejects.toMatchObject({
      code: OFFLINE_ERROR_CODE,
      host: 'api.anthropic.com',
    })
    expect(events.some((e) => e.kind === 'fetch' && e.host === 'api.anthropic.com')).toBe(true)
  })

  it('fails http.get to a public host like a refused connection', async () => {
    const error = await new Promise<NodeJS.ErrnoException>((resolve) => {
      http
        .get('http://example.com/', () => resolve(new Error('connected') as NodeJS.ErrnoException))
        .on('error', resolve)
    })
    expect(error.code).toBe(OFFLINE_ERROR_CODE)
  })

  it('fails raw TCP and TLS sockets to public hosts', async () => {
    const tcp = await new Promise<NodeJS.ErrnoException>((resolve) => {
      net.connect(80, '93.184.215.14').on('error', resolve)
    })
    expect(tcp.code).toBe(OFFLINE_ERROR_CODE)
    const secure = await new Promise<NodeJS.ErrnoException>((resolve) => {
      tls.connect(443, 'github.com').on('error', resolve)
    })
    expect(secure.code).toBe(OFFLINE_ERROR_CODE)
  })

  it('refuses to listen on a public interface', () => {
    const s = net.createServer()
    expect(() => s.listen(0, '0.0.0.0')).toThrow(OfflineBlockedError)
  })
})

describe('Electron guard', () => {
  function fakes(choice: number) {
    const handlers: Array<(s: FakeSession) => void> = []
    type Listener = (
      details: { url: string; resourceType?: string },
      cb: (r: { cancel?: boolean }) => void,
    ) => void
    class FakeSession {
      listener: Listener | null = null
      setProxy = vi.fn(async () => {})
      webRequest = { onBeforeRequest: (l: Listener) => void (this.listener = l) }
    }
    const defaultSession = new FakeSession()
    const opened: string[] = []
    const showMessageBox = vi.fn(async () => ({ response: choice }))
    const deps = {
      app: {
        on: (_event: string, l: (s: FakeSession) => void) => handlers.push(l),
        whenReady: () => Promise.resolve(),
      },
      session: { defaultSession },
      shell: { openExternal: async (url: string) => void opened.push(url) },
      dialog: { showMessageBox },
    }
    const events: EgressEvent[] = []
    installElectronEgressGuard(deps as unknown as ElectronEgressDeps, (e) => events.push(e))
    return { deps, defaultSession, FakeSession, handlers, opened, showMessageBox, events }
  }

  function decide(session: { listener: unknown }, url: string): boolean | undefined {
    let cancel: boolean | undefined
    ;(session.listener as (d: { url: string }, cb: (r: { cancel?: boolean }) => void) => void)(
      { url },
      (r) => (cancel = r.cancel),
    )
    return cancel
  }

  it('guards the default session and every new one', async () => {
    const { defaultSession, FakeSession, handlers } = fakes(0)
    await Promise.resolve()
    expect(defaultSession.setProxy).toHaveBeenCalledWith(
      expect.objectContaining({ proxyRules: BLACK_HOLE_PROXY }),
    )
    expect(decide(defaultSession, 'https://www.genspark.ai/api')).toBe(true)
    expect(decide(defaultSession, 'http://localhost:5173/')).toBeUndefined()
    expect(decide(defaultSession, 'file:///x.docx')).toBeUndefined()

    const partition = new FakeSession()
    for (const h of handlers) h(partition)
    expect(decide(partition, 'wss://example.com/')).toBe(true)
  })

  it('opens public links only after the user confirms', async () => {
    const declined = fakes(0)
    await declined.deps.shell.openExternal('https://genoffice.ai/join')
    expect(declined.opened).toEqual([])
    expect(declined.events).toEqual([{ kind: 'open-external', host: 'genoffice.ai' }])

    const accepted = fakes(1)
    await accepted.deps.shell.openExternal('https://genoffice.ai/join')
    expect(accepted.opened).toEqual(['https://genoffice.ai/join'])

    await accepted.deps.shell.openExternal('http://127.0.0.1:3000/')
    expect(accepted.showMessageBox).toHaveBeenCalledTimes(1)
  })
})
