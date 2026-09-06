/// Integration tests for the HTTP/SSE IPC bridge — real loopback servers, real
/// fetch requests, real SSE streams (no transport mocks), so the dual-protocol
/// contract is exercised exactly the way the web version drives it.

import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  attachIpcMain,
  createBridgeServer,
  installHttpIpcBridge,
  IPC_NO_HANDLER,
  IpcHandlerRegistry,
  WEB_UNSUPPORTED,
  type BridgeServer,
  type IpcMainLike,
} from '../src/index'

/** Minimal electron-shaped ipcMain double: same duplicate/once semantics. */
class FakeIpcMain implements IpcMainLike {
  readonly handles = new Map<string, (event: unknown, ...args: unknown[]) => unknown>()
  private readonly listeners = new Map<
    string,
    Array<(event: unknown, ...args: unknown[]) => void>
  >()

  handle(channel: string, handler: (event: unknown, ...args: unknown[]) => unknown): void {
    if (this.handles.has(channel)) {
      throw new Error(`Attempted to register a second handler for '${channel}'`)
    }
    this.handles.set(channel, handler)
  }

  removeHandler(channel: string): void {
    this.handles.delete(channel)
  }

  on(channel: string, listener: (event: unknown, ...args: unknown[]) => void): void {
    this.listenersOf(channel).push(listener)
  }

  once(channel: string, listener: (event: unknown, ...args: unknown[]) => void): void {
    const wrapped = (event: unknown, ...args: unknown[]) => {
      this.removeListener(channel, wrapped)
      listener(event, ...args)
    }
    this.listenersOf(channel).push(wrapped)
  }

  removeListener(channel: string, listener: unknown): void {
    const set = this.listeners.get(channel)
    const index = set?.indexOf(listener as (event: unknown, ...args: unknown[]) => void) ?? -1
    if (index >= 0) set?.splice(index, 1)
  }

  removeAllListeners(channel: string): void {
    this.listeners.delete(channel)
  }

  private listenersOf(channel: string) {
    let set = this.listeners.get(channel)
    if (!set) {
      set = []
      this.listeners.set(channel, set)
    }
    return set
  }
}

function invokeUrl(port: number, channel: string): string {
  return `http://127.0.0.1:${port}/api/ipc/${encodeURIComponent(channel)}`
}

/** Opens the SSE push stream and collects `count` data frames. */
function collectSseFrames(
  port: number,
  session: string,
  count: number,
): Promise<Array<{ channel: string; args: unknown[] }>> {
  return (async () => {
    const response = await fetch(`http://127.0.0.1:${port}/api/ipc/events?session=${session}`, {
      signal: AbortSignal.timeout(5_000),
    })
    expect(response.status).toBe(200)
    const reader = response.body!.getReader()
    const decoder = new TextDecoder()
    let buffered = ''
    const frames: Array<{ channel: string; args: unknown[] }> = []
    while (frames.length < count) {
      const { done, value } = await reader.read()
      if (done) break
      buffered += decoder.decode(value, { stream: true })
      let boundary = buffered.indexOf('\n\n')
      while (boundary !== -1 && frames.length < count) {
        const raw = buffered.slice(0, boundary)
        buffered = buffered.slice(boundary + 2)
        if (raw.startsWith('data: ')) frames.push(JSON.parse(raw.slice('data: '.length)))
        boundary = buffered.indexOf('\n\n')
      }
    }
    return frames
  })()
}

describe('bridge server over real HTTP', () => {
  let registry: IpcHandlerRegistry
  let server: BridgeServer

  beforeAll(async () => {
    registry = new IpcHandlerRegistry()
    registry.registerHandle('test:echo', (_event, value: unknown) => value)
    registry.registerHandle('test:bin', (_event, bytes: unknown) => {
      const view = bytes as Uint8Array
      return {
        size: view.byteLength,
        copy: new Uint8Array(view),
        raw: new ArrayBuffer(4),
      }
    })
    registry.registerHandle('test:boom', () => {
      throw new Error('kaboom')
    })
    registry.registerHandle('test:push', (event, ...values: unknown[]) => {
      event.sender.send('push:chunk', values[0], { n: 1 })
      event.sender.send('push:done', true)
      return 'pushed'
    })
    registry.registerHandle('test:native-touch', (event) => {
      // handlers that poke non-send webContents members get a structured throw
      return (event.sender as unknown as { loadURL: unknown }).loadURL
    })
    server = await createBridgeServer({
      registry,
      port: 0,
      nativeOnlyChannels: [/^dialog:/, 'test:nope'],
    })
  })

  afterAll(async () => {
    await server.close()
  })

  it('returns handler results as JSON with ok:true', async () => {
    const response = await fetch(invokeUrl(server.port, 'test:echo'), {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ args: [{ hello: 'web', n: 42 }] }),
      signal: AbortSignal.timeout(5_000),
    })
    expect(response.status).toBe(200)
    expect(await response.json()).toEqual({ ok: true, result: { hello: 'web', n: 42 } })
  })

  it('survives binary payloads in both directions (ArrayBuffer/typed arrays)', async () => {
    const bytes = Uint8Array.from([0, 1, 2, 250, 251, 255])
    const response = await fetch(invokeUrl(server.port, 'test:bin'), {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({
        args: [{ __ipcBytes: 'u8', b64: Buffer.from(bytes).toString('base64') }],
      }),
      signal: AbortSignal.timeout(5_000),
    })
    const body = (await response.json()) as {
      ok: boolean
      result: {
        size: number
        copy: { __ipcBytes: string; b64: string }
        raw: { __ipcBytes: string; b64: string }
      }
    }
    expect(body.ok).toBe(true)
    expect(body.result.size).toBe(bytes.byteLength)
    expect(new Uint8Array(Buffer.from(body.result.copy.b64, 'base64'))).toEqual(bytes)
    expect(body.result.raw.__ipcBytes).toBe('ab')
  })

  it('maps handler throws to non-2xx with the original message', async () => {
    const response = await fetch(invokeUrl(server.port, 'test:boom'), {
      method: 'POST',
      body: '{}',
      signal: AbortSignal.timeout(5_000),
    })
    expect(response.status).toBe(500)
    expect(await response.json()).toEqual({ error: { message: 'kaboom' } })
  })

  it('returns a structured 404 for channels without handlers', async () => {
    const response = await fetch(invokeUrl(server.port, 'test:missing'), {
      method: 'POST',
      body: '{}',
      signal: AbortSignal.timeout(5_000),
    })
    expect(response.status).toBe(404)
    const body = (await response.json()) as { error: { code: string; message: string } }
    expect(body.error.code).toBe(IPC_NO_HANDLER)
    expect(body.error.message).toContain('test:missing')
  })

  it('answers native-only channels with a structured WEB_UNSUPPORTED error', async () => {
    for (const channel of ['dialog:open', 'test:nope']) {
      const response = await fetch(invokeUrl(server.port, channel), {
        method: 'POST',
        body: '{}',
        signal: AbortSignal.timeout(5_000),
      })
      expect(response.status).toBe(400)
      const body = (await response.json()) as { error: { code: string; message: string } }
      expect(body.error.code).toBe(WEB_UNSUPPORTED)
      expect(body.error.message).toContain('desktop-only')
    }
  })

  it('delivers push frames to the session SSE stream in order', async () => {
    const session = 'sess-push-1'
    const framesPromise = collectSseFrames(server.port, session, 2)
    const response = await fetch(invokeUrl(server.port, 'test:push'), {
      method: 'POST',
      headers: { 'x-ipc-session': session },
      body: JSON.stringify({ args: ['hello'] }),
      signal: AbortSignal.timeout(5_000),
    })
    expect(((await response.json()) as { result: string }).result).toBe('pushed')
    const frames = await framesPromise
    expect(frames).toEqual([
      { channel: 'push:chunk', args: ['hello', { n: 1 }] },
      { channel: 'push:done', args: [true] },
    ])
  })

  it('exposes a health endpoint with channel/session counts', async () => {
    const response = await fetch(`http://127.0.0.1:${server.port}/api/ipc/health`, {
      signal: AbortSignal.timeout(5_000),
    })
    const body = (await response.json()) as { ok: boolean; channels: number }
    expect(body.ok).toBe(true)
    expect(body.channels).toBeGreaterThanOrEqual(5)
  })

  it('rejects malformed bodies with 400', async () => {
    const response = await fetch(invokeUrl(server.port, 'test:echo'), {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: '{"args": "not-an-array"}',
      signal: AbortSignal.timeout(5_000),
    })
    expect(response.status).toBe(400)
  })
})

describe('attachIpcMain wraps an electron-shaped ipcMain', () => {
  it('mirrors handle/on registrations and keeps original behavior', async () => {
    const ipcMain = new FakeIpcMain()
    const registry = new IpcHandlerRegistry()
    attachIpcMain(registry, ipcMain)

    ipcMain.handle('app:lang', () => 'zh')
    ipcMain.on('dirty:changed', (event, value: unknown) => {
      ;(event as { seen: unknown }).seen = value
    })

    expect(ipcMain.handles.has('app:lang')).toBe(true)
    expect(registry.handlerFor('app:lang')).toBeDefined()

    const server = await createBridgeServer({ registry, port: 0 })
    const language = await fetch(invokeUrl(server.port, 'app:lang'), { method: 'POST', body: '{}' })
    expect(((await language.json()) as { result: string }).result).toBe('zh')

    // .on-style channels answer 202 and dispatch to every registered listener
    const seen: unknown[] = []
    registry.addListener('dirty:changed', (event, _value) => void event)
    ipcMain.removeAllListeners?.('dirty:changed')
    registry.addListener('dirty:changed', (_event, value) => seen.push(value))
    const posted = await fetch(invokeUrl(server.port, 'dirty:changed'), {
      method: 'POST',
      body: JSON.stringify({ args: [true] }),
    })
    expect(posted.status).toBe(202)
    expect(((await posted.json()) as { delivered: number }).delivered).toBe(1)
    expect(seen).toEqual([true])
    await server.close()
  })

  it('preserves duplicate-handle rejection and removeHandler sync', () => {
    const ipcMain = new FakeIpcMain()
    const registry = new IpcHandlerRegistry()
    attachIpcMain(registry, ipcMain)
    ipcMain.handle('dup', () => 1)
    expect(() => ipcMain.handle('dup', () => 2)).toThrow(/second handler/)
    expect(registry.handlerFor('dup')).toBeDefined()
    ipcMain.removeHandler('dup')
    expect(registry.handlerFor('dup')).toBeUndefined()
    expect(ipcMain.handles.has('dup')).toBe(false)
  })

  it('is idempotent — a second attach does not double-wrap', () => {
    const ipcMain = new FakeIpcMain()
    const registry = new IpcHandlerRegistry()
    attachIpcMain(registry, ipcMain)
    ipcMain.handle('once-only', () => 1)
    attachIpcMain(registry, ipcMain)
    expect(ipcMain.handles.size).toBe(1)
    expect(registry.invokeChannelCount).toBe(1)
  })
})

describe('installHttpIpcBridge (the app one-liner)', () => {
  it('captures registrations made through the wrapped ipcMain and serves them over HTTP', async () => {
    const ipcMain = new FakeIpcMain()
    const bridge = await installHttpIpcBridge({ ipcMain, port: 0 })
    expect(bridge).not.toBeNull()
    ipcMain.handle('late:channel', (_event, value: unknown) => ({ got: value }))
    const response = await fetch(invokeUrl(bridge!.port, 'late:channel'), {
      method: 'POST',
      body: JSON.stringify({ args: ['v'] }),
    })
    expect(((await response.json()) as { result: { got: string } }).result).toEqual({ got: 'v' })
    await bridge!.close()
  })

  it('returns null instead of throwing when the port is already taken', async () => {
    const first = await installHttpIpcBridge({ ipcMain: new FakeIpcMain(), port: 0 })
    const busy = await installHttpIpcBridge({
      ipcMain: new FakeIpcMain(),
      port: first!.port,
      log: () => {},
    })
    expect(busy).toBeNull()
    await first!.close()
  })
})

describe('static hosting (production web form)', () => {
  let staticDir: string
  let server: BridgeServer

  beforeAll(() => {
    staticDir = join(tmpdir(), `ipc-bridge-static-${Date.now()}`)
    mkdirSync(staticDir, { recursive: true })
    writeFileSync(join(staticDir, 'index.html'), '<html><body>web-docs</body></html>')
    mkdirSync(join(staticDir, 'assets'))
    writeFileSync(join(staticDir, 'assets', 'app.js'), 'console.log(1)')
  })

  afterAll(async () => {
    rmSync(staticDir, { recursive: true, force: true })
    await server.close()
  })

  it('serves index.html and assets from staticDir', async () => {
    server = await createBridgeServer({
      registry: new IpcHandlerRegistry(),
      port: 0,
      staticDir,
    })
    const root = await fetch(`http://127.0.0.1:${server.port}/`)
    expect(root.headers.get('content-type')).toContain('text/html')
    expect(await root.text()).toContain('web-docs')
    const asset = await fetch(`http://127.0.0.1:${server.port}/assets/app.js`)
    expect(asset.headers.get('content-type')).toContain('text/javascript')
  })

  it('rejects path traversal outside staticDir', async () => {
    const response = await fetch(`http://127.0.0.1:${server.port}/..%2F..%2Fetc%2Fpasswd`, {
      signal: AbortSignal.timeout(5_000),
    })
    expect(response.status).toBe(404)
  })
})
