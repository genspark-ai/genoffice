/// HTTP + SSE bridge in front of the existing `ipcMain` handler registry.
///
/// genoffice's dual-protocol story in one place: the bridge wraps the Electron
/// `ipcMain` registration methods ONCE, so every channel any app registers —
/// current or future — is automatically served over both transports:
///
///   • IPC   (Electron renderer)  → unchanged `ipcMain.handle`/`on` behavior
///   • HTTP  (browser web version) → POST /api/ipc/:channel      (invoke/send)
///                                   GET  /api/ipc/events?session= (push stream)
///
/// Handlers receive a bridge event whose `sender.send` routes into the calling
/// session's SSE stream, so main→renderer push (AI stream chunks, …) works
/// unchanged. Handlers that need real native resources (dialogs, printing,
/// webContents registries) either fail naturally or are blocked up-front via
/// `nativeOnlyChannels` — both surface as structured errors on the web side
/// instead of an invisible hang.
///
/// The bridge binds 127.0.0.1 only and a listen failure never breaks the
/// desktop app: `installHttpIpcBridge` logs and returns null. Install it at the
/// very top of the app's main bootstrap, BEFORE the first `register*Ipc()` call,
/// so the wrapper sees every registration.

import { createServer, type IncomingMessage, type ServerResponse } from 'node:http'
import { createReadStream, existsSync, mkdirSync, mkdtempSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { basename, extname, join, resolve, sep } from 'node:path'
import { tmpdir } from 'node:os'
import { decodeTransportValue, encodeTransportValue } from './codec'
import { WEB_UNSUPPORTED } from './client'
import { WEB_FILE_CHANNELS } from './web-native'

export { WEB_UNSUPPORTED, IpcBridgeError } from './client'
export type { IpcTransport } from './client'

/** Error codes carried in the bridge's structured error objects. */
export const IPC_NO_HANDLER = 'IPC_NO_HANDLER'
export const IPC_BRIDGE_UNREACHABLE = 'IPC_BRIDGE_UNREACHABLE'

/** Web-file channels the bridge serves itself (no app registration needed). */
export { WEB_FILE_CHANNELS } from './web-native'

/** Longest allowed POST body — the largest real payloads are document saves. */
const DEFAULT_BODY_LIMIT_BYTES = 256 * 1024 * 1024
/** Push frames buffered for a session while its SSE stream is (re)connecting. */
const MAX_PENDING_FRAMES_PER_SESSION = 500
const PENDING_FRAME_TTL_MS = 60_000
const SSE_HEARTBEAT_MS = 25_000

/** Root for web-uploaded temp files; only files under it are readable back. */
const WEB_TEMP_ROOT = join(tmpdir(), 'genoffice-web-bridge')

/**
 * Generic web-file channels: the browser uploads picked-file bytes and gets a
 * temp path the app's normal open/add flow can consume; saved/exported bytes
 * are read back for a browser download. Reads are confined to the temp root so
 * a loopback caller cannot exfiltrate arbitrary files.
 */
function installWebFileChannels(registry: IpcHandlerRegistry): void {
  mkdirSyncSafe(WEB_TEMP_ROOT)
  registry.registerHandle(WEB_FILE_CHANNELS.writeTempFile, (_event, request: unknown) => {
    const record = request as { name?: unknown; bytes?: unknown } | null
    if (!record || typeof record.name !== 'string' || !(record.bytes instanceof ArrayBuffer)) {
      throw new Error('web:write-temp-file expects { name: string, bytes: ArrayBuffer }')
    }
    const safeName = basename(record.name).replace(/[^\w.\- ]+/g, '_') || 'file'
    const dir = mkdtempSync(join(WEB_TEMP_ROOT, 'upload-'))
    const filePath = join(dir, safeName)
    writeFileSync(filePath, Buffer.from(record.bytes))
    return filePath
  })
  registry.registerHandle(WEB_FILE_CHANNELS.readFileBytes, (_event, path: unknown) => {
    if (typeof path !== 'string' || !path.startsWith(WEB_TEMP_ROOT + sep)) {
      throw new Error('web:read-file-bytes only reads files written by the web bridge')
    }
    const bytes = readFileSync(path)
    return { name: basename(path), bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) }
  })
  registry.registerHandle(WEB_FILE_CHANNELS.makeTempDir, () => mkdtempSync(join(WEB_TEMP_ROOT, 'dir-')))
}

function mkdirSyncSafe(dir: string): void {
  try {
    mkdirSync(dir, { recursive: true })
  } catch {
    // best effort — a failed temp root only disables web file uploads
  }
}

export interface BridgeWebContentsLike {
  readonly id: number
  isDestroyed(): boolean
  send(channel: string, ...args: unknown[]): void
}

/** The subset of `IpcMainInvokeEvent` handlers may rely on across the bridge. */
export interface BridgeIpcEvent {
  readonly processId: number
  readonly frameId: number
  readonly sender: BridgeWebContentsLike
}

export type IpcInvokeHandler = (event: BridgeIpcEvent, ...args: unknown[]) => unknown
export type IpcEventListener = (event: BridgeIpcEvent, ...args: unknown[]) => void

/**
 * Mirrors the registration methods the bridge intercepts. Structural so both
 * Electron's `ipcMain` and test doubles satisfy it without an electron import.
 */
export interface IpcMainLike {
  handle(channel: string, handler: (event: unknown, ...args: unknown[]) => unknown): void
  removeHandler(channel: string): void
  on(channel: string, listener: (event: unknown, ...args: unknown[]) => void): unknown
  once?(channel: string, listener: (event: unknown, ...args: unknown[]) => void): unknown
  removeListener?(channel: string, listener: unknown): unknown
  removeAllListeners?(channel: string): unknown
}

/** Registry of everything the wrapped `ipcMain` methods have seen. */
export class IpcHandlerRegistry {
  private readonly handlers = new Map<string, IpcInvokeHandler>()
  private readonly listeners = new Map<string, Set<IpcEventListener>>()

  /** Mirrors Electron's duplicate-handle rejection with the same message. */
  registerHandle(channel: string, handler: IpcInvokeHandler): void {
    if (this.handlers.has(channel)) {
      throw new Error(`Attempted to register a second handler for '${channel}'`)
    }
    this.handlers.set(channel, handler)
  }

  removeHandle(channel: string): void {
    this.handlers.delete(channel)
  }

  addListener(channel: string, listener: IpcEventListener): void {
    let set = this.listeners.get(channel)
    if (!set) {
      set = new Set()
      this.listeners.set(channel, set)
    }
    set.add(listener)
  }

  removeListener(channel: string, listener: IpcEventListener): void {
    this.listeners.get(channel)?.delete(listener)
  }

  clearListeners(channel: string): void {
    this.listeners.delete(channel)
  }

  get invokeChannelCount(): number {
    return this.handlers.size
  }

  get listenChannelCount(): number {
    return this.listeners.size
  }

  handlerFor(channel: string): IpcInvokeHandler | undefined {
    return this.handlers.get(channel)
  }

  /** Delivers to every live listener; returns how many accepted the event. */
  emitToListeners(channel: string, event: BridgeIpcEvent, args: unknown[]): number {
    const set = this.listeners.get(channel)
    if (!set) return 0
    for (const listener of [...set]) listener(event, ...args)
    return set.size
  }
}

const ATTACH_FLAG = Symbol('genofficeIpcBridgeAttached')

/**
 * Wrap an `ipcMain`-like object so every registration is mirrored into
 * `registry` while still executing the original Electron behavior. Idempotent:
 * attaching twice is a no-op. `handle`/`removeHandler`/`on`/`once`/
 * `removeListener`/`removeAllListeners` all stay in sync.
 */
export function attachIpcMain(registry: IpcHandlerRegistry, ipcMain: IpcMainLike): void {
  const writable = ipcMain as unknown as Record<PropertyKey, unknown>
  if (writable[ATTACH_FLAG] === true) return
  writable[ATTACH_FLAG] = true

  const originalHandle = ipcMain.handle.bind(ipcMain)
  writable.handle = (channel: string, handler: IpcInvokeHandler) => {
    // Throws on duplicates before Electron does — same message, same outcome.
    registry.registerHandle(channel, handler)
    return originalHandle(channel, handler as (event: unknown, ...args: unknown[]) => unknown)
  }

  const originalRemoveHandler = ipcMain.removeHandler.bind(ipcMain)
  writable.removeHandler = (channel: string) => {
    registry.removeHandle(channel)
    return originalRemoveHandler(channel)
  }

  const originalOn = ipcMain.on.bind(ipcMain)
  writable.on = (channel: string, listener: IpcEventListener) => {
    registry.addListener(channel, listener)
    return originalOn(channel, listener as (event: unknown, ...args: unknown[]) => void)
  }

  if (ipcMain.once) {
    const originalOnce = ipcMain.once.bind(ipcMain)
    writable.once = (channel: string, listener: IpcEventListener) => {
      const wrapped: IpcEventListener = (event, ...args) => {
        registry.removeListener(channel, wrapped)
        listener(event, ...args)
      }
      registry.addListener(channel, wrapped)
      return originalOnce(channel, wrapped as (event: unknown, ...args: unknown[]) => void)
    }
  }

  if (ipcMain.removeListener) {
    const originalRemoveListener = ipcMain.removeListener.bind(ipcMain)
    writable.removeListener = (channel: string, listener: unknown) => {
      registry.removeListener(channel, listener as IpcEventListener)
      return originalRemoveListener(
        channel,
        listener as (event: unknown, ...args: unknown[]) => void,
      )
    }
  }

  if (ipcMain.removeAllListeners) {
    const originalRemoveAllListeners = ipcMain.removeAllListeners.bind(ipcMain)
    writable.removeAllListeners = (channel: string) => {
      registry.clearListeners(channel)
      return originalRemoveAllListeners(channel)
    }
  }
}

/**
 * The `event.sender` handlers receive across HTTP. `send` routes into the
 * session's SSE stream; any other webContents member is a structured
 * desktop-only error instead of an undefined that would mislead handler code.
 */
function createBridgeSender(
  push: (channel: string, args: unknown[]) => void,
): BridgeWebContentsLike {
  const base = {
    id: -1,
    isDestroyed: () => false,
    send: (channel: string, ...args: unknown[]) => {
      push(channel, args)
    },
  }
  return new Proxy(base, {
    get(target, prop) {
      if (prop in target) return (target as Record<PropertyKey, unknown>)[prop]
      throw new Error(
        `${WEB_UNSUPPORTED}: event.sender.${String(prop)} is desktop-only and unavailable over the HTTP bridge`,
      )
    },
  })
}

interface PendingFrames {
  frames: string[]
  timer: NodeJS.Timeout
}

/** Session-isolated SSE fan-out with a reconnect buffer per session. */
class SessionHub {
  private readonly connections = new Map<string, Set<ServerResponse>>()
  private readonly pending = new Map<string, PendingFrames>()

  get sessionCount(): number {
    return this.connections.size + this.pending.size
  }

  push(session: string, frame: string): void {
    const connections = this.connections.get(session)
    if (connections) {
      for (const response of connections) response.write(frame)
      return
    }
    let entry = this.pending.get(session)
    if (!entry) {
      entry = {
        frames: [],
        timer: setTimeout(() => this.pending.delete(session), PENDING_FRAME_TTL_MS),
      }
      this.pending.set(session, entry)
    }
    entry.frames.push(frame)
    if (entry.frames.length > MAX_PENDING_FRAMES_PER_SESSION) entry.frames.shift()
  }

  attach(session: string, response: ServerResponse): () => void {
    let connections = this.connections.get(session)
    if (!connections) {
      connections = new Set()
      this.connections.set(session, connections)
    }
    connections.add(response)
    const entry = this.pending.get(session)
    if (entry) {
      clearTimeout(entry.timer)
      this.pending.delete(session)
      for (const frame of entry.frames) response.write(frame)
    }
    return () => {
      connections.delete(response)
      if (connections.size === 0) this.connections.delete(session)
    }
  }
}

export interface BridgeServerOptions {
  registry: IpcHandlerRegistry
  port: number
  /** Serve the built renderer for the production web form (out/renderer). */
  staticDir?: string
  /**
   * Channels that can never serve the web version (native dialogs, printing,
   * screen capture, …). Calls receive an immediate structured WEB_UNSUPPORTED
   * error instead of reaching a handler that would open an invisible dialog
   * and hang the web caller.
   */
  nativeOnlyChannels?: Array<string | RegExp>
  bodyLimitBytes?: number
  log?: (message: string) => void
}

export interface BridgeServer {
  port: number
  url: string
  close(): Promise<void>
}

function matchesNativeOnly(channel: string, patterns: Array<string | RegExp>): boolean {
  return patterns.some((pattern) =>
    typeof pattern === 'string' ? pattern === channel : pattern.test(channel),
  )
}

function sendJson(response: ServerResponse, status: number, payload: unknown): void {
  const body = JSON.stringify(payload)
  response.writeHead(status, { 'content-type': 'application/json' })
  response.end(body)
}

function readBody(request: IncomingMessage, limitBytes: number): Promise<string> {
  return new Promise((resolvePromise, rejectPromise) => {
    let size = 0
    const chunks: Buffer[] = []
    request.on('data', (chunk: Buffer) => {
      size += chunk.length
      if (size > limitBytes) {
        rejectPromise(new Error(`IPC bridge request body exceeds ${limitBytes} bytes`))
        request.destroy()
        return
      }
      chunks.push(chunk)
    })
    request.on('end', () => resolvePromise(Buffer.concat(chunks).toString('utf8')))
    request.on('error', rejectPromise)
  })
}

/** `{"args":[…]}` from the client; bare JSON accepted for curl-style checks. */
function extractArgs(parsed: unknown): unknown[] {
  if (parsed === null || parsed === undefined) return []
  if (
    typeof parsed === 'object' &&
    !Array.isArray(parsed) &&
    'args' in (parsed as Record<string, unknown>)
  ) {
    const args = (parsed as Record<string, unknown>).args
    if (!Array.isArray(args)) throw new Error("'args' must be an array")
    return args
  }
  return Array.isArray(parsed) ? parsed : [parsed]
}

const STATIC_MIME_TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript',
  '.mjs': 'text/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.ico': 'image/x-icon',
  '.txt': 'text/plain; charset=utf-8',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.wasm': 'application/wasm',
  '.map': 'application/json',
}

export function createBridgeServer(options: BridgeServerOptions): Promise<BridgeServer> {
  const log = options.log ?? (() => {})
  const bodyLimitBytes = options.bodyLimitBytes ?? DEFAULT_BODY_LIMIT_BYTES
  const nativeOnly = options.nativeOnlyChannels ?? []
  const hub = new SessionHub()
  const staticRoot = options.staticDir ? resolve(options.staticDir) : null

  const server = createServer((request, response) => {
    void handleRequest(request, response).catch((cause: unknown) => {
      if (!response.headersSent) {
        sendJson(response, 500, { error: { message: String((cause as Error)?.message ?? cause) } })
      } else {
        response.end()
      }
    })
  })

  async function handleRequest(request: IncomingMessage, response: ServerResponse): Promise<void> {
    const url = new URL(request.url ?? '/', 'http://127.0.0.1')

    if (url.pathname === '/api/ipc/health' && request.method === 'GET') {
      sendJson(response, 200, {
        ok: true,
        channels: options.registry.invokeChannelCount + options.registry.listenChannelCount,
        sessions: hub.sessionCount,
      })
      return
    }

    const invokeMatch = /^\/api\/ipc\/(.+)$/.exec(url.pathname)
    if (invokeMatch && request.method === 'POST') {
      const channel = decodeURIComponent(invokeMatch[1] ?? '')
      const session = request.headers['x-ipc-session']
      const sessionId = typeof session === 'string' ? session : undefined
      const push = (pushedChannel: string, args: unknown[]) => {
        if (!sessionId) return
        hub.push(
          sessionId,
          `data: ${JSON.stringify({
            channel: pushedChannel,
            args: args.map((arg) => encodeTransportValue(arg)),
          })}\n\n`,
        )
      }
      const event: BridgeIpcEvent = {
        processId: 0,
        frameId: 0,
        sender: createBridgeSender(push),
      }

      let args: unknown[]
      try {
        args = extractArgs(JSON.parse((await readBody(request, bodyLimitBytes)) || 'null')).map(
          (arg) => decodeTransportValue(arg),
        )
      } catch (cause) {
        sendJson(response, 400, { error: { message: String((cause as Error)?.message ?? cause) } })
        return
      }

      if (matchesNativeOnly(channel, nativeOnly)) {
        sendJson(response, 400, {
          error: {
            code: 'WEB_UNSUPPORTED',
            message: `'${channel}' is desktop-only and unavailable in the web version`,
          },
        })
        return
      }

      const handler = options.registry.handlerFor(channel)
      if (handler) {
        try {
          const result = await handler(event, ...args)
          sendJson(response, 200, { ok: true, result: encodeTransportValue(result) })
        } catch (cause) {
          // Same shape a rejected ipcRenderer.invoke surfaces to the renderer.
          sendJson(response, 500, {
            error: { message: String((cause as Error)?.message ?? cause) },
          })
        }
        return
      }

      const delivered = options.registry.emitToListeners(channel, event, args)
      if (delivered > 0) {
        sendJson(response, 202, { ok: true, delivered })
        return
      }

      sendJson(response, 404, {
        error: {
          code: IPC_NO_HANDLER,
          message: `No IPC handler registered for '${channel}'`,
        },
      })
      return
    }

    if (invokeMatch && request.method === 'GET' && url.pathname === '/api/ipc/events') {
      const session = url.searchParams.get('session')
      if (!session) {
        sendJson(response, 400, { error: { message: "Missing 'session' query parameter" } })
        return
      }
      response.writeHead(200, {
        'content-type': 'text/event-stream',
        'cache-control': 'no-cache',
        connection: 'keep-alive',
        'x-accel-buffering': 'no',
      })
      response.write(': connected\n\n')
      const detach = hub.attach(session, response)
      const heartbeat = setInterval(() => response.write(': hb\n\n'), SSE_HEARTBEAT_MS)
      request.on('close', () => {
        clearInterval(heartbeat)
        detach()
      })
      return
    }

    if (staticRoot) {
      if (serveStatic(url.pathname, staticRoot, response)) return
    }

    sendJson(response, 404, {
      error: { message: `No bridge route for ${request.method} ${url.pathname}` },
    })
  }

  function serveStatic(pathname: string, root: string, response: ServerResponse): boolean {
    const relative =
      pathname === '/' ? 'index.html' : decodeURIComponent(pathname).replace(/^\/+/, '')
    const filePath = resolve(root, relative)
    if (filePath !== root && !filePath.startsWith(root + sep)) return false
    if (!existsSync(filePath) || !statSync(filePath).isFile()) return false
    response.writeHead(200, {
      'content-type': STATIC_MIME_TYPES[extname(filePath)] ?? 'application/octet-stream',
    })
    createReadStream(filePath).pipe(response)
    return true
  }

  return new Promise((resolvePromise, rejectPromise) => {
    server.once('error', rejectPromise)
    server.listen(options.port, '127.0.0.1', () => {
      const address = server.address()
      const port = typeof address === 'object' && address ? address.port : options.port
      log(`[ipc-bridge] HTTP IPC bridge on http://127.0.0.1:${port}/api/ipc/* (web dual protocol)`)
      resolvePromise({
        port,
        url: `http://127.0.0.1:${port}`,
        close: () =>
          new Promise<void>((closeResolve, closeReject) => {
            server.close((cause) => (cause ? closeReject(cause) : closeResolve()))
          }),
      })
    })
  })
}

export interface HttpIpcBridgeOptions {
  ipcMain: IpcMainLike
  /** Loopback port for the HTTP form; per-app default, env-overridable at the call site. */
  port: number
  staticDir?: string
  nativeOnlyChannels?: Array<string | RegExp>
  bodyLimitBytes?: number
  log?: (message: string) => void
}

export interface HttpIpcBridge {
  port: number
  close(): Promise<void>
}

/**
 * One bridge per process: the shell aggregates every editor module's register*
 * calls on a single `ipcMain`, so only the FIRST install in a process wins and
 * later module installs reuse it (their registrations land in the same wrapped
 * registry anyway). Keyed by ipcMain identity so test doubles stay independent.
 */
const installedBridges = new WeakMap<IpcMainLike, HttpIpcBridge | null>()

/**
 * The one-line app integration. Installs the registration interceptor and
 * starts the loopback HTTP server. Never throws: when the port is taken the
 * desktop app keeps running without the web form and null is returned.
 */
export async function installHttpIpcBridge(
  options: HttpIpcBridgeOptions,
): Promise<HttpIpcBridge | null> {
  const existing = installedBridges.get(options.ipcMain)
  if (existing !== undefined) return existing
  const log = options.log ?? ((message: string) => console.log(message))
  const registry = new IpcHandlerRegistry()
  attachIpcMain(registry, options.ipcMain)
  installWebFileChannels(registry)
  let bridge: HttpIpcBridge | null
  try {
    bridge = await createBridgeServer({ ...options, registry, log })
  } catch (cause) {
    log(
      `[ipc-bridge] disabled — port ${options.port} unavailable: ${String((cause as Error)?.message ?? cause)}`,
    )
    bridge = null
  }
  installedBridges.set(options.ipcMain, bridge)
  return bridge
}
