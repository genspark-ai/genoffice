import { appendFileSync } from 'node:fs'
import net from 'node:net'
import { egressDecision, externalOpenNeedsConfirm, isLoopbackHost } from './allowlist'

/**
 * Local-only fork: deny-by-default network guard for the shell main process.
 *
 * Layers (each one alone blocks a public host):
 * - Node: `net.Socket.prototype.connect` (covers http, https, tls, ws and the
 *   built-in fetch, which all open sockets through it) plus `globalThis.fetch`
 *   for an early, clear error.
 * - Node: `net.Server.prototype.listen` binds loopback only.
 * - Chromium, per session: a black-hole proxy (loopback bypassed) and a
 *   `webRequest` filter. Together they cover renderers, iframes, `net.fetch`
 *   and Chromium's own background requests.
 * - `shell.openExternal`: a web link opens in the browser only after a confirm.
 * - Child processes: proxy environment variables point at the black hole.
 *
 * See docs/local-only-fork.md.
 */

export const OFFLINE_ERROR_CODE = 'EGENOFFICE_OFFLINE'

/** Closed port on loopback: every proxied connection is refused at once. */
export const BLACK_HOLE_PROXY = 'http://127.0.0.1:9'
const LOOPBACK_BYPASS = 'localhost,127.0.0.1,[::1]'

export type EgressKind = 'socket' | 'fetch' | 'listen' | 'chromium' | 'open-external'

export interface EgressEvent {
  kind: EgressKind
  host: string
  /** where the attempt came from: a stack frame, or Chromium's resource type */
  via?: string
}

export type EgressReporter = (event: EgressEvent) => void

export class OfflineBlockedError extends Error {
  readonly code = OFFLINE_ERROR_CODE
  constructor(
    readonly host: string,
    readonly kind: EgressKind,
  ) {
    super(`Local-only build: blocked ${kind} connection to ${host}`)
    this.name = 'OfflineBlockedError'
  }
}

/** Frames that belong to the guard itself; by name too, since bundles merge files. */
const GUARD_FRAME = /egress-guard|node:|callerFrame|guarded(Connect|Listen|Fetch)/

/** First stack frame outside this module and Node internals. */
function callerFrame(): string | undefined {
  const lines = (new Error().stack ?? '').split('\n').slice(1)
  for (const line of lines) {
    const frame = line.trim()
    if (GUARD_FRAME.test(frame)) continue
    return frame.replace(/^at\s+/, '').slice(0, 300)
  }
  return undefined
}

/**
 * Appends one JSON line per new (kind, host) pair: time, kind, host and the
 * caller. Never URLs, paths or payloads.
 */
export function createBlockLog(logPath: () => string | null): EgressReporter {
  const seen = new Set<string>()
  return (event) => {
    const key = `${event.kind} ${event.host}`
    if (seen.has(key)) return
    seen.add(key)
    console.warn(`[local-only] blocked ${event.kind} to ${event.host}`)
    try {
      const path = logPath()
      if (!path) return
      const line = { time: new Date().toISOString(), ...event }
      appendFileSync(path, `${JSON.stringify(line)}\n`)
    } catch {
      // logging must never break the guard
    }
  }
}

// ---- Node layer ----

type ConnectTarget = { host: string } | { path: string }

/** Where a `Socket.connect(...)` call points, across its overloads. */
export function connectTarget(args: readonly unknown[]): ConnectTarget {
  // net.createConnection passes its normalized [options, callback] pair
  const first = Array.isArray(args[0]) ? args[0][0] : args[0]
  if (first && typeof first === 'object') {
    const o = first as { host?: unknown; path?: unknown }
    if (typeof o.path === 'string' && o.path) return { path: o.path }
    return { host: typeof o.host === 'string' && o.host ? o.host : 'localhost' }
  }
  if (typeof first === 'string' && !/^\d+$/.test(first)) return { path: first }
  const host = args[1]
  return { host: typeof host === 'string' && host ? host : 'localhost' }
}

/**
 * `Server.listen(...)` arguments rewritten to bind loopback. An unspecified
 * host (all interfaces) becomes 127.0.0.1; an explicit non-loopback host
 * throws. Pipes, paths and inherited handles pass through unchanged.
 */
export function loopbackListenArgs(args: readonly unknown[]): unknown[] {
  const first = args[0]
  if (first && typeof first === 'object') {
    const o = first as { host?: unknown; path?: unknown; fd?: unknown; _handle?: unknown }
    if (typeof o.path === 'string' || o.fd !== undefined || o._handle !== undefined) {
      return [...args]
    }
    if (typeof o.host === 'string' && o.host) {
      if (!isLoopbackHost(o.host)) throw new OfflineBlockedError(o.host, 'listen')
      return [...args]
    }
    return [{ ...o, host: '127.0.0.1' }, ...args.slice(1)]
  }
  if (typeof first === 'string' && !/^\d+$/.test(first)) return [...args]
  if (first === undefined || typeof first === 'function') {
    return [0, '127.0.0.1', ...args.filter((a) => a !== undefined)]
  }
  const host = args[1]
  if (typeof host === 'string') {
    if (!isLoopbackHost(host)) throw new OfflineBlockedError(host, 'listen')
    return [...args]
  }
  return [first, '127.0.0.1', ...args.slice(1)]
}

function requestUrl(input: unknown): string | null {
  if (typeof input === 'string') return input
  if (input instanceof URL) return input.href
  if (input && typeof input === 'object' && typeof (input as { url?: unknown }).url === 'string') {
    return (input as { url: string }).url
  }
  return null
}

/** Patches Node's socket, listen and fetch. Returns an uninstall for tests. */
export function installNodeEgressGuard(report: EgressReporter): () => void {
  const socketProto = net.Socket.prototype
  const serverProto = net.Server.prototype
  const originalConnect = socketProto.connect
  const originalListen = serverProto.listen
  const originalFetch = globalThis.fetch

  socketProto.connect = function guardedConnect(this: net.Socket, ...args: unknown[]) {
    const target = connectTarget(args)
    if ('host' in target && !isLoopbackHost(target.host)) {
      const error = new OfflineBlockedError(target.host, 'socket')
      report({ kind: 'socket', host: target.host, via: callerFrame() })
      // fail like a refused connection, so callers' error handling runs
      process.nextTick(() => this.destroy(error))
      return this
    }
    return (originalConnect as (...a: unknown[]) => net.Socket).apply(this, args)
  } as typeof socketProto.connect

  serverProto.listen = function guardedListen(this: net.Server, ...args: unknown[]) {
    let fixed: unknown[]
    try {
      fixed = loopbackListenArgs(args)
    } catch (error) {
      if (error instanceof OfflineBlockedError) {
        report({ kind: 'listen', host: error.host, via: callerFrame() })
      }
      throw error
    }
    return (originalListen as (...a: unknown[]) => net.Server).apply(this, fixed)
  } as typeof serverProto.listen

  if (typeof originalFetch === 'function') {
    globalThis.fetch = async function guardedFetch(input, init) {
      const url = requestUrl(input)
      const decision = url === null ? { allowed: true as const } : egressDecision(url)
      if (!decision.allowed) {
        report({ kind: 'fetch', host: decision.host, via: callerFrame() })
        throw new OfflineBlockedError(decision.host, 'fetch')
      }
      return originalFetch(input, init)
    } as typeof fetch
  }

  return () => {
    socketProto.connect = originalConnect
    serverProto.listen = originalListen
    if (typeof originalFetch === 'function') globalThis.fetch = originalFetch
  }
}

/** Points child processes that honor proxy variables at the black hole. */
export function pinChildProxyEnv(env: NodeJS.ProcessEnv): void {
  for (const name of ['HTTP_PROXY', 'HTTPS_PROXY', 'ALL_PROXY']) {
    env[name] = BLACK_HOLE_PROXY
    env[name.toLowerCase()] = BLACK_HOLE_PROXY
  }
  env.NO_PROXY = LOOPBACK_BYPASS
  env.no_proxy = LOOPBACK_BYPASS
}

// ---- Electron layer (structural types, so tests need no Electron) ----

interface GuardableSession {
  setProxy(config: { mode?: string; proxyRules?: string; proxyBypassRules?: string }): Promise<void>
  webRequest: {
    onBeforeRequest(
      listener: (
        details: { url: string; resourceType?: string },
        callback: (response: { cancel?: boolean }) => void,
      ) => void,
    ): void
  }
}

export interface ElectronEgressDeps {
  app: {
    on(event: 'session-created', listener: (session: GuardableSession) => void): unknown
    whenReady(): Promise<unknown>
  }
  session: { readonly defaultSession: GuardableSession }
  shell: { openExternal(url: string, options?: unknown): Promise<void> }
  dialog: {
    showMessageBox(options: {
      type: string
      buttons: string[]
      defaultId: number
      cancelId: number
      message: string
      detail: string
    }): Promise<{ response: number }>
  }
}

export function guardSession(session: GuardableSession, report: EgressReporter): void {
  void session
    .setProxy({
      mode: 'fixed_servers',
      proxyRules: BLACK_HOLE_PROXY,
      proxyBypassRules: LOOPBACK_BYPASS,
    })
    .catch(() => {
      // the webRequest filter below still blocks
    })
  session.webRequest.onBeforeRequest((details, callback) => {
    const decision = egressDecision(details.url)
    if (decision.allowed) {
      callback({})
      return
    }
    report({ kind: 'chromium', host: decision.host, via: details.resourceType })
    callback({ cancel: true })
  })
}

export function installElectronEgressGuard(deps: ElectronEgressDeps, report: EgressReporter): void {
  const guarded = new WeakSet<GuardableSession>()
  const guard = (session: GuardableSession): void => {
    if (guarded.has(session)) return
    guarded.add(session)
    guardSession(session, report)
  }
  deps.app.on('session-created', guard)
  void deps.app.whenReady().then(() => guard(deps.session.defaultSession))

  const originalOpen = deps.shell.openExternal.bind(deps.shell)
  deps.shell.openExternal = async (url: string, options?: unknown): Promise<void> => {
    if (!externalOpenNeedsConfirm(url)) return originalOpen(url, options)
    const host = new URL(url).hostname
    const { response } = await deps.dialog.showMessageBox({
      type: 'warning',
      buttons: ['Cancel', 'Open in browser'],
      defaultId: 0,
      cancelId: 0,
      message: `Open ${host} in your web browser?`,
      detail:
        'This local-only build of GenOffice never connects to the internet. ' +
        'The link opens in your default browser, outside GenOffice.',
    })
    if (response === 1) return originalOpen(url, options)
    report({ kind: 'open-external', host })
  }
}
