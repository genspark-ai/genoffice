/// Browser-safe transport layer for the dual-protocol IPC bridge.
///
/// The renderer-facing `window.*` APIs are built on an {@link IpcTransport}.
/// Two implementations exist for the same channel surface:
///   • Electron renderer → {@link createElectronIpcTransport} (thin wrapper over
///     `ipcRenderer.invoke/send/on`, byte-for-byte the previous behavior);
///   • plain browser (web version) → {@link createHttpIpcTransport}, which maps
///     `invoke` to `POST /api/ipc/:channel` and push listeners to one SSE stream
///     per session (`GET /api/ipc/events`).
///
/// This module must stay free of node/electron imports: the sandboxed preload
/// bundles it and the web bridge loads it directly in the browser.

import { decodeTransportValue, encodeTransportValue } from './codec'

export interface IpcTransport {
  /**
   * `ipcRenderer.invoke(channel, ...args)` equivalent; resolves with the handler
   * result. Typed like Electron (`Promise<any>`) so preload-extracted API
   * factories keep their contextual return types across both transports.
   */
  invoke(channel: string, ...args: unknown[]): Promise<any>
  /** `ipcRenderer.send(channel, ...args)` equivalent — fire-and-forget, no result. */
  send(channel: string, ...args: unknown[]): void
  /**
   * Subscribe to a main-process push channel. Listeners receive the payload
   * arguments only (no event object) — the same shape the SSE bridge delivers.
   * Returns an unsubscribe function, matching the preload listener contract.
   */
  on(channel: string, listener: (...args: unknown[]) => void): () => void
}

/** Error code carried by structured desktop-only rejections (see the bridge server). */
export const WEB_UNSUPPORTED = 'WEB_UNSUPPORTED'

/** Error thrown by the HTTP transport; `code` mirrors the bridge error object. */
export class IpcBridgeError extends Error {
  readonly code: string | undefined

  constructor(message: string, code?: string) {
    super(message)
    this.name = 'IpcBridgeError'
    this.code = code
  }
}

/** True when running inside an Electron renderer (preload already exposed window.* APIs). */
export function isElectronRuntime(): boolean {
  return typeof navigator !== 'undefined' && /Electron/i.test(navigator.userAgent)
}

/** Structural subset of Electron's IpcRenderer the Electron transport needs. */
export interface IpcRendererLike {
  invoke(channel: string, ...args: unknown[]): Promise<unknown>
  send(channel: string, ...args: unknown[]): void
  on(channel: string, listener: (event: unknown, ...args: unknown[]) => void): unknown
  removeListener(channel: string, listener: unknown): unknown
}

export function createElectronIpcTransport(ipcRenderer: IpcRendererLike): IpcTransport {
  return {
    invoke: (channel, ...args) => ipcRenderer.invoke(channel, ...args),
    send: (channel, ...args) => {
      ipcRenderer.send(channel, ...args)
    },
    on: (channel, listener) => {
      const wrapped = (_event: unknown, ...args: unknown[]) => listener(...args)
      ipcRenderer.on(channel, wrapped)
      return () => {
        ipcRenderer.removeListener(channel, wrapped)
      }
    },
  }
}

export interface HttpIpcTransportOptions {
  /**
   * Bridge origin, e.g. 'http://127.0.0.1:5273'. Empty (default) talks to the
   * serving origin itself — in dev the vite dev server proxies `/api` to the
   * bridge, and in production web mode the bridge serves the app statically,
   * so everything stays same-origin under the page CSP (connect-src 'self').
   */
  baseUrl?: string
}

export function createHttpIpcTransport(options: HttpIpcTransportOptions = {}): IpcTransport {
  const base = (options.baseUrl ?? '').replace(/\/+$/, '')
  const session = createSessionId()
  const pushHub = createPushHub(base, session)

  async function invoke(channel: string, ...args: unknown[]): Promise<unknown> {
    let response: Response
    try {
      response = await fetch(`${base}/api/ipc/${encodeURIComponent(channel)}`, {
        method: 'POST',
        headers: { 'content-type': 'application/json', 'x-ipc-session': session },
        body: JSON.stringify({ args: args.map((arg) => encodeTransportValue(arg)) }),
      })
    } catch (cause) {
      throw new IpcBridgeError(
        `IPC bridge unreachable (${base || 'same origin'}): ${String(cause)}`,
        'IPC_BRIDGE_UNREACHABLE',
      )
    }
    const body = (await response.json().catch(() => null)) as
      { ok: true; result: unknown } | { error: { message?: string; code?: string } } | null
    if (response.ok && body && 'ok' in body && body.ok === true) {
      return decodeTransportValue(body.result)
    }
    const error = body && 'error' in body && body.error ? body.error : {}
    throw new IpcBridgeError(
      error.message || `IPC call to '${channel}' failed with HTTP ${response.status}`,
      error.code,
    )
  }

  return {
    invoke,
    send: (channel, ...args) => {
      // No delivery guarantee, matching ipcRenderer.send semantics.
      void invoke(channel, ...args).catch((cause) => {
        console.warn(`[ipc-bridge] send('${channel}') failed:`, cause)
      })
    },
    on: (channel, listener) => pushHub.on(channel, listener),
  }
}

function createSessionId(): string {
  if (typeof crypto !== 'undefined' && 'randomUUID' in crypto) return crypto.randomUUID()
  return `s-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`
}

/**
 * One SSE stream per transport instance, shared by every `on()` subscriber.
 * Pushes emitted before the stream connects (or during a reconnect) are
 * buffered server-side per session, so subscribing after an `invoke` — or a
 * stream that reconnects — still receives every frame.
 */
function createPushHub(base: string, session: string) {
  const listeners = new Map<string, Set<(...args: unknown[]) => void>>()
  let source: EventSource | null = null

  function ensureSource(): void {
    if (source || typeof EventSource === 'undefined') return
    source = new EventSource(`${base}/api/ipc/events?session=${encodeURIComponent(session)}`)
    source.onmessage = (event) => {
      let frame: { channel: string; args: unknown[] }
      try {
        frame = JSON.parse(event.data as string)
      } catch {
        return
      }
      const args = (frame.args ?? []).map((arg) => decodeTransportValue(arg))
      for (const listener of listeners.get(frame.channel) ?? []) {
        try {
          listener(...args)
        } catch (cause) {
          console.warn(`[ipc-bridge] listener for '${frame.channel}' threw:`, cause)
        }
      }
    }
    // EventSource reconnects on its own; dropped frames are covered by the
    // server-side session buffer while the stream is down.
    source.onerror = () => {}
  }

  return {
    on(channel: string, listener: (...args: unknown[]) => void): () => void {
      ensureSource()
      let set = listeners.get(channel)
      if (!set) {
        set = new Set()
        listeners.set(channel, set)
      }
      set.add(listener)
      return () => {
        set.delete(listener)
      }
    },
  }
}
