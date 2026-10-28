/**
 * The `electron` module as it exists inside a browser tab running a GenOffice
 * preload bundle. contextBridge degrades to plain window assignment (the page
 * is our own origin; there is no isolated world to bridge into), ipcRenderer
 * rides the web bridge client (a single /__web/client.js script owns the
 * WebSocket; these stubs reach it through window.__genofficeWebClient), and
 * webUtils cannot resolve OS file paths in a browser (the drop-open bridge
 * tolerates that and stays out of the way).
 */

interface WebClient {
  invoke(channel: string, args: unknown[]): Promise<unknown>
  send(channel: string, args: unknown[]): void
}

type Listener = (event: unknown, ...args: unknown[]) => void

const listeners = new Map<string, Set<Listener>>()

function client(): WebClient {
  const c = (window as unknown as Record<string, unknown>).__genofficeWebClient as
    WebClient | undefined
  if (!c) throw new Error('web bridge client not loaded — /__web/client.js must run first')
  return c
}

export const ipcRenderer = {
  invoke(channel: string, ...args: unknown[]): Promise<unknown> {
    return client().invoke(channel, args)
  },
  send(channel: string, ...args: unknown[]): void {
    client().send(channel, args)
  },
  /** the web bridge has no synchronous round trip; the one sendSync caller
   *  (docs context-menu claim) ignores the return value */
  sendSync(channel: string, ...args: unknown[]): undefined {
    client().send(channel, args)
    return undefined
  },
  on(channel: string, handler: Listener): void {
    const set = listeners.get(channel) ?? new Set<Listener>()
    set.add(handler)
    listeners.set(channel, set)
  },
  once(channel: string, handler: Listener): void {
    const wrapped: Listener = (event, ...args) => {
      ipcRenderer.removeListener(channel, wrapped)
      handler(event, ...args)
    }
    ipcRenderer.on(channel, wrapped)
  },
  removeListener(channel: string, handler: Listener): void {
    listeners.get(channel)?.delete(handler)
  },
  removeAllListeners(channel: string): void {
    listeners.delete(channel)
  },
}

export const contextBridge = {
  exposeInMainWorld(name: string, value: unknown): void {
    ;(window as unknown as Record<string, unknown>)[name] = value
  },
  exposeInIsolatedWorld(): void {
    // no isolated world in the browser shim
  },
}

export const webUtils = {
  getPathForFile(): string {
    return ''
  },
}

export default { ipcRenderer, contextBridge, webUtils }

/** the client script dispatches main→renderer pushes through this hook; one
 *  preload bundle runs per page, so a single dispatch slot is enough */
;(window as unknown as Record<string, unknown>).__genofficeWebDispatch = (
  channel: string,
  args: unknown[],
): void => {
  const event = { sender: null, senderId: -1 }
  for (const handler of [...(listeners.get(channel) ?? [])]) {
    try {
      handler(event, ...args)
    } catch (err) {
      console.error(`[web-shim] listener for ${channel}:`, err)
    }
  }
}
