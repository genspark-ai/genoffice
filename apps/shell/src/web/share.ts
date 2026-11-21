/**
 * Runtime start/stop for the web service, behind the editors' status-bar
 * share chips. `startWebServer` itself stays launch-shaped (one boot path);
 * this controller owns the at-most-one-live handle, derives the shareable
 * LAN URL, and reports state so every open editor reflects the same service.
 */
import { networkInterfaces } from 'node:os'
import { randomBytes } from 'node:crypto'
import { webContents } from 'electron'
import { startWebServer, webBroadcast, type WebServerConfig, type WebServerHandle } from './index'

export interface WebShareStatus {
  running: boolean
  /** `http://<lan-ip>:<port>/?token=…` when running — the URL to hand out */
  url?: string
  /** the URL that only reaches the server machine itself */
  localUrl?: string
  port?: number
  /** set when the last start attempt failed (port busy, …) */
  error?: string
}

/** DI hook for tests: the live handle, so tests never bind a real socket */
type HandleFactory = (overrides: Partial<WebServerConfig>) => Promise<WebServerHandle>
let handleFactory: HandleFactory = (overrides) => startWebServer(overrides)

/** @internal test seam */
export function setWebShareHandleFactory(factory: HandleFactory): void {
  handleFactory = factory
}

let active: WebServerHandle | null = null
let lastError: string | undefined

/** first non-internal IPv4 across the machine's interfaces, else loopback */
export function lanAddress(
  interfaces: ReturnType<typeof networkInterfaces> = networkInterfaces(),
): string {
  for (const list of Object.values(interfaces)) {
    for (const entry of list ?? []) {
      if (entry.family !== 'IPv4' || entry.internal) continue
      return entry.address
    }
  }
  return '127.0.0.1'
}

/** the URL to share: a LAN address when one exists, loopback otherwise */
export function shareUrlFor(config: WebServerConfig, address = lanAddress()): string {
  return `http://${address}:${config.port}/?token=${config.token}`
}

const localUrlFor = (config: WebServerConfig): string =>
  `http://127.0.0.1:${config.port}/?token=${config.token}`

export function webShareStatus(): WebShareStatus {
  if (!active) return { running: false, ...(lastError ? { error: lastError } : {}) }
  return {
    running: true,
    url: shareUrlFor(active.config),
    localUrl: localUrlFor(active.config),
    port: active.config.port,
  }
}

function announce(): void {
  const status = webShareStatus()
  webBroadcast('web-share:changed', [status])
  for (const wc of webContents.getAllWebContents()) {
    if (!wc.isDestroyed()) wc.send('web-share:changed', status)
  }
}

/** Start the service for sharing (LAN-reachable, fresh token) unless it is
 *  already running — an existing service is kept, so a second chip press
 *  anywhere in the suite never disturbs sessions someone is using. */
export async function startWebShare(
  /** boot-path argv overrides (--web-port/--web-token…); chip presses pass
   *  nothing and get the LAN defaults. Ignored while a server is live. */
  overrides: Partial<WebServerConfig> = {},
): Promise<WebShareStatus> {
  if (active) return webShareStatus()
  lastError = undefined
  try {
    // 0.0.0.0 = reachable from the network (that is the point of sharing);
    // the token is always freshly generated, never reused across runs
    active = await handleFactory({
      host: '0.0.0.0',
      token: randomBytes(24).toString('hex'),
      ...overrides,
    })
  } catch (err) {
    lastError = err instanceof Error ? err.message : String(err)
    announce()
    return webShareStatus()
  }
  announce()
  return webShareStatus()
}

export function stopWebShare(): WebShareStatus {
  if (active) {
    active.close()
    active = null
    announce()
  }
  lastError = undefined
  return webShareStatus()
}

/** @internal — the boot path (--web-server argv) reports through the same
 *  controller so the chips reflect a server that started before any UI. */
export function adoptWebServerHandle(handle: WebServerHandle): void {
  active = handle
  announce()
}

export function webShareActiveHandle(): WebServerHandle | null {
  return active
}
