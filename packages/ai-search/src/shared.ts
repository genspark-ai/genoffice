/** Search result types and shared constants (used by both the index and gsk backends) */
import { isBridgeableSocksUrl, startSocksBridge, stopSocksBridge } from './socks-bridge'

export interface WebSearchResult {
  title: string
  url: string
  snippet: string
}

export interface ImageSearchResult {
  title: string
  imageUrl: string
  sourceUrl: string
  source: string
  width?: number
  height?: number
}

// Known stock-photo hosts skipped during image search (matches the upstream filter list)
export const COPYRIGHT_HOSTS = ['gettyimages', 'istockphoto', 'shutterstock', 'corbis']

export function safeHost(url: unknown): string {
  try {
    return new URL(String(url)).hostname
  } catch {
    return ''
  }
}

/**
 * View untrusted JSON as a string-keyed record so properties can be probed
 * without `any`; non-object inputs read as an empty record.
 */
export function asRecord(v: unknown): Record<string, unknown> {
  return typeof v === 'object' && v !== null ? (v as Record<string, unknown>) : {}
}

/** First element when the value is an array, otherwise undefined (loose JSON probing). */
export function firstItem(v: unknown): unknown {
  return Array.isArray(v) ? (v as unknown[])[0] : undefined
}

let explicitProxyUrl = ''
/** loopback http:// bridge standing in for a SOCKS proxy, '' when not needed */
let bridgedProxyUrl = ''
/** resolves once the bridge for the current proxy is listening (or known not to be) */
let bridgeReady: Promise<void> = Promise.resolve()

/**
 * Proxy resolved by the apps' proxy bootstraps (env vars, else the system
 * proxy via session.resolveProxy); consumed by gskChildEnv() and the login
 * flow's proxy fallback.
 *
 * A SOCKS proxy additionally starts a loopback HTTP bridge, because the CLI
 * child can only be pointed at an http(s) proxy through its environment.
 */
export function setGskProxyUrl(url: string): void {
  if (url === explicitProxyUrl) return
  explicitProxyUrl = url
  bridgedProxyUrl = ''
  if (!isBridgeableSocksUrl(url)) {
    stopSocksBridge()
    bridgeReady = Promise.resolve()
    return
  }
  bridgeReady = startSocksBridge(url)
    .then((bridge) => {
      // a later change may have won the race; only claim the url we started for
      if (explicitProxyUrl === url) bridgedProxyUrl = bridge
    })
    .catch((err: unknown) => {
      console.warn('[proxy] could not start the SOCKS bridge for gsk children:', String(err))
    })
}

export function gskProxyUrl(): string {
  return explicitProxyUrl
}

/**
 * The proxy a spawned CLI child can actually use: the SOCKS bridge when one is
 * running, otherwise the configured proxy as-is.
 */
export function gskChildProxyUrl(): string {
  return bridgedProxyUrl || explicitProxyUrl
}

/** Await before spawning a child, so the first call cannot outrun the bridge. */
export function gskChildProxyReady(): Promise<void> {
  return bridgeReady
}
