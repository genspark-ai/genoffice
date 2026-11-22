/**
 * Local-only fork: which network destinations the app may reach.
 *
 * Loopback only. Anything else (public hosts, LAN addresses, bare intranet
 * names, 0.0.0.0) is refused. See docs/local-only-fork.md.
 */

const LOOPBACK_NAMES = new Set(['localhost', '127.0.0.1', '::1', '::ffff:127.0.0.1'])

/** Schemes that open a network connection; every other scheme is local. */
const NETWORK_SCHEMES = new Set(['http:', 'https:', 'ws:', 'wss:', 'ftp:'])

export function isLoopbackHost(host: string): boolean {
  let h = host.trim().toLowerCase()
  if (h.startsWith('[') && h.endsWith(']')) h = h.slice(1, -1)
  if (h.endsWith('.')) h = h.slice(0, -1)
  if (LOOPBACK_NAMES.has(h)) return true
  // the whole 127.0.0.0/8 block is loopback
  return /^127\.\d{1,3}\.\d{1,3}\.\d{1,3}$/.test(h)
}

export type EgressDecision = { allowed: true } | { allowed: false; host: string }

/**
 * Decide a URL. Non-network schemes (file:, data:, blob:, devtools:, the
 * app's own protocols) are allowed; network schemes only to loopback.
 * An unparseable URL is allowed here: the caller's own parser rejects it.
 */
export function egressDecision(raw: string): EgressDecision {
  let url: URL
  try {
    url = new URL(raw)
  } catch {
    return { allowed: true }
  }
  if (!NETWORK_SCHEMES.has(url.protocol)) return { allowed: true }
  return isLoopbackHost(url.hostname) ? { allowed: true } : { allowed: false, host: url.hostname }
}

/** True when opening this URL in the user's browser should be confirmed first. */
export function externalOpenNeedsConfirm(raw: string): boolean {
  let url: URL
  try {
    url = new URL(raw)
  } catch {
    return false
  }
  return (url.protocol === 'http:' || url.protocol === 'https:') && !isLoopbackHost(url.hostname)
}
