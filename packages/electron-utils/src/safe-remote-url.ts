/// SSRF gate for main-process fetches whose target is influenced by untrusted
/// input — most importantly AI tool calls, where a poisoned web-search result
/// can steer the model into requesting an internal address. Every hop of a
/// redirect chain has to pass, because validating only the initial URL lets a
/// public host bounce the request to loopback or a cloud metadata endpoint.

import { lookup } from 'node:dns/promises'
import { BlockList, isIP } from 'node:net'

/**
 * Non-public address ranges. BlockList does the range math and maps
 * ::ffff:0:0/96 addresses back onto the IPv4 rules, which hand-rolled string
 * checks get wrong for normalized forms such as [::ffff:127.0.0.1], whose
 * hostname becomes ::ffff:7f00:1.
 */
const blockedAddresses = (() => {
  const list = new BlockList()
  list.addSubnet('0.0.0.0', 8) // "this network"
  list.addSubnet('10.0.0.0', 8) // private
  list.addSubnet('100.64.0.0', 10) // carrier NAT
  list.addSubnet('127.0.0.0', 8) // loopback
  list.addSubnet('169.254.0.0', 16) // link-local (cloud metadata)
  list.addSubnet('172.16.0.0', 12) // private
  list.addSubnet('192.0.0.0', 24) // IETF protocol assignments
  list.addSubnet('192.168.0.0', 16) // private
  list.addSubnet('198.18.0.0', 15) // benchmarking
  list.addSubnet('224.0.0.0', 4) // multicast
  list.addSubnet('240.0.0.0', 4) // reserved + broadcast
  list.addAddress('::', 'ipv6') // unspecified
  list.addAddress('::1', 'ipv6') // loopback
  list.addSubnet('::', 96, 'ipv6') // deprecated IPv4-compatible (::127.0.0.1)
  list.addSubnet('64:ff9b::', 96, 'ipv6') // NAT64
  list.addSubnet('fc00::', 7, 'ipv6') // unique local
  list.addSubnet('fe80::', 10, 'ipv6') // link-local
  list.addSubnet('ff00::', 8, 'ipv6') // multicast
  return list
})()

/** Hostname suffixes that never denote a public host. */
const BLOCKED_HOST_SUFFIXES: readonly string[] = ['.localhost', '.local', '.internal']

/**
 * The same 198.18.0.0/15 range on its own list: VPN clients in fake-ip mode
 * answer every DNS query with a placeholder address from here (the real
 * resolution happens at the proxy exit), so a hostname that resolves into this
 * range is a proxy artifact, not an intranet target. Nothing else deploys the
 * range (RFC 2544 benchmarking), and a connection to it either rides the VPN
 * tunnel outward — never into the LAN — or simply fails to connect.
 */
const fakeIpRange = (() => {
  const list = new BlockList()
  list.addSubnet('198.18.0.0', 15)
  return list
})()

/** True when `ip` is a fake-ip placeholder a VPN resolver handed out for a real hostname. */
export function isFakeIpDnsArtifact(ip: string): boolean {
  return isIP(ip) === 4 && fakeIpRange.check(ip, 'ipv4')
}

/** True when `ip` is a literal address outside the public ranges (non-addresses count as unsafe). */
export function isBlockedAddress(ip: string): boolean {
  const family = isIP(ip)
  if (family === 0) return true
  return blockedAddresses.check(ip, family === 4 ? 'ipv4' : 'ipv6')
}

/**
 * True when the URL is http(s) and its host is public. Literal addresses are
 * checked directly — a URL naming an internal address outright is refused with
 * no exception. Hostnames are resolved and every returned address must pass,
 * so a name pointing at an internal address is rejected; the single exception
 * is a resolved address in the VPN fake-ip range (see `isFakeIpDnsArtifact`),
 * which only a proxy resolver can produce for a hostname an attacker does not
 * already control the machine's DNS for.
 *
 * This is a best-effort control, not a guarantee: DNS can change between this
 * check and the request (rebinding). Pair it with per-hop revalidation via
 * `fetchWithSsrfGuard` rather than relying on a single up-front check.
 */
export async function isSafeRemoteUrl(raw: unknown): Promise<boolean> {
  if (typeof raw !== 'string') return false
  let url: URL
  try {
    url = new URL(raw)
  } catch {
    return false
  }
  if (url.protocol !== 'http:' && url.protocol !== 'https:') return false
  // URL keeps IPv6 literals bracketed; isIP does not accept the brackets.
  // Strip a trailing dot: DNS treats "localhost." as "localhost", so the
  // hostname gate must too (otherwise the check passes and only the later
  // DNS lookup catches it, after a needless external resolution).
  const host = url.hostname
    .replace(/^\[|\]$/g, '')
    .toLowerCase()
    .replace(/\.+$/, '')
  if (host === '') return false
  if (isIP(host)) return !isBlockedAddress(host)
  if (host === 'localhost' || BLOCKED_HOST_SUFFIXES.some((s) => host.endsWith(s))) return false
  try {
    const addrs = await lookup(host, { all: true })
    return (
      addrs.length > 0 &&
      addrs.every((a) => !isBlockedAddress(a.address) || isFakeIpDnsArtifact(a.address))
    )
  } catch {
    return false
  }
}

export interface FetchWithSsrfGuardOptions {
  /** Redirect hops to follow. Each hop is revalidated. Defaults to 5. */
  maxRedirects?: number
  headers?: Record<string, string>
  /** Injectable for tests; defaults to global fetch. */
  fetchImpl?: typeof fetch
}

/**
 * fetch() that validates the target before every hop, following redirects
 * manually. Returns null when any hop fails validation, a redirect lacks a
 * usable Location, or the hop budget is exhausted.
 */
export async function fetchWithSsrfGuard(
  rawUrl: string,
  options: FetchWithSsrfGuardOptions = {},
): Promise<Response | null> {
  const { maxRedirects = 5, headers, fetchImpl = fetch } = options
  // maxRedirects arrives from callers (CLI flags, IPC-adjacent config): an
  // Infinity value would follow redirects forever, so normalize to 0..10.
  const hopBudget = Number.isFinite(maxRedirects)
    ? Math.min(Math.max(0, Math.floor(maxRedirects)), 10)
    : 5
  let current = rawUrl
  for (let hop = 0; hop <= hopBudget; hop++) {
    if (!(await isSafeRemoteUrl(current))) return null
    // headers stays absent rather than explicitly undefined (exactOptionalPropertyTypes)
    const resp = await fetchImpl(current, {
      ...(headers ? { headers } : {}),
      redirect: 'manual',
    })
    if (resp.status < 300 || resp.status >= 400) return resp
    const location = resp.headers.get('location')
    if (!location) return null
    try {
      current = new URL(location, current).toString()
    } catch {
      return null
    }
  }
  return null
}
