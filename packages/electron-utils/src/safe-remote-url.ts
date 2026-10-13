/// SSRF gate for main-process fetches whose target is influenced by untrusted
/// input — most importantly AI tool calls, where a poisoned web-search result
/// can steer the model into requesting an internal address. Every hop of a
/// redirect chain has to pass, because validating only the initial URL lets a
/// public host bounce the request to loopback or a cloud metadata endpoint.
///
/// Validation and connection are bound: the gate resolves the hostname once,
/// validates those addresses, and pins them into the connection through a
/// per-hop undici Agent whose `lookup` hands back exactly the validated list.
/// A separate resolution at connect time — the DNS rebinding window every
/// check-then-connect guard otherwise carries — never happens.

import { lookup } from 'node:dns/promises'
import type { LookupAddress, LookupOptions } from 'node:dns'
import { BlockList, isIP } from 'node:net'
import { Agent, fetch as undiciFetch, getGlobalDispatcher, type Dispatcher } from 'undici'

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

/** Hostname suffixes that never denote a public host. */
const BLOCKED_HOST_SUFFIXES: readonly string[] = ['.localhost', '.local', '.internal']

/** One resolved, validated connection target. */
export interface RemoteTargetAddress {
  address: string
  family: 4 | 6
}

export interface SafeRemoteTarget {
  /** The URL the addresses were validated for (fetch this exact string). */
  url: string
  /** Normalized hostname (no brackets, no trailing dot) the addresses answer for. */
  hostname: string
  addresses: RemoteTargetAddress[]
  /** The URL carried an address literal; no resolution happened, nothing to pin. */
  literal: boolean
}

const normalizeHost = (hostname: string): string =>
  hostname
    .replace(/^\[|\]$/g, '')
    .toLowerCase()
    .replace(/\.+$/, '')

/**
 * Resolves and validates `raw`: http(s) only, a public literal, or a hostname
 * whose every resolved address is public. Literal addresses are checked
 * directly — a URL naming an internal address outright is refused with no
 * exception. The single allowance on the resolved side is an address in the
 * VPN fake-ip range (see `isFakeIpDnsArtifact`), which only a proxy resolver
 * can produce for a hostname an attacker does not already control the
 * machine's DNS for.
 *
 * Returns the validated addresses so the caller can bind the connection to
 * them (see `fetchWithSsrfGuard`); a plain boolean check loses exactly the
 * information that makes rebinding impossible.
 */
export async function resolveSafeRemoteTarget(raw: unknown): Promise<SafeRemoteTarget | null> {
  if (typeof raw !== 'string') return null
  let url: URL
  try {
    url = new URL(raw)
  } catch {
    return null
  }
  if (url.protocol !== 'http:' && url.protocol !== 'https:') return null
  // URL keeps IPv6 literals bracketed; isIP does not accept the brackets.
  // Strip a trailing dot: DNS treats "localhost." as "localhost", so the
  // hostname gate must too (otherwise the check passes and only the later
  // DNS lookup catches it, after a needless external resolution).
  const host = normalizeHost(url.hostname)
  if (host === '') return null
  const literalFamily = isIP(host)
  if (literalFamily !== 0) {
    if (isBlockedAddress(host)) return null
    return {
      url: url.toString(),
      hostname: host,
      literal: true,
      addresses: [{ address: host, family: literalFamily === 4 ? 4 : 6 }],
    }
  }
  if (host === 'localhost' || BLOCKED_HOST_SUFFIXES.some((s) => host.endsWith(s))) return null
  try {
    const addrs = await lookup(host, { all: true })
    if (
      addrs.length === 0 ||
      !addrs.every((a) => !isBlockedAddress(a.address) || isFakeIpDnsArtifact(a.address))
    ) {
      return null
    }
    return {
      url: url.toString(),
      hostname: host,
      literal: false,
      addresses: addrs.map((a) => ({
        address: a.address,
        family: a.family === 6 ? 6 : 4,
      })),
    }
  } catch {
    return null
  }
}

/**
 * True when the URL is http(s) and its host is public (see
 * `resolveSafeRemoteTarget` for the exact rules, including the fake-ip
 * allowance for resolved hostnames).
 */
export async function isSafeRemoteUrl(raw: unknown): Promise<boolean> {
  return (await resolveSafeRemoteTarget(raw)) !== null
}

type PinnedLookup = (
  hostname: string,
  options: LookupOptions,
  callback: (
    err: NodeJS.ErrnoException | null,
    address: string | LookupAddress[],
    family?: number,
  ) => void,
) => void

/**
 * A connect-time `lookup` that ignores the OS resolver and answers with the
 * addresses the gate validated for this hostname — the binding half of the
 * guard. Any other hostname fails closed: the dispatcher is built per hop and
 * must never be asked about a host that was not validated.
 */
export function pinnedLookup(target: SafeRemoteTarget): PinnedLookup {
  return (hostname, options, callback) => {
    if (normalizeHost(hostname) !== target.hostname) {
      const err: NodeJS.ErrnoException = new Error(
        `ssrf guard: refusing to connect to unvalidated host ${hostname}`,
      )
      err.code = 'ESSRFHOST'
      callback(err, '')
      return
    }
    // net.connect only passes all:true on the happy-eyeballs path; the
    // single-address path expects (err, address, family) and would reject an
    // array with ERR_INVALID_IP_ADDRESS.
    if (options.all) {
      callback(null, target.addresses)
      return
    }
    const pick =
      (options.family !== undefined
        ? target.addresses.find((a) => a.family === options.family)
        : undefined) ?? target.addresses[0]
    if (!pick) {
      // unreachable: resolveSafeRemoteTarget never returns an empty list, but
      // a pinned lookup must never answer "connect anywhere" instead
      const err: NodeJS.ErrnoException = new Error('ssrf guard: no validated address to pin')
      err.code = 'ESSRFHOST'
      callback(err, '')
      return
    }
    callback(null, pick.address, pick.family)
  }
}

export interface FetchWithSsrfGuardOptions {
  /** Redirect hops to follow. Each hop is revalidated. Defaults to 5. */
  maxRedirects?: number
  headers?: Record<string, string>
  /** Injectable for tests; defaults to undici's fetch (which honors the pinned dispatcher). */
  fetchImpl?: typeof fetch
}

/**
 * Own symbols that only undici's proxy dispatchers carry (every copy of the
 * library, npm or Node's bundled one, derives them from the same source):
 * a plain Agent — the default global dispatcher, and the one Node's bundled
 * copy wraps for its own slot — has none of these. Detection is duck-typed
 * on symbol descriptions rather than `instanceof`, because the global
 * dispatcher slot is shared through a registered symbol across all copies,
 * so the object in it may belong to a different copy's class hierarchy.
 */
const PROXY_DISPATCHER_SYMBOLS: ReadonlySet<string> = new Set([
  'proxy agent options',
  'proxy client',
  'proxy agent',
  'proxy headers',
  'tunnel proxy',
])

function globalDispatcherRoutesThroughProxy(): boolean {
  const d: unknown = getGlobalDispatcher()
  if (d === null || typeof d !== 'object') return false
  return Object.getOwnPropertySymbols(d).some((s) =>
    PROXY_DISPATCHER_SYMBOLS.has(s.description ?? ''),
  )
}

/**
 * fetch() that validates the target before every hop and binds the connection
 * to the validated addresses, following redirects manually. Returns null when
 * any hop fails validation, a redirect lacks a usable Location, or the hop
 * budget is exhausted.
 *
 * Hostname hops connect through a per-hop undici Agent whose lookup answers
 * with the validated address list (TLS keeps the URL hostname for SNI and
 * certificate verification); literal-IP URLs need no pinning. When a forward
 * proxy is installed globally, the hop rides it instead — the proxy resolves,
 * so pinning is impossible — and the guard degrades to per-hop validation.
 * A `fetchImpl` injected for tests receives the dispatcher in the init but
 * owns its transport, so binding is a production-path guarantee.
 */
export async function fetchWithSsrfGuard(
  rawUrl: string,
  options: FetchWithSsrfGuardOptions = {},
): Promise<Response | null> {
  const { maxRedirects = 5, headers, fetchImpl } = options
  // maxRedirects arrives from callers (CLI flags, IPC-adjacent config): an
  // Infinity value would follow redirects forever, so normalize to 0..10.
  const hopBudget = Number.isFinite(maxRedirects)
    ? Math.min(Math.max(0, Math.floor(maxRedirects)), 10)
    : 5
  let current = rawUrl
  for (let hop = 0; hop <= hopBudget; hop++) {
    const target = await resolveSafeRemoteTarget(current)
    if (!target) return null
    // headers stays absent rather than explicitly undefined (exactOptionalPropertyTypes)
    const init: Record<string, unknown> = { ...(headers ? { headers } : {}), redirect: 'manual' }
    let dispatcher: Dispatcher | undefined
    if (!target.literal && !globalDispatcherRoutesThroughProxy()) {
      // A forward proxy resolves at its own exit, so pinning is neither
      // possible (the proxy owns the connection) nor meaningful — when the
      // shell has installed one via setGlobalDispatcher, let the hop ride it
      // instead of dialing the validated address directly and breaking
      // proxied users. Behind a proxy the gate therefore degrades to its
      // per-hop validation; rebinding is then the proxy's to prevent.
      dispatcher = new Agent({ connect: { lookup: pinnedLookup(target) } })
      init.dispatcher = dispatcher
    }
    try {
      const resp =
        fetchImpl !== undefined
          ? await fetchImpl(current, init as RequestInit)
          : ((await undiciFetch(current, init as never)) as unknown as Response)
      if (resp.status < 300 || resp.status >= 400) return resp
      const location = resp.headers.get('location')
      if (!location) return null
      // The redirect hop's body is never read; cancel it so the socket — a
      // keep-alive global-dispatcher socket especially — returns to its pool
      // now instead of lingering until the keep-alive timeout.
      try {
        await resp.body?.cancel()
      } catch {
        // a locked or already-consumed body has nowhere to drain to
      }
      try {
        current = new URL(location, current).toString()
      } catch {
        return null
      }
    } finally {
      // graceful: in-flight bodies keep streaming, only new dispatches are refused
      void dispatcher?.close()
    }
  }
  return null
}
