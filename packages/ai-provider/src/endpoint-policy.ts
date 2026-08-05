import type { AiEndpointKind, AiEndpointPolicy, AiEndpointValidation } from './types'

/** Defaults are suitable for user-configured providers: local model servers are allowed,
 * while plaintext public endpoints and URL-embedded credentials are rejected. */
export const DEFAULT_AI_ENDPOINT_POLICY: AiEndpointPolicy = {
  allowLocal: true,
  allowInsecureHttp: false,
  allowUrlCredentials: false,
}

function unbracket(hostname: string): string {
  return hostname.startsWith('[') && hostname.endsWith(']') ? hostname.slice(1, -1) : hostname
}

function isPrivateIpv4(hostname: string): boolean {
  const parts = hostname.split('.')
  if (parts.length !== 4 || parts.some((part) => !/^\d+$/.test(part))) return false
  const octets = parts.map(Number)
  if (octets.some((octet) => octet < 0 || octet > 255)) return false
  const a = octets[0] ?? -1
  const b = octets[1] ?? -1
  return (
    a === 0 ||
    a === 10 ||
    a === 127 ||
    (a === 100 && b >= 64 && b <= 127) ||
    (a === 169 && b === 254) ||
    (a === 172 && b >= 16 && b <= 31) ||
    (a === 192 && b === 168)
  )
}

function isPrivateIpv6(hostname: string): boolean {
  const normalized = hostname.toLowerCase()
  if (normalized.startsWith('::ffff:')) return isPrivateIpv4(normalized.slice('::ffff:'.length))
  return (
    normalized === '::1' ||
    normalized === '::' ||
    normalized.startsWith('fc') ||
    normalized.startsWith('fd') ||
    normalized.startsWith('fe8') ||
    normalized.startsWith('fe9') ||
    normalized.startsWith('fea') ||
    normalized.startsWith('feb') ||
    normalized.startsWith('::ffff:127.')
  )
}

export function isLocalEndpointHost(hostname: string): boolean {
  const host = unbracket(hostname.trim().toLowerCase()).replace(/\.$/, '')
  return (
    host === 'localhost' ||
    host.endsWith('.localhost') ||
    host.endsWith('.local') ||
    isPrivateIpv4(host) ||
    (host.includes(':') && isPrivateIpv6(host))
  )
}

export function endpointKindForUrl(endpoint: string): AiEndpointKind {
  try {
    const url = new URL(endpoint)
    return isLocalEndpointHost(url.hostname) ? 'local' : 'cloud'
  } catch {
    return 'custom'
  }
}

/** Validate and normalize a provider base URL before any network request. */
export function validateAiEndpoint(
  endpoint: string,
  policy: AiEndpointPolicy = DEFAULT_AI_ENDPOINT_POLICY,
): AiEndpointValidation {
  const raw = endpoint.trim()
  if (!raw) return { ok: false, reason: 'Endpoint URL is required' }

  let url: URL
  try {
    url = new URL(raw)
  } catch {
    return { ok: false, reason: 'Endpoint must be a valid URL' }
  }
  if (url.protocol !== 'https:' && url.protocol !== 'http:') {
    return { ok: false, reason: 'Endpoint must use http:// or https://' }
  }
  if (!url.hostname) return { ok: false, reason: 'Endpoint hostname is required' }

  const local = isLocalEndpointHost(url.hostname)
  if (local && !policy.allowLocal) {
    return { ok: false, reason: 'Local network endpoints are disabled by policy' }
  }
  if (url.protocol === 'http:' && !local && !policy.allowInsecureHttp) {
    return { ok: false, reason: 'Public endpoints must use HTTPS' }
  }
  if ((url.username || url.password) && !policy.allowUrlCredentials) {
    return { ok: false, reason: 'Credentials in endpoint URLs are not allowed' }
  }
  if ((url.search || url.hash) && !policy.allowUrlCredentials) {
    return { ok: false, reason: 'Query strings and fragments are not allowed in endpoint URLs' }
  }

  // URL.pathname keeps provider prefixes such as `/v1`; only remove trailing slashes.
  url.pathname = url.pathname.replace(/\/+$/, '') || '/'
  url.search = ''
  url.hash = ''
  return {
    ok: true,
    normalized: url.toString().replace(/\/$/, ''),
    kind: local ? 'local' : 'cloud',
  }
}

export function assertValidAiEndpoint(
  endpoint: string,
  policy: AiEndpointPolicy = DEFAULT_AI_ENDPOINT_POLICY,
): string {
  const result = validateAiEndpoint(endpoint, policy)
  if (!result.ok || !result.normalized) throw new Error(result.reason ?? 'Invalid endpoint URL')
  return result.normalized
}
