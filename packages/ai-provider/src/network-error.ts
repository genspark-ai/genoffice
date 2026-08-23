/**
 * Classifies transport-level connectivity failures (DNS, refused/reset
 * connections, dead proxies/VPNs…) so the apps can show a localized
 * "network problem, check your connection" message (errorCode 'network')
 * instead of the raw `fetch failed cause=ECONNRESET` text.
 */

// Node/undici error codes plus Chromium net-stack error names (net.fetch rescue path)
const NETWORK_ERROR_PATTERN = new RegExp(
  [
    'fetch failed',
    'socket hang up',
    'network error',
    'ECONNRESET',
    'ECONNREFUSED',
    'ECONNABORTED',
    'ETIMEDOUT',
    'ENOTFOUND',
    'EAI_AGAIN',
    'ENETUNREACH',
    'ENETDOWN',
    'EHOSTUNREACH',
    'EHOSTDOWN',
    'EPIPE',
    'UND_ERR_CONNECT_TIMEOUT',
    'UND_ERR_SOCKET',
    'ERR_NETWORK',
    'ERR_INTERNET_DISCONNECTED',
    'ERR_NAME_NOT_RESOLVED',
    'ERR_CONNECTION_[A-Z_]+',
    'ERR_PROXY_CONNECTION_FAILED',
    'ERR_TIMED_OUT',
    'ERR_ADDRESS_UNREACHABLE',
  ].join('|'),
  'i',
)

/**
 * True when the error (or anything in its `cause` chain) looks like a network
 * connectivity failure. Provider wrappers embed the cause code in the message
 * (e.g. "Claude fetch failed: … cause=ECONNRESET"), so the message text is
 * checked as well as structured `code` fields.
 */
export function isAiNetworkError(err: unknown): boolean {
  let current: unknown = err
  for (let depth = 0; current && depth < 5; depth++) {
    const e = current as { message?: unknown; code?: unknown; cause?: unknown }
    if (typeof e.code === 'string' && NETWORK_ERROR_PATTERN.test(e.code)) return true
    if (typeof e.message === 'string' && NETWORK_ERROR_PATTERN.test(e.message)) return true
    if (typeof current === 'string' && NETWORK_ERROR_PATTERN.test(current)) return true
    current = e.cause
  }
  return false
}
