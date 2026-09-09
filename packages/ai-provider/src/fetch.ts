/**
 * In Electron main processes AI requests run on Node's fetch (undici), which
 * connects directly instead of going through Chromium's network stack. Under
 * VPN/tun setups those direct connections can get reset (ECONNRESET) while
 * Chromium traffic — login, renderer fetches — works fine. Main processes
 * inject Electron's net.fetch here as a rescue path: when the primary fetch
 * fails at the network layer, the request is retried once over the Chromium
 * stack. Renderers never inject one (their fetch already is Chromium's).
 */

type FetchLike = (url: string, init?: RequestInit) => Promise<Response>

let rescueFetch: FetchLike | null = null

export function setRescueFetch(fn: FetchLike | null): void {
  rescueFetch = fn
}

/**
 * Node's fetch announces itself as a bare `node`, which gateways that watch
 * for anonymous automation treat as abusive traffic (OpenCode Go requires
 * clients to identify themselves and flags "broad" user agents). Main
 * processes refine the default with the app version.
 */
export const AI_DEFAULT_USER_AGENT = 'GenOffice'

let userAgent = AI_DEFAULT_USER_AGENT

export function setAiUserAgent(ua: string): void {
  userAgent = ua || AI_DEFAULT_USER_AGENT
}

/** protocols pass plain header records; keep that shape so callers can read the request back */
function withUserAgent(init: RequestInit): RequestInit {
  const given = init.headers
  const headers: Record<string, string> = {}
  if (given instanceof Headers) given.forEach((value, name) => (headers[name] = value))
  else if (Array.isArray(given)) for (const [name, value] of given) headers[name] = value
  else Object.assign(headers, given)
  if (!Object.keys(headers).some((name) => name.toLowerCase() === 'user-agent')) {
    headers['User-Agent'] = userAgent
  }
  return { ...init, headers }
}

/**
 * A 403 with an HTML body is an edge block page (bot management, WAF), not an API
 * answer. Those edges score the Node network stack's TLS fingerprint differently
 * from Chromium's, so the rescue fetch (Electron net.fetch) often gets through.
 */
function isBlockPage(response: Response): boolean {
  return (
    response.status === 403 && (response.headers.get('content-type') || '').includes('text/html')
  )
}

export async function aiFetch(url: string, rawInit: RequestInit): Promise<Response> {
  const init = withUserAgent(rawInit)
  const signal = init.signal as AbortSignal | null | undefined
  let response: Response
  try {
    response = await fetch(url, init)
  } catch (primaryError) {
    if (!rescueFetch || signal?.aborted) throw primaryError
    console.warn('[ai-provider] fetch failed, retrying via rescue fetch:', String(primaryError))
    try {
      return await rescueFetch(url, init)
    } catch {
      throw primaryError
    }
  }
  const resendable = init.body == null || typeof init.body === 'string'
  if (rescueFetch && resendable && !signal?.aborted && isBlockPage(response)) {
    console.warn(
      '[ai-provider] 403 block page from',
      new URL(url).host,
      '- retrying via rescue fetch',
    )
    try {
      const rescued = await rescueFetch(url, init)
      if (!isBlockPage(rescued)) return rescued
    } catch {
      // fall through to the primary response
    }
  }
  return response
}
