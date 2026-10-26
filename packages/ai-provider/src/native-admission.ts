/**
 * Optional signed edge admission (native-tok-v1) for Genspark LLM proxy requests whose final
 * credential is a gsk key. Off unless the host process sets GENOFFICE_NATIVE_WAF=1; renderers and
 * browsers have no process env and never opt in.
 *
 * Eligible: HTTPS to an admitted host (default proxy host, GSK_BASE_URL, GENOFFICE_NATIVE_WAF_HOSTS), path /api/llm_proxy/ or /api/anthropic (family
 * "llm"), and exactly one credential header carrying a gsk key — `Authorization: Bearer gsk-…`
 * (OpenAI-compatible) or `x-api-key: gsk-…` (Anthropic). BYOK keys, custom endpoints and every
 * other host are never sent to the issuer. Issuer failure leaves the request unchanged.
 */

const PASS_HEADER = 'X-Genspark-Token-Pass'

/**
 * The default proxy host (providers.ts always targets it), the configured service host and any
 * extra deployment hosts; no hostlist lives in source.
 */
export function nativeAdmissionHosts(env: NodeJS.ProcessEnv = process.env): Set<string> {
  const hosts = new Set(['www.genspark.ai'])
  try {
    if (env.GSK_BASE_URL) hosts.add(new URL(env.GSK_BASE_URL).hostname)
  } catch {
    // Malformed configuration only narrows the allowlist to the default host.
  }
  for (const host of (env.GENOFFICE_NATIVE_WAF_HOSTS ?? '').split(',')) {
    const trimmed = host.trim().toLowerCase()
    if (trimmed) hosts.add(trimmed)
  }
  return hosts
}

const PASS_SHAPE = /^\d{10}-[A-Za-z0-9_-]{43}$/
const MAX_BODY_BYTES = 4096
const MAX_ENTRIES = 32

type Fetcher = (url: string, init: RequestInit) => Promise<Response>
type Subject = { origin: string; name: 'authorization' | 'x-api-key'; value: string }
type Entry = { token: string; until: number; expires: number }
type Pending = { controller: AbortController; promise: Promise<string>; waiters: number }

export function nativeAdmissionEnabled(): boolean {
  return typeof process !== 'undefined' && process.env?.GENOFFICE_NATIVE_WAF === '1'
}

function subjectOf(raw: string, headers: Headers): Subject | undefined {
  let url: URL
  try {
    url = new URL(raw)
  } catch {
    return undefined
  }
  if (
    url.protocol !== 'https:' ||
    url.port ||
    url.username ||
    url.password ||
    !nativeAdmissionHosts().has(url.hostname)
  ) {
    return undefined
  }
  if (!url.pathname.startsWith('/api/llm_proxy/') && !url.pathname.startsWith('/api/anthropic')) {
    return undefined
  }
  const key = headers.get('x-api-key')
  const bearer = headers.get('authorization')
  if (key !== null && bearer !== null) return undefined
  const name = key !== null ? 'x-api-key' : 'authorization'
  const value = key ?? bearer
  if (
    !value ||
    !value.startsWith(name === 'x-api-key' ? 'gsk-' : 'Bearer gsk-') ||
    /[,\r\n]/.test(value)
  ) {
    return undefined
  }
  return { origin: url.origin, name, value }
}

async function bootstrap(
  subject: Subject,
  fetcher: Fetcher,
  signal: AbortSignal,
  now: () => number,
): Promise<string> {
  try {
    const response = await fetcher(`${subject.origin}/api/native-waf/token-pass?family=llm`, {
      method: 'POST',
      headers: { [subject.name]: subject.value },
      redirect: 'error',
      signal,
    })
    if (
      !response.ok ||
      !response.body ||
      Number(response.headers.get('content-length')) > MAX_BODY_BYTES
    ) {
      await response.body?.cancel()
      return ''
    }
    const reader = response.body.getReader()
    const chunks: Uint8Array[] = []
    let size = 0
    try {
      for (;;) {
        const part = await reader.read()
        if (part.done) break
        size += part.value.byteLength
        if (size > MAX_BODY_BYTES) return ''
        chunks.push(part.value)
      }
    } finally {
      await reader.cancel()
      reader.releaseLock()
    }
    const bytes = new Uint8Array(size)
    let offset = 0
    for (const chunk of chunks) {
      bytes.set(chunk, offset)
      offset += chunk.byteLength
    }
    const data = (
      JSON.parse(new TextDecoder().decode(bytes)) as {
        data?: { token?: unknown; family?: unknown }
      }
    ).data
    const token = data?.token
    if (data?.family !== 'llm' || typeof token !== 'string' || !PASS_SHAPE.test(token)) return ''
    const issued = Number(token.slice(0, 10)) * 1000
    return issued <= now() + 30_000 && issued + 270_000 > now() ? token : ''
  } catch {
    // Admission is optional; the business request keeps its own retry policy.
    return ''
  }
}

/** One bounded, memory-only cache shared by every provider request of this process. */
export function createNativeAdmission(
  options: { bootstrapFetch?: Fetcher; now?: () => number; timeoutMs?: number } = {},
) {
  const cache = new Map<string, Entry>()
  const pending = new Map<string, Pending>()
  const now = options.now ?? Date.now

  function start(key: string, subject: Subject, fetcher: Fetcher): Pending {
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), options.timeoutMs ?? 2000)
    const promise = bootstrap(subject, fetcher, controller.signal, now)
      .then((token) => {
        const expires = token
          ? Math.min(now() + 240_000, Number(token.slice(0, 10)) * 1000 + 270_000)
          : now() + 15_000
        if (pending.get(key)?.controller === controller) {
          if (cache.size >= MAX_ENTRIES) cache.delete(cache.keys().next().value!)
          cache.set(key, { token, expires, until: performance.now() + expires - now() })
        }
        return token
      })
      .finally(() => {
        clearTimeout(timer)
        if (pending.get(key)?.controller === controller) pending.delete(key)
      })
    const work = { controller, promise, waiters: 0 }
    pending.set(key, work)
    return work
  }

  async function pass(
    url: string,
    headers: Headers,
    fetcher: Fetcher,
    signal?: AbortSignal | null,
  ): Promise<string> {
    signal?.throwIfAborted()
    const subject = subjectOf(url, headers)
    if (!subject) return ''
    const key = `${subject.origin}\n${subject.name}\n${subject.value}`
    const entry = cache.get(key)
    if (entry && entry.until > performance.now() && entry.expires > now()) return entry.token
    const work = pending.get(key) ?? start(key, subject, options.bootstrapFetch ?? fetcher)
    work.waiters++
    // Each waiter owns its cancellation; only the last one to leave aborts the bootstrap.
    let onAbort: () => void = () => {}
    const cancelled = new Promise<never>((_, reject) => {
      onAbort = () => reject(signal?.reason)
    })
    signal?.addEventListener('abort', onAbort, { once: true })
    try {
      const token = await Promise.race([work.promise, cancelled])
      signal?.throwIfAborted()
      return token
    } finally {
      signal?.removeEventListener('abort', onAbort)
      work.waiters--
      if (!work.waiters && pending.get(key) === work) {
        pending.delete(key)
        work.controller.abort()
      }
    }
  }

  /** Returns the init to send: unchanged unless a pass was obtained. */
  async function prepare(url: string, init: RequestInit, fetcher: Fetcher): Promise<RequestInit> {
    if (!nativeAdmissionEnabled()) return init
    const headers = new Headers(init.headers)
    const token = await pass(url, headers, fetcher, init.signal)
    if (!token) return init
    const record: Record<string, string> = {}
    headers.forEach((value, name) => (record[name] = value))
    record[PASS_HEADER] = token
    // A request carrying a pass must never follow a redirect to another origin.
    return { ...init, headers: record, redirect: 'error' }
  }

  return { pass, prepare }
}

export const nativeAdmission = createNativeAdmission()
