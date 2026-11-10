type Fetcher = (url: string, init: RequestInit) => Promise<Response>

/** The configured service host plus any extra deployment hosts; no hostlist lives in source. */
export function nativeWafHosts(env: NodeJS.ProcessEnv = process.env): Set<string> {
  const hosts = new Set<string>()
  try {
    hosts.add(new URL(env.GSK_BASE_URL || 'https://www.genspark.ai').hostname)
  } catch {
    hosts.add('www.genspark.ai')
  }
  for (const host of (env.GENOFFICE_NATIVE_WAF_HOSTS ?? '').split(',')) {
    const trimmed = host.trim().toLowerCase()
    if (trimmed) hosts.add(trimmed)
  }
  return hosts
}

export class NativeWaf {
  private generation = 0
  private subject = ''
  private token: string | undefined
  private until = 0
  private pending: Promise<string | undefined> | undefined

  clear(): void {
    this.generation++
    this.subject = ''
    this.token = undefined
    this.pending = undefined
    this.until = 0
  }

  async prepare(url: string, init: RequestInit, fetcher: Fetcher): Promise<RequestInit> {
    const headers = new Headers(init.headers)
    headers.delete('X-Genspark-Native-Pass')
    const target = new URL(url)
    const cookies = (headers.get('Cookie') ?? '')
      .split(';')
      .map((part) => part.trim())
      .filter((part) => part.startsWith('session_id='))
    if (
      target.protocol !== 'https:' ||
      target.port ||
      target.username ||
      target.password ||
      !nativeWafHosts().has(target.hostname) ||
      !target.pathname.startsWith('/api/') ||
      target.pathname === '/api/native-waf/pass' ||
      cookies.length !== 1 ||
      !/^session_id=[A-Za-z0-9_.:-]+$/.test(cookies[0]!)
    ) {
      return { ...init, headers }
    }
    init.signal?.throwIfAborted()
    const cookie = cookies[0]!
    const subject = `${target.origin}|${cookie}`
    if (this.subject !== subject) {
      this.clear()
      this.subject = subject
    }
    const generation = this.generation
    let pass = this.token
    if (Date.now() >= this.until) {
      if (!this.pending) {
        const pending = this.load(target.origin, cookie, headers, fetcher).then((token) => {
          if (generation === this.generation && this.pending === pending) {
            this.token = token
            this.until = Date.now() + (token ? Math.min(240_000, this.remaining(token)) : 15_000)
            this.pending = undefined
          }
          return token
        })
        this.pending = pending
      }
      pass = await this.wait(this.pending, init.signal)
    }
    init.signal?.throwIfAborted()
    if (generation === this.generation && pass && this.remaining(pass) > 0) {
      headers.set('X-Genspark-Native-Pass', pass)
    }
    // Explicit cookies and passes may never ride a redirect or an ambient Chromium cookie jar.
    return { ...init, headers, credentials: 'omit', redirect: 'manual' }
  }

  private remaining(token: string): number {
    const issued = Number(token.split('-')[0]) * 1000
    const age = Math.max(0, Date.now() - issued)
    return age >= 0 && age < 300_000 ? 300_000 - age : 0
  }

  private async wait<T>(pending: Promise<T>, signal?: AbortSignal | null): Promise<T> {
    if (!signal) return pending
    signal.throwIfAborted()
    return new Promise((resolve, reject) => {
      const abort = (): void => reject(signal.reason)
      signal.addEventListener('abort', abort, { once: true })
      void pending.then(resolve, reject).finally(() => signal.removeEventListener('abort', abort))
    })
  }

  private async load(
    origin: string,
    cookie: string,
    original: Headers,
    fetcher: Fetcher,
  ): Promise<string | undefined> {
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), 5000)
    try {
      const headers = new Headers({ Cookie: cookie })
      original.forEach((value, name) => {
        if (name === 'user-agent' || name === 'x-route-identifier' || name.startsWith('gs-')) {
          headers.set(name, value)
        }
      })
      const response = await fetcher(`${origin}/api/native-waf/pass`, {
        method: 'POST',
        headers,
        signal: controller.signal,
        redirect: 'manual',
        credentials: 'omit',
      })
      if (response.status !== 200 || !response.body) {
        await response.body?.cancel()
        return undefined
      }
      const reader = response.body.getReader()
      const decoder = new TextDecoder()
      let size = 0
      let text = ''
      try {
        for (;;) {
          const chunk = await reader.read()
          if (chunk.done) break
          size += chunk.value.byteLength
          if (size > 4096) return undefined
          text += decoder.decode(chunk.value, { stream: true })
        }
        text += decoder.decode()
      } finally {
        await reader.cancel()
      }
      const envelope: unknown = JSON.parse(text)
      if (
        typeof envelope !== 'object' ||
        envelope === null ||
        !('data' in envelope) ||
        typeof envelope.data !== 'object' ||
        envelope.data === null ||
        !('token' in envelope.data)
      )
        return undefined
      const token = envelope.data.token
      return typeof token === 'string' &&
        /^\d{10}-[A-Za-z0-9_-]{43}$/.test(token) &&
        this.remaining(token) > 0
        ? token
        : undefined
    } catch {
      return undefined
    } finally {
      clearTimeout(timer)
    }
  }
}

export const nativeWaf = new NativeWaf()
