import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { aiFetch } from '../src/fetch'
import { createNativeAdmission, nativeAdmissionHosts } from '../src/native-admission'

const OPENAI = 'https://www.genspark.ai/api/llm_proxy/v1/chat/completions'
const ANTHROPIC = 'https://www.genspark.ai/api/anthropic/v1/messages'
const pass = (seconds = Math.floor(Date.now() / 1000)) => `${seconds}-${'A'.repeat(43)}`
const issued = (token: string, family = 'llm') => Response.json({ data: { token, family } })

beforeEach(() => {
  process.env.GENOFFICE_NATIVE_WAF = '1'
})

afterEach(() => {
  delete process.env.GENOFFICE_NATIVE_WAF
  vi.unstubAllGlobals()
})

function recorder(reply: (url: string, init: RequestInit) => Response | Promise<Response>) {
  const calls: { url: string; init: RequestInit }[] = []
  const fetcher = vi.fn(async (url: string, init: RequestInit) => {
    calls.push({ url, init })
    return reply(url, init)
  })
  return { calls, fetcher }
}

describe('native admission for Genspark provider requests', () => {
  it('is off by default: no issuer call and the init passes through untouched', async () => {
    delete process.env.GENOFFICE_NATIVE_WAF
    const { fetcher } = recorder(() => issued(pass()))
    const init = { headers: { Authorization: 'Bearer gsk-k' } }
    expect(await createNativeAdmission().prepare(OPENAI, init, fetcher)).toBe(init)
    expect(fetcher).not.toHaveBeenCalled()
  })

  it('binds the exact credential header for both protocols and caches per subject', async () => {
    const token = pass()
    const { calls, fetcher } = recorder(() => issued(token))
    const admission = createNativeAdmission()
    const openai = await admission.prepare(
      OPENAI,
      { headers: { Authorization: 'Bearer gsk-k' } },
      fetcher,
    )
    await admission.prepare(OPENAI, { headers: { Authorization: 'Bearer gsk-k' } }, fetcher)
    const anthropic = await admission.prepare(
      ANTHROPIC,
      { headers: { 'x-api-key': 'gsk-k' } },
      fetcher,
    )
    expect(calls).toHaveLength(2)
    expect(calls[0]!.url).toBe('https://www.genspark.ai/api/native-waf/token-pass?family=llm')
    expect(new Headers(calls[0]!.init.headers).get('authorization')).toBe('Bearer gsk-k')
    expect(new Headers(calls[1]!.init.headers).get('x-api-key')).toBe('gsk-k')
    expect(calls[0]!.init.redirect).toBe('error')
    for (const init of [openai, anthropic]) {
      expect(new Headers(init.headers).get('x-genspark-token-pass')).toBe(token)
      expect(init.redirect).toBe('error')
    }
    expect(new Headers(openai.headers).get('authorization')).toBe('Bearer gsk-k')
  })

  it('never sends BYOK keys, other hosts or mixed credentials to the issuer', async () => {
    const { fetcher } = recorder(() => {
      throw new Error('must not bootstrap')
    })
    const admission = createNativeAdmission()
    const cases: [string, Record<string, string>][] = [
      ['https://api.openai.com/v1/chat/completions', { Authorization: 'Bearer gsk-k' }],
      ['https://api.anthropic.com/v1/messages', { 'x-api-key': 'gsk-k' }],
      [OPENAI, { Authorization: 'Bearer sk-byok' }],
      [ANTHROPIC, { 'x-api-key': 'sk-ant-byok' }],
      [OPENAI.replace('https:', 'http:'), { Authorization: 'Bearer gsk-k' }],
      [
        OPENAI.replace('www.genspark.ai', 'www.genspark.ai.evil.test'),
        { Authorization: 'Bearer gsk-k' },
      ],
      [OPENAI.replace('/api/llm_proxy/', '/api/tool_cli/'), { Authorization: 'Bearer gsk-k' }],
      [OPENAI, { Authorization: 'Bearer gsk-k', 'x-api-key': 'gsk-k' }],
      [OPENAI, { Authorization: 'Bearer gsk-a, gsk-b' }],
    ]
    for (const [url, headers] of cases) {
      const init = { headers }
      expect(await admission.prepare(url, init, fetcher)).toBe(init)
    }
    expect(fetcher).not.toHaveBeenCalled()
  })

  it('falls back unchanged and negative-caches an unusable issuer', async () => {
    for (const reply of [
      () => new Response('', { status: 503 }),
      () => new Response('not json'),
      () => new Response('x'.repeat(4097)),
      () => issued(pass(), 'tool'),
      () => issued(pass(Math.floor(Date.now() / 1000) - 600)),
    ]) {
      const { calls, fetcher } = recorder(reply)
      const admission = createNativeAdmission()
      const init = { headers: { Authorization: 'Bearer gsk-k' } }
      expect(await admission.prepare(OPENAI, init, fetcher)).toBe(init)
      expect(await admission.prepare(OPENAI, init, fetcher)).toBe(init)
      expect(calls).toHaveLength(1)
    }
  })

  it('bounds a stalled issuer by the bootstrap deadline', async () => {
    const { fetcher } = recorder(
      (_url, init) =>
        new Promise<Response>((_, reject) =>
          init.signal!.addEventListener('abort', () => reject(init.signal!.reason)),
        ),
    )
    const started = performance.now()
    const init = { headers: { Authorization: 'Bearer gsk-k' } }
    expect(await createNativeAdmission({ timeoutMs: 30 }).prepare(OPENAI, init, fetcher)).toBe(init)
    expect(performance.now() - started).toBeLessThan(1000)
  })

  it('never caches past the signed lifetime', async () => {
    let now = 1_800_000_000_000
    const { calls, fetcher } = recorder(() => issued(`1799999740-${'A'.repeat(43)}`))
    const admission = createNativeAdmission({ now: () => now })
    const headers = new Headers({ Authorization: 'Bearer gsk-k' })
    expect(await admission.pass(OPENAI, headers, fetcher)).not.toBe('')
    now += 11_000
    await admission.pass(OPENAI, headers, fetcher)
    expect(calls).toHaveLength(2)
  })

  it('coalesces callers; one cancelled waiter does not cancel the others', async () => {
    let release!: () => void
    const gate = new Promise<void>((resolve) => (release = resolve))
    const signals: AbortSignal[] = []
    const { fetcher } = recorder(async (_url, init) => {
      signals.push(init.signal!)
      await gate
      return issued(pass())
    })
    const admission = createNativeAdmission()
    const headers = new Headers({ Authorization: 'Bearer gsk-k' })
    const controller = new AbortController()
    const first = admission.pass(OPENAI, headers, fetcher, controller.signal)
    const second = admission.pass(OPENAI, headers, fetcher)
    controller.abort(new Error('cancelled'))
    await expect(first).rejects.toThrow('cancelled')
    expect(signals[0]!.aborted).toBe(false)
    release()
    expect(await second).not.toBe('')
    expect(fetcher).toHaveBeenCalledTimes(1)
  })

  it('reaches the actual provider transport (aiFetch) without changing the credential or body', async () => {
    const token = pass()
    const sent: { url: string; init: RequestInit }[] = []
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string, init: RequestInit) => {
        sent.push({ url, init })
        return url.includes('/api/native-waf/') ? issued(token) : new Response('ok')
      }),
    )
    await aiFetch(OPENAI, {
      method: 'POST',
      headers: { Authorization: 'Bearer gsk-live' },
      body: '{"a":1}',
    })
    const business = sent.find((call) => call.url === OPENAI)!
    const headers = new Headers(business.init.headers)
    expect(headers.get('x-genspark-token-pass')).toBe(token)
    expect(headers.get('authorization')).toBe('Bearer gsk-live')
    expect(headers.get('user-agent')).toBe('GenOffice')
    expect(business.init.body).toBe('{"a":1}')
    expect(business.init.redirect).toBe('error')
  })
})

describe('nativeAdmissionHosts', () => {
  it('keeps the default proxy host and adds the configured ones', () => {
    expect([...nativeAdmissionHosts({})]).toEqual(['www.genspark.ai'])
    expect([...nativeAdmissionHosts({ GSK_BASE_URL: 'https://office.example.com/x' })]).toEqual([
      'www.genspark.ai',
      'office.example.com',
    ])
    expect([
      ...nativeAdmissionHosts({
        GENOFFICE_NATIVE_WAF_HOSTS: ' Tenant.Example.com, ,second.example.com',
      }),
    ]).toEqual(['www.genspark.ai', 'tenant.example.com', 'second.example.com'])
  })
})
