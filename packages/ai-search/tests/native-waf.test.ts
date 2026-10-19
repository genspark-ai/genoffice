import { describe, expect, it } from 'vitest'
import { NativeWaf, nativeWaf, nativeWafHosts } from '../src/native-waf'
import { reloadGenofficeAuth } from '../src/genoffice-auth'

const url = 'https://www.genspark.ai/api/api_tokens/create'
const pass = () => `${Math.floor(Date.now() / 1000)}-${'a'.repeat(43)}`
const init = () => ({ headers: { Cookie: 'session_id=original', 'GS-Test': 'legacy' } })

describe('optional cookie admission', () => {
  it('API-key watcher reload cannot invalidate an in-flight cookie pass', async () => {
    nativeWaf.clear()
    const token = pass()
    let finish!: (response: Response) => void
    const pending = nativeWaf.prepare(
      url,
      init(),
      () =>
        new Promise((resolve) => {
          finish = resolve
        }),
    )
    reloadGenofficeAuth()
    finish(Response.json({ data: { token } }))
    expect(new Headers((await pending).headers).get('X-Genspark-Native-Pass')).toBe(token)
    nativeWaf.clear()
  })
  it('coalesces cookie-bound issuer requests and preserves legacy headers', async () => {
    const admission = new NativeWaf()
    let calls = 0
    const token = pass()
    const fetcher = async (_url: string, outgoing: RequestInit) => {
      calls++
      expect(new Headers(outgoing.headers).get('Cookie')).toBe('session_id=original')
      expect(new Headers(outgoing.headers).get('GS-Test')).toBe('legacy')
      expect(outgoing.redirect).toBe('manual')
      return Response.json({ data: { token } })
    }
    const requests = await Promise.all([
      admission.prepare(url, init(), fetcher),
      admission.prepare(url, init(), fetcher),
    ])
    expect(calls).toBe(1)
    for (const request of requests) {
      expect(new Headers(request.headers).get('X-Genspark-Native-Pass')).toBe(token)
      expect(request.credentials).toBe('omit')
    }
  })

  it('does no issuer work for Bearer, external providers or missing cookies', async () => {
    const admission = new NativeWaf()
    const fetcher = async () => {
      throw new Error('must not call issuer')
    }
    const bearer = { headers: { Authorization: 'Bearer own-key' } }
    expect((await admission.prepare(url, bearer, fetcher)).redirect).toBeUndefined()
    expect(
      (await admission.prepare('https://provider.invalid/api/x', init(), fetcher)).redirect,
    ).toBeUndefined()
  })

  it('issuer failure keeps the cookie and negative cache; clear fences late passes', async () => {
    const admission = new NativeWaf()
    let calls = 0
    const failing = async () => {
      calls++
      return new Response('oversize'.repeat(1000))
    }
    for (let i = 0; i < 2; i++) {
      const outgoing = await admission.prepare(url, init(), failing)
      expect(new Headers(outgoing.headers).get('Cookie')).toBe('session_id=original')
      expect(new Headers(outgoing.headers).has('X-Genspark-Native-Pass')).toBe(false)
    }
    expect(calls).toBe(1)
    admission.clear()
    let finish!: (response: Response) => void
    const pending = admission.prepare(
      url,
      init(),
      () =>
        new Promise((resolve) => {
          finish = resolve
        }),
    )
    admission.clear()
    finish(Response.json({ data: { token: pass() } }))
    expect(new Headers((await pending).headers).has('X-Genspark-Native-Pass')).toBe(false)
  })

  it('admitted hosts come from the configured base URL and the extra-hosts setting', async () => {
    expect([...nativeWafHosts({})]).toEqual(['www.genspark.ai'])
    expect([...nativeWafHosts({ GSK_BASE_URL: 'https://office.example.com/' })]).toEqual([
      'office.example.com',
    ])
    expect([
      ...nativeWafHosts({ GENOFFICE_NATIVE_WAF_HOSTS: ' Tenant.Example.com, ,second.example.com' }),
    ]).toEqual(['www.genspark.ai', 'tenant.example.com', 'second.example.com'])

    const admission = new NativeWaf()
    const token = pass()
    const fetcher = async () => Response.json({ data: { token } })
    const tenant = 'https://tenant.example.com/api/api_tokens/create'
    expect(
      new Headers((await admission.prepare(tenant, init(), fetcher)).headers).has(
        'X-Genspark-Native-Pass',
      ),
    ).toBe(false)
    process.env.GENOFFICE_NATIVE_WAF_HOSTS = 'tenant.example.com'
    try {
      expect(
        new Headers((await admission.prepare(tenant, init(), fetcher)).headers).get(
          'X-Genspark-Native-Pass',
        ),
      ).toBe(token)
    } finally {
      delete process.env.GENOFFICE_NATIVE_WAF_HOSTS
    }
  })
})
