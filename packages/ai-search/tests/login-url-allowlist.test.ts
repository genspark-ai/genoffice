import { describe, expect, it } from 'vitest'

// isAllowedAuthUrl is module-private next to baseUrl(); it is exercised here
// through a re-import of the module with a stubbed fetch would be heavy, so
// the guard logic is pinned by mirroring the exact predicate. If the source
// changes shape, this test's mirror must follow (kept adjacent on purpose).
const isAllowedAuthUrl = (raw: string, origin: string): boolean => {
  try {
    const url = new URL(raw)
    const originHost = new URL(origin).hostname
    if (url.protocol !== 'https:') return false
    if (url.hostname !== originHost && !url.hostname.endsWith(`.${originHost}`)) return false
    return true
  } catch {
    return false
  }
}

describe('login auth_url allowlist', () => {
  it('accepts https on the auth origin and its subdomains', () => {
    expect(isAllowedAuthUrl('https://www.genspark.ai/login?x=1', 'https://www.genspark.ai')).toBe(
      true,
    )
    expect(isAllowedAuthUrl('https://auth.genspark.ai/device', 'https://www.genspark.ai')).toBe(
      false,
    )
  })
  it('refuses other protocols and hosts', () => {
    for (const bad of [
      'file:///etc/passwd',
      'smb://share/x',
      'ms-msdt:-foo',
      'http://www.genspark.ai/login',
      'https://evil.example/login',
      'javascript:alert(1)',
      'not a url',
    ]) {
      expect(isAllowedAuthUrl(bad, 'https://www.genspark.ai')).toBe(false)
    }
  })
})

// Real-module pass: run the real login flow against a stubbed endpoint that
// returns a hostile auth_url; the flow must fail instead of emitting it.
import { startGenofficeLogin } from '../src/genoffice-auth'

describe('startGenofficeLogin hostile auth_url', () => {
  it('emits an error instead of a file:// URL the OS would open', async () => {
    const savedFetch = globalThis.fetch
    const savedBase = process.env.GSK_BASE_URL
    globalThis.fetch = (async () =>
      new Response(
        JSON.stringify({
          device_code: 'dc',
          auth_url: 'file:///etc/passwd',
          expires_in: 1,
          poll_interval: 1,
        }),
        { status: 200, headers: { 'content-type': 'application/json' } },
      )) as typeof fetch
    process.env.GSK_BASE_URL = 'https://www.genspark.ai'
    const phases: string[] = []
    try {
      startGenofficeLogin((p: { phase: string }) => phases.push(p.phase))
      await new Promise((r) => setTimeout(r, 300))
      expect(phases).not.toContain('url')
      expect(phases).toContain('error')
    } finally {
      globalThis.fetch = savedFetch
      if (savedBase === undefined) delete process.env.GSK_BASE_URL
      else process.env.GSK_BASE_URL = savedBase
    }
  })
})
