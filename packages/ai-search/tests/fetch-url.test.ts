import { describe, it, expect, beforeAll } from 'vitest'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fetchUrl } from '../src/index'

// fetchUrl now routes real fetches through TinyFish over MCP. Without a stored
// TinyFish token it must fail fast — never touch the network — and tell the caller
// to connect. Point auth storage at an empty temp dir so "not signed in" is
// deterministic, then pin two behaviours: unsafe URLs are refused up front by the
// SSRF guard, and a safe URL with no sign-in returns a clear connect-first error.
beforeAll(() => {
  process.env.GENOFFICE_AUTH_DIR = join(tmpdir(), 'genoffice-tinyfish-test-noauth')
})

describe('fetchUrl (SSRF gate + connect-first)', () => {
  it('asks the user to connect TinyFish when a safe URL is fetched with no sign-in', async () => {
    const r = await fetchUrl('https://1.1.1.1/page')
    expect(r.method).toBe('error')
    expect(r.error?.toLowerCase()).toContain('connect')
  })

  it('refuses localhost before any TinyFish call', async () => {
    const r = await fetchUrl('http://localhost/admin')
    expect(r.method).toBe('error')
    expect(r.error?.toLowerCase()).toContain('refused')
  })

  it('refuses a private/link-local literal address', async () => {
    const r = await fetchUrl('http://169.254.169.254/latest/meta-data/')
    expect(r.method).toBe('error')
    expect(r.error?.toLowerCase()).toContain('refused')
  })

  it('refuses a non-http(s) scheme', async () => {
    const r = await fetchUrl('file:///etc/passwd')
    expect(r.method).toBe('error')
    expect(r.error?.toLowerCase()).toContain('refused')
  })

  it('refuses a malformed URL', async () => {
    const r = await fetchUrl('not a url')
    expect(r.method).toBe('error')
    expect(r.error?.toLowerCase()).toContain('refused')
  })
})
