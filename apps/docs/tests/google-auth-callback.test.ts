/**
 * OAuth loopback callback: the redirect_uri sent in the token exchange must
 * match the port the loopback server actually listened on. Regression test
 * for AOF-437 P0 (server.close() before server.address() returned null,
 * producing "http://127.0.0.1:null/callback").
 */
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

const userDataDir = mkdtempSync(join(tmpdir(), 'genoffice-auth-test-'))

vi.mock('electron', () => ({
  app: { getPath: () => userDataDir },
  shell: { openExternal: vi.fn() },
}))

const { runGoogleAuthFlow } = await import('../src/main/google-auth')

describe('runGoogleAuthFlow callback', () => {
  const prevId = process.env.GENOFFICE_GOOGLE_CLIENT_ID
  const prevSecret = process.env.GENOFFICE_GOOGLE_CLIENT_SECRET

  beforeEach(() => {
    process.env.GENOFFICE_GOOGLE_CLIENT_ID = 'client-123'
    process.env.GENOFFICE_GOOGLE_CLIENT_SECRET = 'secret-456'
  })

  afterEach(() => {
    process.env.GENOFFICE_GOOGLE_CLIENT_ID = prevId
    process.env.GENOFFICE_GOOGLE_CLIENT_SECRET = prevSecret
    vi.restoreAllMocks()
  })

  it('sends a redirect_uri with the real loopback port, not null', async () => {
    let capturedRedirectUri: string | null = null
    const realFetch = globalThis.fetch.bind(globalThis)
    const fetchSpy = vi.spyOn(globalThis, 'fetch').mockImplementation(async (url, init) => {
      // Only intercept the token-exchange POST to Google; let the test's own
      // GET to the local loopback server (the simulated browser redirect)
      // pass through untouched.
      if (url !== 'https://oauth2.googleapis.com/token') return realFetch(url as string, init)
      const body = new URLSearchParams((init as RequestInit).body as string)
      capturedRedirectUri = body.get('redirect_uri')
      return new Response(
        JSON.stringify({
          access_token: 'at-1',
          refresh_token: 'rt-1',
          expires_in: 3600,
          scope: 'drive.file',
        }),
        { status: 200 },
      )
    })

    const flowPromise = runGoogleAuthFlow()

    // Grab the port the loopback server actually bound by intercepting the
    // authorization URL passed to shell.openExternal.
    const { shell } = await import('electron')
    await vi.waitFor(() => {
      expect((shell.openExternal as ReturnType<typeof vi.fn>).mock.calls.length).toBeGreaterThan(0)
    })
    const authUrl = new URL(
      (shell.openExternal as ReturnType<typeof vi.fn>).mock.calls[0][0] as string,
    )
    const redirectUri = authUrl.searchParams.get('redirect_uri')!
    const state = authUrl.searchParams.get('state')!
    expect(redirectUri).toMatch(/^http:\/\/127\.0\.0\.1:\d+\/callback$/)

    const callbackUrl = new URL(redirectUri)
    callbackUrl.searchParams.set('code', 'fake-code')
    callbackUrl.searchParams.set('state', state)
    const res = await fetch(callbackUrl.toString())
    expect(res.ok).toBe(true)

    const result = await flowPromise
    expect(result.ok).toBe(true)
    expect(fetchSpy).toHaveBeenCalled()
    expect(capturedRedirectUri).toBe(redirectUri)
    expect(capturedRedirectUri).not.toContain('null')
  })
})

afterEach(() => {
  rmSync(userDataDir, { recursive: true, force: true })
})
