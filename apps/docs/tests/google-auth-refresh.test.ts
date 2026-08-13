/**
 * Token refresh math: refreshAccessToken exchanges a refresh token for a fresh
 * access token, and tokenNeedsRefresh decides when that's due. Electron is
 * mocked (app.getPath / shell) since google-auth.ts touches it at import time
 * even though these two functions never call it directly.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest'

vi.mock('electron', () => ({
  app: { getPath: () => '/tmp/genoffice-test-userdata' },
  shell: { openExternal: vi.fn() },
}))

const { refreshAccessToken, tokenNeedsRefresh } = await import('../src/main/google-auth')

describe('tokenNeedsRefresh', () => {
  it('is true once within the refresh skew window', () => {
    const now = Date.now()
    expect(
      tokenNeedsRefresh({
        accessToken: 'a',
        refreshToken: 'r',
        expiresAt: now + 30_000,
        scope: '',
      }),
    ).toBe(true)
  })

  it('is false with plenty of time left', () => {
    const now = Date.now()
    expect(
      tokenNeedsRefresh({
        accessToken: 'a',
        refreshToken: 'r',
        expiresAt: now + 10 * 60_000,
        scope: '',
      }),
    ).toBe(false)
  })

  it('respects a custom skew', () => {
    const now = Date.now()
    expect(
      tokenNeedsRefresh(
        { accessToken: 'a', refreshToken: 'r', expiresAt: now + 5_000, scope: '' },
        1000,
      ),
    ).toBe(false)
  })
})

describe('refreshAccessToken', () => {
  const creds = { clientId: 'client-123', clientSecret: 'secret-456' }

  beforeEach(() => vi.restoreAllMocks())

  it('exchanges the refresh token via the token endpoint and computes expiresAt', async () => {
    const fixedNow = 1_700_000_000_000
    vi.spyOn(Date, 'now').mockReturnValue(fixedNow)
    const fetchMock = vi.fn(async (url: string, init: RequestInit) => {
      expect(url).toBe('https://oauth2.googleapis.com/token')
      const body = new URLSearchParams(init.body as string)
      expect(body.get('grant_type')).toBe('refresh_token')
      expect(body.get('refresh_token')).toBe('rt-abc')
      expect(body.get('client_id')).toBe(creds.clientId)
      expect(body.get('client_secret')).toBe(creds.clientSecret)
      return new Response(
        JSON.stringify({ access_token: 'at-new', expires_in: 3600, scope: 'drive.file' }),
        { status: 200 },
      )
    })

    const result = await refreshAccessToken(creds, 'rt-abc', fetchMock as unknown as typeof fetch)
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.tokens).toEqual({
      accessToken: 'at-new',
      refreshToken: 'rt-abc', // preserved: Google doesn't reissue it on refresh
      expiresAt: fixedNow + 3600 * 1000,
      scope: 'drive.file',
    })
  })

  it('surfaces a non-2xx response as a typed failure, never throwing', async () => {
    const fetchMock = vi.fn(async () => new Response('invalid_grant', { status: 400 }))
    const result = await refreshAccessToken(creds, 'stale-rt', fetchMock as unknown as typeof fetch)
    expect(result.ok).toBe(false)
    if (result.ok) return
    expect(result.error).toContain('400')
  })
})
