import { describe, expect, it, vi } from 'vitest'
import { CodexAuthService, type CodexCredentialStore } from '../src/auth'

const claims = {
  email: 'person@example.test',
  'https://api.openai.com/auth': { chatgpt_account_id: 'account-123' },
}
const idToken = `header.${Buffer.from(JSON.stringify(claims)).toString('base64url')}.signature`

function tokenResponse(overrides: Record<string, unknown> = {}) {
  return new Response(
    JSON.stringify({
      access_token: 'access-token',
      refresh_token: 'refresh-token',
      id_token: idToken,
      expires_in: 3600,
      ...overrides,
    }),
    { status: 200, headers: { 'Content-Type': 'application/json' } },
  )
}

function memoryStore(initial?: Parameters<CodexCredentialStore['set']>[0]) {
  let value = initial
  return {
    get: vi.fn(async () => value),
    set: vi.fn(async (next) => {
      value = next
    }),
    delete: vi.fn(async () => {
      value = undefined
    }),
  } satisfies CodexCredentialStore
}

function service(overrides: Partial<ConstructorParameters<typeof CodexAuthService>[0]> = {}) {
  const store = memoryStore()
  const callback = {
    wait: Promise.resolve({ state: '', code: 'authorization-code' }),
    cancel: vi.fn(),
  }
  const deps = {
    store,
    clock: () => 1_000_000,
    fetch: vi.fn(async () => tokenResponse()),
    openBrowser: vi.fn(async () => {}),
    beginCallback: vi.fn((state) => ({
      ...callback,
      wait: Promise.resolve({ state, code: 'authorization-code' }),
    })),
    ...overrides,
  }
  return { auth: new CodexAuthService(deps), deps, store, callback }
}

describe('CodexAuthService', () => {
  it('exchanges a PKCE callback and exposes only redacted account status', async () => {
    const { auth, deps, store } = service()

    await auth.login()

    expect(deps.openBrowser).toHaveBeenCalledWith(expect.stringContaining('code_challenge='))
    expect(deps.fetch).toHaveBeenCalledWith(
      'https://auth.openai.com/oauth/token',
      expect.objectContaining({ method: 'POST' }),
    )
    expect(store.set).toHaveBeenCalledWith(
      expect.objectContaining({
        accessToken: 'access-token',
        refreshToken: 'refresh-token',
        accountId: 'account-123',
      }),
    )
    expect(await auth.status()).toEqual({ loggedIn: true, email: 'person@example.test' })
  })

  it('confirms browser success only after secure credential save', async () => {
    const events: string[] = []
    const store = memoryStore()
    store.set.mockImplementation(async (next) => {
      events.push(`save:${next.accountId}`)
    })
    const complete = vi.fn((ok: boolean) => events.push(`complete:${ok}`))
    const auth = new CodexAuthService({
      store,
      clock: () => 1_000_000,
      fetch: vi.fn(async () => tokenResponse()),
      openBrowser: () => {},
      beginCallback: (state) => ({
        wait: Promise.resolve({ state, code: 'authorization-code' }),
        cancel: () => {},
        complete,
      }),
    })

    await auth.login()

    expect(events).toEqual(['save:account-123', 'complete:true'])
  })

  it('reports safe browser failure when token exchange fails', async () => {
    const complete = vi.fn()
    const { auth } = service({
      fetch: async () => new Response('forbidden', { status: 403 }),
      beginCallback: (state) => ({
        wait: Promise.resolve({ state, code: 'authorization-code' }),
        cancel: () => {},
        complete,
      }),
    })

    await expect(auth.login()).rejects.toThrow('ChatGPT sign-in failed (HTTP 403)')
    expect(complete).toHaveBeenCalledWith(false)
  })

  it('rejects a callback state mismatch without exchanging or storing credentials', async () => {
    const { auth, deps, store } = service({
      beginCallback: () => ({
        wait: Promise.resolve({ state: 'wrong-state', code: 'authorization-code' }),
        cancel: () => {},
      }),
    })

    await expect(auth.login()).rejects.toThrow('ChatGPT sign-in state mismatch')
    expect(deps.fetch).not.toHaveBeenCalled()
    expect(store.set).not.toHaveBeenCalled()
  })

  it('propagates callback timeout and supports cancellation', async () => {
    const cancel = vi.fn()
    let rejectCallback!: (error: Error) => void
    const { auth } = service({
      beginCallback: () => ({
        wait: new Promise((_, reject) => {
          rejectCallback = reject
        }),
        cancel,
      }),
    })

    const login = auth.login()
    auth.cancelLogin()
    expect(cancel).toHaveBeenCalledOnce()
    rejectCallback(new Error('ChatGPT sign-in timed out'))
    await expect(login).rejects.toThrow('ChatGPT sign-in timed out')
  })

  it('refreshes expired credentials once for concurrent requests', async () => {
    const store = memoryStore({
      accessToken: 'old-access',
      refreshToken: 'refresh-token',
      accountId: 'account-123',
      expiresAt: 999_999,
    })
    const fetch = vi.fn(async () => tokenResponse({ access_token: 'new-access' }))
    const auth = new CodexAuthService({
      store,
      clock: () => 1_000_000,
      fetch,
      openBrowser: () => {},
      beginCallback: () => ({ wait: Promise.reject(new Error('not used')), cancel: () => {} }),
    })

    const [left, right] = await Promise.all([auth.getContext(), auth.getContext()])

    expect(left.accessToken).toBe('new-access')
    expect(right.accessToken).toBe('new-access')
    expect(fetch).toHaveBeenCalledOnce()
  })

  it('deletes credentials after a confirmed refresh failure', async () => {
    const store = memoryStore({
      accessToken: 'old-access',
      refreshToken: 'refresh-token',
      accountId: 'account-123',
      expiresAt: 999_999,
    })
    const auth = new CodexAuthService({
      store,
      clock: () => 1_000_000,
      fetch: async () => new Response('denied', { status: 401 }),
      openBrowser: () => {},
      beginCallback: () => ({ wait: Promise.reject(new Error('not used')), cancel: () => {} }),
    })

    await expect(auth.getContext()).rejects.toThrow('ChatGPT sign-in expired')
    expect(store.delete).toHaveBeenCalledOnce()
  })

  it('rejects non-200 authorization exchanges without persisting credentials', async () => {
    const { auth, store } = service({
      fetch: async () => new Response('forbidden', { status: 403 }),
    })

    await expect(auth.login()).rejects.toThrow('ChatGPT sign-in failed (HTTP 403)')
    expect(store.set).not.toHaveBeenCalled()
  })

  it('deletes credentials on logout', async () => {
    const { auth, store } = service()

    await auth.logout()

    expect(store.delete).toHaveBeenCalledOnce()
    expect(await auth.status()).toEqual({ loggedIn: false })
  })
})
