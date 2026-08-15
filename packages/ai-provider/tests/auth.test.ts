import { describe, expect, it, vi } from 'vitest'
import {
  CodexAuthService,
  type CodexAuthDependencies,
  type CodexCredentialStore,
} from '../src/auth'

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

const transientRefreshFailures: Array<[string, CodexAuthDependencies['fetch']]> = [
  [
    'network failure',
    async () => {
      throw new Error('network failure')
    },
  ],
  ['HTTP 408', async () => new Response(null, { status: 408 })],
  ['HTTP 429', async () => new Response(null, { status: 429 })],
  ['HTTP 500', async () => new Response(null, { status: 500 })],
  ['malformed success', async () => new Response('not-json', { status: 200 })],
]

function service(overrides: Partial<ConstructorParameters<typeof CodexAuthService>[0]> = {}) {
  const store = memoryStore()
  const callback = {
    ready: Promise.resolve(),
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

  it('waits for callback readiness before opening the browser', async () => {
    let resolveReady!: () => void
    const ready = new Promise<void>((resolve) => {
      resolveReady = resolve
    })
    const openBrowser = vi.fn()
    const auth = new CodexAuthService({
      store: memoryStore(),
      clock: () => 1_000_000,
      fetch: vi.fn(async () => tokenResponse()),
      openBrowser,
      beginCallback: (state) => ({
        ready,
        wait: Promise.resolve({ state, code: 'authorization-code' }),
        cancel: () => {},
      }),
    })

    const login = auth.login()
    await Promise.resolve()
    expect(openBrowser).not.toHaveBeenCalled()

    resolveReady()
    await login
    expect(openBrowser).toHaveBeenCalledOnce()
  })

  it('does not open the browser when callback readiness fails', async () => {
    const complete = vi.fn()
    const openBrowser = vi.fn()
    const ready = Promise.reject(new Error('callback unavailable'))
    const auth = new CodexAuthService({
      store: memoryStore(),
      clock: () => 1_000_000,
      fetch: vi.fn(async () => tokenResponse()),
      openBrowser,
      beginCallback: () => ({
        ready,
        wait: Promise.resolve({ state: 'unused', code: 'unused' }),
        cancel: () => {},
        complete,
      }),
    })

    await expect(auth.login()).rejects.toMatchObject({ code: 'auth-temporary' })
    expect(openBrowser).not.toHaveBeenCalled()
    expect(complete).toHaveBeenCalledOnce()
    expect(complete).toHaveBeenCalledWith(false)
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
        ready: Promise.resolve(),
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
        ready: Promise.resolve(),
        wait: Promise.resolve({ state, code: 'authorization-code' }),
        cancel: () => {},
        complete,
      }),
    })

    await expect(auth.login()).rejects.toMatchObject({ code: 'auth-temporary', status: 403 })
    expect(complete).toHaveBeenCalledWith(false)
  })

  it('rejects a callback state mismatch without exchanging or storing credentials', async () => {
    const { auth, deps, store } = service({
      beginCallback: () => ({
        ready: Promise.resolve(),
        wait: Promise.resolve({ state: 'wrong-state', code: 'authorization-code' }),
        cancel: () => {},
      }),
    })

    await expect(auth.login()).rejects.toMatchObject({ code: 'provider-failure' })
    expect(deps.fetch).not.toHaveBeenCalled()
    expect(store.set).not.toHaveBeenCalled()
  })

  it('propagates callback timeout and supports cancellation', async () => {
    const cancel = vi.fn()
    let rejectCallback!: (error: Error) => void
    const wait = new Promise<never>((_, reject) => {
      rejectCallback = reject
    })
    const { auth } = service({
      beginCallback: () => ({
        ready: Promise.resolve(),
        wait,
        cancel,
      }),
    })

    const login = auth.login()
    const result = expect(login).rejects.toMatchObject({ code: 'auth-temporary' })
    auth.cancelLogin()
    expect(cancel).toHaveBeenCalledOnce()
    rejectCallback(new Error('ChatGPT sign-in timed out'))
    await result
  })

  it('shares one login attempt across concurrent callers', async () => {
    let resolveCallback!: (result: { state: string; code: string }) => void
    let callbackState = ''
    const wait = new Promise<{ state: string; code: string }>((resolve) => {
      resolveCallback = resolve
    })
    const { auth, deps } = service({
      beginCallback: vi.fn((state) => {
        callbackState = state
        return { ready: Promise.resolve(), wait, cancel: vi.fn() }
      }),
    })

    const left = auth.login()
    const right = auth.login()
    await vi.waitFor(() => expect(deps.openBrowser).toHaveBeenCalledOnce())
    resolveCallback({ state: callbackState, code: 'authorization-code' })

    await expect(Promise.all([left, right])).resolves.toEqual([
      { loggedIn: true, email: 'person@example.test' },
      { loggedIn: true, email: 'person@example.test' },
    ])
    expect(deps.beginCallback).toHaveBeenCalledOnce()
    expect(deps.openBrowser).toHaveBeenCalledOnce()
  })

  it('cancels the shared login and allows a later attempt', async () => {
    const callbackStates: string[] = []
    const callbacks: Array<{
      resolve: (result: { state: string; code: string }) => void
      reject: (error: Error) => void
    }> = []
    const cancels: Array<ReturnType<typeof vi.fn>> = []
    const { auth, deps } = service({
      beginCallback: vi.fn((state) => {
        callbackStates.push(state)
        let resolve!: (result: { state: string; code: string }) => void
        let reject!: (error: Error) => void
        const cancel = vi.fn()
        const wait = new Promise<{ state: string; code: string }>((resolveWait, rejectWait) => {
          resolve = resolveWait
          reject = rejectWait
        })
        callbacks.push({ resolve, reject })
        cancel.mockImplementation(() => reject(new Error('ChatGPT sign-in cancelled')))
        cancels.push(cancel)
        return { ready: Promise.resolve(), wait, cancel }
      }),
    })

    const first = auth.login()
    const shared = auth.login()
    await vi.waitFor(() => expect(deps.openBrowser).toHaveBeenCalledOnce())
    auth.cancelLogin()
    expect(cancels[0]).toHaveBeenCalledOnce()
    await expect(Promise.allSettled([first, shared])).resolves.toEqual([
      { status: 'rejected', reason: expect.any(Error) },
      { status: 'rejected', reason: expect.any(Error) },
    ])

    const later = auth.login()
    await vi.waitFor(() => expect(deps.openBrowser).toHaveBeenCalledTimes(2))
    callbacks[1]!.resolve({ state: callbackStates[1]!, code: 'authorization-code' })
    await expect(later).resolves.toEqual({ loggedIn: true, email: 'person@example.test' })
    expect(deps.beginCallback).toHaveBeenCalledTimes(2)
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
      beginCallback: () => ({
        ready: Promise.resolve(),
        wait: Promise.reject(new Error('not used')),
        cancel: () => {},
      }),
    })

    const [left, right] = await Promise.all([auth.getContext(), auth.getContext()])

    expect(left.accessToken).toBe('new-access')
    expect(right.accessToken).toBe('new-access')
    expect(fetch).toHaveBeenCalledOnce()
  })

  it('reports no account without checking the network when credentials are absent', async () => {
    const { auth, deps } = service()

    await expect(auth.status()).resolves.toEqual({ loggedIn: false })
    expect(deps.fetch).not.toHaveBeenCalled()
  })

  it('reports fresh credentials without refreshing', async () => {
    const store = memoryStore({
      accessToken: 'access',
      refreshToken: 'refresh',
      accountId: 'account-123',
      email: 'person@example.test',
      expiresAt: 2_000_000,
    })
    const fetch = vi.fn(async () => tokenResponse())
    const auth = new CodexAuthService({
      store,
      clock: () => 1_000_000,
      fetch,
      openBrowser: () => {},
      beginCallback: () => ({
        ready: Promise.resolve(),
        wait: Promise.reject(new Error('not used')),
        cancel: () => {},
      }),
    })

    await expect(auth.status()).resolves.toEqual({
      loggedIn: true,
      email: 'person@example.test',
    })
    expect(fetch).not.toHaveBeenCalled()
  })

  it('refreshes expired credentials before reporting logged in', async () => {
    const store = memoryStore({
      accessToken: 'old-access',
      refreshToken: 'refresh-token',
      accountId: 'account-123',
      email: 'person@example.test',
      expiresAt: 999_999,
    })
    const fetch = vi.fn(async () => tokenResponse({ access_token: 'new-access' }))
    const auth = new CodexAuthService({
      store,
      clock: () => 1_000_000,
      fetch,
      openBrowser: () => {},
      beginCallback: () => ({
        ready: Promise.resolve(),
        wait: Promise.reject(new Error('not used')),
        cancel: () => {},
      }),
    })

    await expect(auth.status()).resolves.toEqual({
      loggedIn: true,
      email: 'person@example.test',
    })
    expect(fetch).toHaveBeenCalledOnce()
    expect(store.set).toHaveBeenCalledWith(expect.objectContaining({ accessToken: 'new-access' }))
  })

  it('reports confirmed expiry as logged out after deleting credentials', async () => {
    const store = memoryStore({
      accessToken: 'old-access',
      refreshToken: 'refresh-token',
      accountId: 'account-123',
      email: 'person@example.test',
      expiresAt: 999_999,
    })
    const auth = new CodexAuthService({
      store,
      clock: () => 1_000_000,
      fetch: async () => new Response(null, { status: 401 }),
      openBrowser: () => {},
      beginCallback: () => ({
        ready: Promise.resolve(),
        wait: Promise.reject(new Error('not used')),
        cancel: () => {},
      }),
    })

    await expect(auth.status()).resolves.toEqual({
      loggedIn: false,
      errorCode: 'auth-expired',
    })
    expect(store.delete).toHaveBeenCalledOnce()
    expect(await store.get()).toBeUndefined()
  })

  it('retains credentials and reports a temporary status failure', async () => {
    const credentials = {
      accessToken: 'old-access',
      refreshToken: 'refresh-token',
      accountId: 'account-123',
      email: 'person@example.test',
      expiresAt: 999_999,
    }
    const store = memoryStore(credentials)
    const auth = new CodexAuthService({
      store,
      clock: () => 1_000_000,
      fetch: async () => {
        throw new Error('network failure')
      },
      openBrowser: () => {},
      beginCallback: () => ({
        ready: Promise.resolve(),
        wait: Promise.reject(new Error('not used')),
        cancel: () => {},
      }),
    })

    await expect(auth.status()).resolves.toEqual({
      loggedIn: true,
      email: 'person@example.test',
      errorCode: 'auth-temporary',
    })
    expect(store.delete).not.toHaveBeenCalled()
    expect(await store.get()).toEqual(credentials)
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
      beginCallback: () => ({
        ready: Promise.resolve(),
        wait: Promise.reject(new Error('not used')),
        cancel: () => {},
      }),
    })

    await expect(auth.getContext()).rejects.toMatchObject({ code: 'auth-expired' })
    expect(store.delete).toHaveBeenCalledOnce()
  })

  it.each([400, 401])(
    'deletes credentials after a confirmed HTTP %s refresh rejection',
    async (status) => {
      const store = memoryStore({
        accessToken: 'old-access',
        refreshToken: 'refresh-token',
        accountId: 'account-123',
        expiresAt: 999_999,
      })
      const auth = new CodexAuthService({
        store,
        clock: () => 1_000_000,
        fetch: async () => new Response(null, { status }),
        openBrowser: () => {},
        beginCallback: () => ({
          ready: Promise.resolve(),
          wait: Promise.reject(new Error('not used')),
          cancel: () => {},
        }),
      })

      await expect(auth.getContext()).rejects.toMatchObject({ code: 'auth-expired' })
      expect(store.delete).toHaveBeenCalledOnce()
      expect(await store.get()).toBeUndefined()
    },
  )

  it.each(transientRefreshFailures)(
    'preserves credentials after a transient %s refresh failure',
    async (_label, fetch) => {
      const credentials = {
        accessToken: 'old-access',
        refreshToken: 'refresh-token',
        accountId: 'account-123',
        expiresAt: 999_999,
      }
      const store = memoryStore(credentials)
      const auth = new CodexAuthService({
        store,
        clock: () => 1_000_000,
        fetch,
        openBrowser: () => {},
        beginCallback: () => ({
          ready: Promise.resolve(),
          wait: Promise.reject(new Error('not used')),
          cancel: () => {},
        }),
      })

      await expect(auth.getContext()).rejects.toMatchObject({ code: 'auth-temporary' })
      expect(store.delete).not.toHaveBeenCalled()
      expect(await store.get()).toEqual(credentials)
    },
  )

  it('can retry after a transient refresh failure', async () => {
    const store = memoryStore({
      accessToken: 'old-access',
      refreshToken: 'refresh-token',
      accountId: 'account-123',
      expiresAt: 999_999,
    })
    const fetch = vi
      .fn<CodexAuthDependencies['fetch']>()
      .mockRejectedValueOnce(new Error('network failure'))
      .mockResolvedValueOnce(tokenResponse({ access_token: 'new-access' }))
    const auth = new CodexAuthService({
      store,
      clock: () => 1_000_000,
      fetch,
      openBrowser: () => {},
      beginCallback: () => ({
        ready: Promise.resolve(),
        wait: Promise.reject(new Error('not used')),
        cancel: () => {},
      }),
    })

    await expect(auth.getContext()).rejects.toMatchObject({ code: 'auth-temporary' })
    await expect(auth.getContext()).resolves.toMatchObject({ accessToken: 'new-access' })
    expect(fetch).toHaveBeenCalledTimes(2)
    expect(store.delete).not.toHaveBeenCalled()
  })

  it('rejects non-200 authorization exchanges without persisting credentials', async () => {
    const { auth, store } = service({
      fetch: async () => new Response('forbidden', { status: 403 }),
    })

    await expect(auth.login()).rejects.toMatchObject({ code: 'auth-temporary', status: 403 })
    expect(store.set).not.toHaveBeenCalled()
  })

  it('deletes credentials on logout', async () => {
    const { auth, store } = service()

    await auth.logout()

    expect(store.delete).toHaveBeenCalledOnce()
    expect(await auth.status()).toEqual({ loggedIn: false })
  })
})
