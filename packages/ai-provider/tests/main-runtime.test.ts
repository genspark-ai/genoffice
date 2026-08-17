import { describe, expect, it, vi } from 'vitest'
import { CodexError, defaultAiSettings } from '../src/index'
import {
  AiMainRuntime,
  type AiMainRuntimeOptions,
  type AiRuntimeAuth,
  sanitizeAiSettings,
} from '../src/main-runtime'

const messages = {
  noGensparkApiKey: 'gsk login required',
  noApiKey: (provider: string) => `missing ${provider}`,
  noModel: 'model required',
}

function request(overrides: Partial<ReturnType<typeof defaultAiSettings>> = {}) {
  const settings = defaultAiSettings()
  return {
    requestId: 'request-1',
    settings: { ...settings, ...overrides },
    system: 'system',
    messages: [{ role: 'user' as const, text: 'hello' }],
  }
}

function auth(overrides: Partial<AiRuntimeAuth> = {}): AiRuntimeAuth {
  return {
    status: vi.fn(async () => ({ loggedIn: false })),
    login: vi.fn(async () => ({ loggedIn: false })),
    cancelLogin: vi.fn(),
    logout: vi.fn(async () => {}),
    getContext: vi.fn(async () => ({ accessToken: 'token', accountId: 'account', expiresAt: 1 })),
    ...overrides,
  }
}

describe('sanitizeAiSettings', () => {
  it('allows only the account-backed selection and normalizes Codex settings', () => {
    const settings = sanitizeAiSettings({
      provider: 'anthropic',
      providers: {
        genspark: { apiKey: 'persisted', model: 'gpt-5.2' },
        'openai-codex': {
          apiKey: 'must not persist',
          model: 'gpt-5.5',
          reasoningEffort: 'not-real',
          serviceTier: 'not real',
        },
      },
    })

    expect(settings.provider).toBe('genspark')
    expect(settings.providers.genspark).toMatchObject({ apiKey: '', model: 'gpt-5.2' })
    expect(settings.providers['openai-codex']).toMatchObject({
      apiKey: '',
      model: 'gpt-5.5',
      reasoningEffort: 'none',
      serviceTier: 'default',
    })
  })
})

describe('AiMainRuntime', () => {
  it('injects the runtime Genspark key and preserves stream chunks', async () => {
    let receivedConfig: unknown
    const streamProvider: NonNullable<AiMainRuntimeOptions['streamProvider']> = async (
      _provider,
      config,
      _system,
      _messages,
      _tools,
      _maxTokens,
      callbacks,
    ) => {
      receivedConfig = config
      callbacks.onActivity?.()
      callbacks.onDelta('hello')
      callbacks.onStopReason?.('max_tokens')
    }
    const runtime = new AiMainRuntime({
      auth: auth(),
      getGensparkApiKey: () => 'runtime-key',
      streamProvider,
    })
    const chunks: unknown[] = []

    await runtime.stream(request(), (chunk) => chunks.push(chunk), messages)

    expect(receivedConfig).toMatchObject({ apiKey: 'runtime-key' })
    expect(chunks).toEqual([
      { requestId: 'request-1', type: 'ping' },
      { requestId: 'request-1', type: 'delta', text: 'hello' },
      { requestId: 'request-1', type: 'done', stopReason: 'max_tokens' },
    ])
  })

  it('validates Codex capabilities and passes the auth context to the provider', async () => {
    const context = { accessToken: 'token', accountId: 'account', expiresAt: 1 }
    const getContext = vi.fn(async () => context)
    const streamProvider = vi.fn<NonNullable<AiMainRuntimeOptions['streamProvider']>>(
      async (
        _provider,
        _config,
        _system,
        _messages,
        _tools,
        _maxTokens,
        callbacks,
        authContext,
      ) => {
        expect(authContext).toEqual(context)
        callbacks.onDelta('codex')
      },
    )
    const runtime = new AiMainRuntime({
      auth: auth({ getContext }),
      getGensparkApiKey: () => '',
      fetchCapabilities: async () => ({
        models: [
          {
            id: 'gpt-5.5',
            reasoningEfforts: ['none', 'high'],
            serviceTiers: [{ id: 'default', name: 'Standard' }],
          },
        ],
      }),
      streamProvider,
    })
    const settings = defaultAiSettings()
    settings.provider = 'openai-codex'
    const chunks: unknown[] = []

    await runtime.stream(
      request({ provider: 'openai-codex', providers: settings.providers }),
      (chunk) => chunks.push(chunk),
      messages,
    )

    expect(getContext).toHaveBeenCalledOnce()
    expect(streamProvider).toHaveBeenCalledOnce()
    expect(chunks.at(-1)).toEqual({ requestId: 'request-1', type: 'done' })
  })

  it('returns safe Codex codes for auth, capability, and provider failures', async () => {
    const outputs: unknown[] = []
    const runtime = new AiMainRuntime({
      auth: auth({
        getContext: vi.fn(async () => {
          throw new CodexError('auth-expired')
        }),
      }),
      getGensparkApiKey: () => '',
      streamProvider: vi.fn(),
    })
    const settings = defaultAiSettings()
    settings.provider = 'openai-codex'

    await runtime.stream(
      request({ provider: 'openai-codex', providers: settings.providers }),
      (chunk) => outputs.push(chunk),
      messages,
    )
    expect(outputs).toEqual([{ requestId: 'request-1', type: 'error', errorCode: 'auth-expired' }])

    const providerFailure: NonNullable<AiMainRuntimeOptions['streamProvider']> = async () => {
      throw new Error('raw provider response')
    }
    const failureRuntime = new AiMainRuntime({
      auth: auth(),
      getGensparkApiKey: () => 'key',
      fetchCapabilities: async () => ({ models: [{ id: 'gpt-5.5', reasoningEfforts: ['none'] }] }),
      streamProvider: providerFailure,
    })
    const failure: unknown[] = []
    await failureRuntime.stream(
      request({ provider: 'openai-codex', providers: settings.providers }),
      (chunk) => failure.push(chunk),
      messages,
    )
    expect(failure).toEqual([
      { requestId: 'request-1', type: 'error', errorCode: 'provider-failure' },
    ])
  })

  it('keeps concurrent request cancellation scoped by request id', async () => {
    const streamProvider: NonNullable<AiMainRuntimeOptions['streamProvider']> = async (
      _provider,
      _config,
      system,
      _messages,
      _tools,
      _maxTokens,
      callbacks,
    ) => {
      if (system === 'one') {
        await new Promise<void>((resolve) =>
          callbacks.signal.addEventListener('abort', () => resolve(), { once: true }),
        )
        return
      }
      callbacks.onDelta('two')
    }
    const runtime = new AiMainRuntime({
      auth: auth(),
      getGensparkApiKey: () => 'key',
      streamProvider,
    })
    const first = request({})
    first.system = 'one'
    const second = request({})
    second.requestId = 'request-2'
    second.system = 'two'
    const firstChunks: unknown[] = []
    const secondChunks: unknown[] = []
    const firstRun = runtime.stream(first, (chunk) => firstChunks.push(chunk), messages)
    const secondRun = runtime.stream(second, (chunk) => secondChunks.push(chunk), messages)
    await secondRun
    runtime.cancel('request-1')
    await firstRun

    expect(firstChunks).toEqual([{ requestId: 'request-1', type: 'done' }])
    expect(secondChunks).toEqual([
      { requestId: 'request-2', type: 'delta', text: 'two' },
      { requestId: 'request-2', type: 'done' },
    ])
  })
})
