import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { registerAiSettingsIpc, type AiSettingsIpcDeps } from '../src/ipc'
import { defaultAiSettings } from '../src/providers'
import type { AiSettings } from '../src/types'

type Handler = (event: unknown, ...args: unknown[]) => unknown

/** in-memory stand-in for ipcMain.handle; captures the per-channel handlers */
function fakeIpc() {
  const handlers = new Map<string, Handler>()
  return {
    handlers,
    handle: vi.fn((channel: string, fn: Handler) => {
      handlers.set(channel, fn)
    }),
  }
}

function invoke(handlers: Map<string, Handler>, channel: string, ...args: unknown[]): unknown {
  const handler = handlers.get(channel)
  if (!handler) throw new Error(`no handler for ${channel}`)
  return handler({}, ...args)
}

function makeDeps(overrides: Partial<AiSettingsIpcDeps> = {}): AiSettingsIpcDeps {
  return {
    settingsPath: () => '/tmp/ai-settings.json',
    readJson: vi.fn(),
    writeJson: vi.fn(),
    gensparkApiKey: () => 'gsk-key',
    ...overrides,
  }
}

describe('registerAiSettingsIpc', () => {
  let fetchSpy: ReturnType<typeof vi.fn>

  beforeEach(() => {
    fetchSpy = vi.fn()
    vi.stubGlobal('fetch', fetchSpy)
  })

  afterEach(() => {
    vi.restoreAllMocks()
  })

  it('registers the four provider-configuration channels', () => {
    const ipc = fakeIpc()
    registerAiSettingsIpc(makeDeps(), ipc)
    expect(ipc.handle).toHaveBeenCalledTimes(4)
    for (const channel of [
      'ai:get-settings',
      'ai:set-settings',
      'ai:ollama-models',
      'ai:test-connection',
    ]) {
      expect(ipc.handlers.has(channel)).toBe(true)
    }
  })

  it('get-settings resolves legacy single-endpoint settings into the custom provider', () => {
    const ipc = fakeIpc()
    const deps = makeDeps()
    ;(deps.readJson as ReturnType<typeof vi.fn>).mockReturnValueOnce({
      apiKey: 'legacy-key',
      model: 'legacy-model',
      baseUrl: 'https://api.openai.com/v1',
    })
    registerAiSettingsIpc(deps, ipc)
    const settings = invoke(ipc.handlers, 'ai:get-settings') as AiSettings
    expect(settings.providers.custom).toEqual({
      apiKey: 'legacy-key',
      model: 'legacy-model',
      baseUrl: 'https://api.openai.com/v1',
    })
    expect(settings.provider).toBe('genspark')
  })

  it('get-settings runs the beforeAccess guard (sheets session check)', () => {
    const ipc = fakeIpc()
    const beforeAccess = vi.fn()
    const deps = makeDeps({ beforeAccess })
    ;(deps.readJson as ReturnType<typeof vi.fn>).mockReturnValueOnce({})
    registerAiSettingsIpc(deps, ipc)
    invoke(ipc.handlers, 'ai:get-settings')
    expect(beforeAccess).toHaveBeenCalledTimes(1)
  })

  it('set-settings persists the raw settings via writeJson', () => {
    const ipc = fakeIpc()
    const writeJson = vi.fn()
    registerAiSettingsIpc(makeDeps({ writeJson }), ipc)
    const input = defaultAiSettings()
    invoke(ipc.handlers, 'ai:set-settings', input)
    expect(writeJson).toHaveBeenCalledWith('/tmp/ai-settings.json', input)
  })

  it('set-settings validates through the injected schema (sheets zod)', () => {
    const ipc = fakeIpc()
    const validateSettings = vi.fn((input: unknown) => input as AiSettings)
    const writeJson = vi.fn()
    registerAiSettingsIpc(makeDeps({ validateSettings, writeJson }), ipc)
    invoke(ipc.handlers, 'ai:set-settings', { provider: 'ollama' })
    expect(validateSettings).toHaveBeenCalledWith({ provider: 'ollama' })
    expect(writeJson).toHaveBeenCalledTimes(1)
  })

  it('ollama-models lists /api/tags models (strips the /v1 suffix)', async () => {
    const ipc = fakeIpc()
    registerAiSettingsIpc(makeDeps(), ipc)
    fetchSpy.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        models: [
          { name: 'nomic-embed-text:latest', details: { parameter_size: '137M' } },
          { name: 'qwen3:8b', modified_at: '2026-08-01' },
        ],
      }),
    })
    const result = (await invoke(
      ipc.handlers,
      'ai:ollama-models',
      'http://localhost:11434/v1',
    )) as { models: Array<{ name: string; parameterSize?: string; modifiedAt?: string }> }
    expect(fetchSpy).toHaveBeenCalledWith(
      'http://localhost:11434/api/tags',
      expect.objectContaining({ method: 'GET' }),
    )
    expect(result.models).toEqual([
      { name: 'nomic-embed-text:latest', parameterSize: '137M' },
      { name: 'qwen3:8b', modifiedAt: '2026-08-01' },
    ])
  })

  it('ollama-models surfaces a fetch failure as an error result', async () => {
    const ipc = fakeIpc()
    registerAiSettingsIpc(makeDeps(), ipc)
    fetchSpy.mockRejectedValueOnce(new Error('ECONNREFUSED'))
    const result = (await invoke(ipc.handlers, 'ai:ollama-models')) as {
      models?: unknown[]
      error?: string
    }
    expect(result.models).toEqual([])
    expect(result.error).toBeTruthy()
  })

  it('test-connection injects the gsk key for the genspark provider', async () => {
    const ipc = fakeIpc()
    registerAiSettingsIpc(makeDeps(), ipc)
    fetchSpy.mockResolvedValueOnce({ ok: true, status: 200 })
    const result = (await invoke(ipc.handlers, 'ai:test-connection', {
      provider: 'genspark',
      model: 'claude-opus-4-7',
    })) as { ok: boolean }
    expect(result.ok).toBe(true)
    const [, init] = fetchSpy.mock.calls[0]
    expect((init.headers as Record<string, string>).Authorization).toBe('Bearer gsk-key')
  })

  it('test-connection keeps an explicit key over the injected one', async () => {
    const ipc = fakeIpc()
    registerAiSettingsIpc(makeDeps(), ipc)
    fetchSpy.mockResolvedValueOnce({ ok: true, status: 200 })
    await invoke(ipc.handlers, 'ai:test-connection', {
      provider: 'genspark',
      apiKey: 'explicit',
    })
    const [, init] = fetchSpy.mock.calls[0]
    expect((init.headers as Record<string, string>).Authorization).toBe('Bearer explicit')
  })

  it('test-connection returns unknown for a missing provider id', async () => {
    const ipc = fakeIpc()
    registerAiSettingsIpc(makeDeps(), ipc)
    const result = await invoke(ipc.handlers, 'ai:test-connection', {})
    expect(result).toEqual({ ok: false, status: 'unknown' })
  })
})
