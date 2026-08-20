// AiSettingsDialog Ollama UX: connection status, model discovery, TTL cache,
// forced refresh, and the model-missing notice — all against mocked discovery
// (no real Ollama, no network).
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import {
  AiSettingsDialog,
  type AiSettingsDialogStrings,
} from '../src/AiSettingsDialog'
import { AI_PROVIDERS, type AiSettings, type OllamaModelsResult } from '@genoffice/ai-provider'

function makeStrings(overrides: Partial<AiSettingsDialogStrings> = {}): AiSettingsDialogStrings {
  return {
    aiSettingsProvider: 'Provider',
    aiSettingsApiKey: 'API key',
    aiSettingsApiKeyHint: 'Optional for local models',
    aiSettingsBaseUrl: 'Base URL',
    aiSettingsDetectedModels: 'Detected models',
    aiSettingsRefresh: 'Refresh',
    aiSettingsNoModel: 'No models found',
    aiSettingsModelMissing: 'The selected model is not installed',
    aiSettingsTestFail: 'Connection failed',
    aiSettingsCancel: 'Cancel',
    aiSettingsSave: 'Save',
    aiSettingsModel: 'Model',
    aiSettingsGensparkLogin: 'Sign in',
    aiSettingsGensparkConnected: 'Connected',
    aiSettingsGensparkDisconnected: 'Disconnected',
    aiSettingsOllamaBaseUrlHint: 'Local endpoint',
    aiSettingsTestButton: 'Test connection',
    aiSettingsTestConnected: 'Connected',
    aiSettingsTestNotRunning: 'Ollama is not running',
    aiSettingsTestRefused: 'Connection refused',
    aiSettingsTestInvalid: 'Invalid endpoint',
    aiSettingsTestAuth: 'Authentication failed',
    aiSettingsTestTimeout: 'Timed out',
    aiSettingsTestFailed: 'Unknown error',
    ...overrides,
  }
}

const strings = makeStrings()

let baseUrlCounter = 0
function makeSettings(overrides: Partial<AiSettings> = {}): AiSettings {
  // unique base URL per fixture so the module-level discovery TTL cache never
  // leaks between tests
  baseUrlCounter += 1
  const providers = Object.fromEntries(
    AI_PROVIDERS.map((p) => [
      p.id,
      { apiKey: '', model: p.defaultModel, baseUrl: p.defaultBaseUrl ?? '' },
    ]),
  ) as AiSettings['providers']
  return {
    provider: 'ollama',
    providers: {
      ...providers,
      ollama: {
        apiKey: '',
        model: '',
        baseUrl: `http://127.0.0.1:${5000 + baseUrlCounter}`,
      },
    },
    ...overrides,
  }
}

let container: HTMLElement
let root: Root

function mount(props: Parameters<typeof AiSettingsDialog>[0]) {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  act(() => root.render(createElement(AiSettingsDialog, props)))
}

function unmount() {
  act(() => root.unmount())
  container.remove()
}

/** flush the discovery promise + state updates */
async function flush() {
  await act(async () => {
    await Promise.resolve()
  })
}

function dialogProps(overrides: Partial<Parameters<typeof AiSettingsDialog>[0]> = {}) {
  return {
    settings: makeSettings(),
    strings,
    listOllamaModels: async () => ({ models: [] }) as OllamaModelsResult,
    onSettingsChange: vi.fn(),
    onClose: vi.fn(),
    ...overrides,
  }
}

function statusChip() {
  return container.querySelector('.ai-settings-ollama-status')
}

/** the model <select> (index 0 is the provider select) */
function modelSelect() {
  return container.querySelectorAll<HTMLSelectElement>('select')[1]
}

afterEach(() => {
  unmount()
  vi.restoreAllMocks()
})

describe('AiSettingsDialog · Ollama connection state', () => {
  it('shows a loading state while discovery is in flight', async () => {
    let resolve!: (r: OllamaModelsResult) => void
    const list = vi.fn(() => new Promise<OllamaModelsResult>((res) => (resolve = res)))
    mount(dialogProps({ listOllamaModels: list }))
    await flush()
    expect(statusChip()).toBeTruthy()
    expect(statusChip()!.textContent).toContain(strings.aiSettingsTestButton)
    resolve({ models: [{ name: 'qwen3:8b' }] })
    await flush()
  })

  it('shows a friendly message when Ollama is not running', async () => {
    const list = vi.fn(async () => ({ models: [], error: 'fetch failed: ECONNREFUSED' }))
    mount(dialogProps({ listOllamaModels: list }))
    await flush()
    expect(statusChip()!.textContent).toBe(strings.aiSettingsTestNotRunning)
    // user-facing text only — no stack traces or raw error dumps
    expect(container.textContent).not.toContain('ECONNREFUSED')
  })

  it('classifies an HTTP-level failure as an invalid endpoint', async () => {
    const list = vi.fn(async () => ({ models: [], error: 'Ollama returned 500' }))
    mount(dialogProps({ listOllamaModels: list }))
    await flush()
    expect(statusChip()!.textContent).toBe(strings.aiSettingsTestInvalid)
  })

  it('reports connected even when zero models are installed', async () => {
    const list = vi.fn(async () => ({ models: [] }))
    mount(dialogProps({ listOllamaModels: list }))
    await flush()
    expect(statusChip()!.textContent).toBe(strings.aiSettingsTestConnected)
    const select = modelSelect()
    expect(select!.textContent).toContain(strings.aiSettingsNoModel)
  })
})

describe('AiSettingsDialog · model discovery', () => {
  it('lists discovered models with sizes', async () => {
    const list = vi.fn(async () => ({
      models: [
        { name: 'qwen3:8b', parameterSize: '8B' },
        { name: 'phi4:14b', parameterSize: '14B' },
      ],
    }))
    mount(dialogProps({ listOllamaModels: list }))
    await flush()
    const select = modelSelect()
    expect(select!.textContent).toContain('qwen3:8b (8B)')
    expect(select!.textContent).toContain('phi4:14b (14B)')
  })

  it('auto-selects the first discovered model when none is configured', async () => {
    const list = vi.fn(async () => ({ models: [{ name: 'phi4:14b' }] }))
    mount(dialogProps({ listOllamaModels: list }))
    await flush()
    const select = modelSelect()
    expect(select!.value).toBe('phi4:14b')
  })

  it('notifies when the selected model is no longer installed (no silent switch)', async () => {
    const settings = makeSettings({
      providers: {
        ...makeSettings().providers,
        ollama: { apiKey: '', model: 'llama3.2:8b', baseUrl: 'http://127.0.0.1:9999' },
      },
    })
    const list = vi.fn(async () => ({ models: [{ name: 'qwen3:8b' }] }))
    mount(dialogProps({ settings, listOllamaModels: list }))
    await flush()
    // the stale model stays selectable (so the user can keep the saved setting)
    const select = modelSelect()
    expect(select!.value).toBe('llama3.2:8b')
    expect(container.textContent).toContain(strings.aiSettingsModelMissing)
  })

  it('does not show the model-missing notice when the model is present', async () => {
    const list = vi.fn(async () => ({ models: [{ name: 'qwen3:8b' }] }))
    mount(dialogProps({ listOllamaModels: list }))
    await flush()
    expect(container.textContent).not.toContain(strings.aiSettingsModelMissing)
  })
})

describe('AiSettingsDialog · discovery caching and refresh', () => {
  it('caches discovery within the TTL and refetches after expiry', async () => {
    vi.useFakeTimers()
    const list = vi
      .fn()
      .mockResolvedValueOnce({ models: [{ name: 'qwen3:8b' }] })
      .mockResolvedValueOnce({ models: [{ name: 'qwen3:8b' }, { name: 'phi4:14b' }] })
    const props = dialogProps({ listOllamaModels: list })
    mount(props)
    await flush()
    expect(list).toHaveBeenCalledTimes(1)

    // reopening within the TTL must not hit /api/tags again
    unmount()
    mount(props)
    await flush()
    expect(list).toHaveBeenCalledTimes(1)

    // after the TTL elapses the next open refetches
    await act(async () => {
      vi.advanceTimersByTime(11_000)
    })
    unmount()
    mount(props)
    await flush()
    expect(list).toHaveBeenCalledTimes(2)
    const select = modelSelect()
    expect(select!.textContent).toContain('phi4:14b')
    vi.useRealTimers()
  })

  it('refresh button forces a fresh probe', async () => {
    const list = vi
      .fn()
      .mockResolvedValueOnce({ models: [{ name: 'qwen3:8b' }] })
      .mockResolvedValueOnce({ models: [{ name: 'phi4:14b' }] })
    mount(dialogProps({ listOllamaModels: list }))
    await flush()
    expect(list).toHaveBeenCalledTimes(1)

    const refresh = [...container.querySelectorAll('button')].find(
      (b) => b.textContent === strings.aiSettingsRefresh,
    )!
    act(() => refresh.click())
    await flush()
    expect(list).toHaveBeenCalledTimes(2)
    const select = modelSelect()
    expect(select!.textContent).toContain('phi4:14b')
  })

  it('does not cache failures — the next open retries', async () => {
    const list = vi
      .fn()
      .mockResolvedValueOnce({ models: [], error: 'fetch failed: ECONNREFUSED' })
      .mockResolvedValueOnce({ models: [{ name: 'qwen3:8b' }] })
    const props = dialogProps({ listOllamaModels: list })
    mount(props)
    await flush()
    expect(statusChip()!.textContent).toBe(strings.aiSettingsTestNotRunning)

    unmount()
    mount(props)
    await flush()
    expect(list).toHaveBeenCalledTimes(2)
    expect(statusChip()!.textContent).toBe(strings.aiSettingsTestConnected)
  })
})
