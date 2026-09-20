/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { defaultAiSettings } from '@genoffice/ai-provider'
import type { AiSettings, CodexModelCatalog } from '@genoffice/ai-provider'
import type { HomeApi } from '../src/shared/home-api'
import { LocaleProvider } from '../src/renderer/src/locale'
import { SettingsModal } from '../src/renderer/src/SettingsModal'

/**
 * The settings-screen half of the custom-endpoint model list is one effect, and
 * these cases pin what that effect is responsible for, in the real modal:
 *
 *   1. It writes the CATALOG and never the settings. `updateConfig` rebuilds the
 *      settings from the render that captured it, so an async write would revert
 *      whatever the user did while the probe was in flight.
 *   2. No default is guessed. The picker opens with nothing selected.
 *   3. Requests are debounced and superseded replies are dropped, so typing an
 *      address is one request and a slow answer from the previous server can
 *      never overwrite a fast answer from the current one.
 *   4. It only probes for the provider actually selected.
 */

const BASE_URL = 'http://127.0.0.1:8800/v1'
const OTHER_URL = 'http://127.0.0.1:11434/v1'
/** the pane's debounce window for the custom model list */
const DEBOUNCE_MS = 400

const actEnvironment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
actEnvironment.IS_REACT_ACT_ENVIRONMENT = true

let host: HTMLDivElement
let root: Root
/** every getCustomModels call, resolvable by hand so replies can be ordered */
let calls: Array<{
  baseUrl: string
  apiKey: string | undefined
  resolve: (catalog: CodexModelCatalog) => void
}>
let setAiSettings: ReturnType<typeof vi.fn>

beforeEach(() => {
  vi.useFakeTimers()
  calls = []
  setAiSettings = vi.fn(async () => undefined)
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  vi.unstubAllGlobals()
  vi.useRealTimers()
})

function storedSettings(
  options: { provider?: AiSettings['provider']; baseUrl?: string; model?: string } = {},
): AiSettings {
  const settings = defaultAiSettings()
  settings.provider = options.provider ?? 'custom'
  settings.providers.custom = {
    apiKey: '',
    model: options.model ?? '',
    baseUrl: options.baseUrl ?? BASE_URL,
  }
  return settings
}

function installApi(settings: AiSettings): void {
  window.aiOffice = {
    getTheme: async () => 'system',
    getDefaultSaveDir: async () => '',
    getAnalyticsEnabled: async () => true,
    setAnalyticsEnabled: async () => true,
    getAiPanelPrefs: async () => ({ fontSize: 'medium', customFontSize: 14, spellcheck: true }),
    setAiPanelPrefs: async (patch: unknown) => patch,
    getUpdateChannel: async () => 'stable',
    getAppVersion: async () => '1.0.0',
    githubStars: async () => null,
    getAiProviders: () => [
      {
        id: 'genspark',
        label: 'Genspark',
        models: [],
        defaultModel: '',
        keyPlaceholder: 'k',
        defaultBaseUrl: '',
      },
      {
        id: 'custom',
        label: 'Custom',
        models: [],
        defaultModel: '',
        keyPlaceholder: 'API Key',
        needsBaseUrl: true,
        defaultBaseUrl: '',
      },
    ],
    getAiSettings: async () => settings,
    setAiSettings,
    getCustomModels: (baseUrl: string, apiKey?: string) =>
      new Promise<CodexModelCatalog>((resolve) => calls.push({ baseUrl, apiKey, resolve })),
  } as unknown as HomeApi
}

async function openAiPane(): Promise<void> {
  await act(async () => {
    root.render(
      createElement(
        LocaleProvider,
        { initial: 'en' },
        createElement(SettingsModal, {
          status: null,
          loggingOut: false,
          loginWaiting: false,
          loginUrl: null,
          urlCopied: false,
          onOpenLoginUrl: vi.fn(),
          onCopyLoginUrl: vi.fn(),
          onClose: vi.fn(),
          onLogin: vi.fn(),
          onLogout: vi.fn(),
        }),
      ),
    )
    await Promise.resolve()
  })
  const nav = Array.from(host.querySelectorAll<HTMLButtonElement>('.set-nav-item')).find((button) =>
    button.textContent?.includes('AI Model'),
  )
  await click(nav!)
}

async function click(button: HTMLButtonElement): Promise<void> {
  await act(async () => {
    button.click()
    await Promise.resolve()
  })
}

/** run the debounce out and let any resolved reply be applied */
async function tick(ms = DEBOUNCE_MS): Promise<void> {
  await act(async () => {
    await vi.advanceTimersByTimeAsync(ms)
  })
}

/** answer the nth outstanding call */
async function answer(index: number, models: string[]): Promise<void> {
  await act(async () => {
    calls[index]!.resolve({ models, defaultModel: '' })
    await Promise.resolve()
  })
}

/** type into a controlled React input the way a user would */
async function type(input: HTMLInputElement, value: string): Promise<void> {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
  await act(async () => {
    setter.call(input, value)
    input.dispatchEvent(new Event('input', { bubbles: true }))
    await Promise.resolve()
  })
}

function field(id: string): HTMLInputElement {
  return host.querySelector<HTMLInputElement>(`#${id}`)!
}

function modelBox(): HTMLInputElement | null {
  return host.querySelector<HTMLInputElement>('#set-ai-model')
}

function picker(label: string): HTMLButtonElement | null {
  return host.querySelector<HTMLButtonElement>(`.gs-dd-btn[aria-label="${label}"]`)
}

/** the model ids the picker offers, in order */
async function modelOptions(): Promise<string[]> {
  const trigger = picker('Model')!
  await click(trigger)
  const options = Array.from(host.querySelectorAll<HTMLButtonElement>('.gs-dd-pop [role="option"]'))
  const labels = options.map((option) => option.getAttribute('aria-label') ?? option.textContent!)
  await click(trigger)
  return labels
}

async function pickModel(model: string): Promise<void> {
  await click(picker('Model')!)
  const option = Array.from(
    host.querySelectorAll<HTMLButtonElement>('.gs-dd-pop [role="option"]'),
  ).find((button) => button.textContent === model)
  await click(option!)
}

async function pickProvider(label: string): Promise<void> {
  await click(picker('Provider')!)
  const option = Array.from(
    host.querySelectorAll<HTMLButtonElement>('.gs-dd-pop [role="option"]'),
  ).find((button) => button.textContent?.includes(label))
  await click(option!)
}

// ── 1. it never writes settings ───────────────────────────

describe('the refresh touches the catalog and nothing else', () => {
  it('does not revert a key typed while the probe is in flight', async () => {
    installApi(storedSettings())
    await openAiPane()
    await tick()
    expect(calls).toHaveLength(1) // the first probe is out, and holding

    await type(field('set-ai-key'), 'sk-typed-while-waiting')
    expect(field('set-ai-key').value).toBe('sk-typed-while-waiting')

    // the first server finally answers — superseded, so it must land nowhere
    await answer(0, ['stale'])
    expect(field('set-ai-key').value).toBe('sk-typed-while-waiting')
    expect(modelBox()).not.toBeNull()

    // the re-probe carrying the typed key is what populates the picker
    await tick()
    expect(calls).toHaveLength(2)
    await answer(1, ['alpha', 'beta'])

    expect(field('set-ai-key').value).toBe('sk-typed-while-waiting')
    expect(modelBox()).toBeNull() // a picker, not a text box
  })

  it('does not revert an unrelated setting changed while the probe is in flight', async () => {
    // The output cap is not one of the effect's dependencies, so editing it does
    // not re-run the effect or cancel the probe. If the reply wrote settings, it
    // would write the snapshot captured before this edit and silently undo it —
    // precisely the fault the catalog-only rule exists to prevent.
    installApi(storedSettings())
    await openAiPane()
    await tick()
    expect(calls).toHaveLength(1)

    const cap = field('set-ai-max-tokens')
    await type(cap, '9216')
    await act(async () => {
      cap.dispatchEvent(new FocusEvent('focusout', { bubbles: true }))
      await Promise.resolve()
    })
    expect(field('set-ai-max-tokens').value).toBe('9216')

    await answer(0, ['alpha', 'beta'])
    expect(field('set-ai-max-tokens').value).toBe('9216')
    expect(modelBox()).toBeNull()
  })

  it('never persists settings of its own accord', async () => {
    installApi(storedSettings())
    await openAiPane()
    await tick()
    await answer(0, ['alpha', 'beta'])
    expect(setAiSettings).not.toHaveBeenCalled()
  })

  it('opens the picker with nothing selected rather than a guessed model', async () => {
    installApi(storedSettings())
    await openAiPane()
    await tick()
    await answer(0, ['text-embedding-nomic-embed-text-v1.5', 'chat'])
    // an embedding model must never be silently adopted as the chat model
    expect(picker('Model')!.textContent).toBe('')
    expect(host.textContent).not.toContain('text-embedding-nomic-embed-text-v1.5')
  })
})

// ── 2. the fold itself ────────────────────────────────────

describe('the fold', () => {
  it('turns the model box into a picker listing what the endpoint advertises', async () => {
    installApi(storedSettings())
    await openAiPane()
    expect(modelBox()).not.toBeNull()
    await tick()
    await answer(0, ['alpha', 'beta'])
    expect(modelBox()).toBeNull()
    expect(await modelOptions()).toEqual(['alpha', 'beta'])
  })

  it('pins a hand-typed model to the top when the live list lacks it', async () => {
    installApi(storedSettings({ model: 'my-local-gguf' }))
    await openAiPane()
    await tick()
    await answer(0, ['alpha', 'beta'])
    expect(await modelOptions()).toEqual(['my-local-gguf', 'alpha', 'beta'])
  })

  it('does not duplicate a stored model the live list already has', async () => {
    installApi(storedSettings({ model: 'beta' }))
    await openAiPane()
    await tick()
    await answer(0, ['alpha', 'beta'])
    expect(await modelOptions()).toEqual(['alpha', 'beta'])
  })

  it('sends the stored key with the request', async () => {
    installApi(storedSettings())
    await openAiPane()
    await type(field('set-ai-key'), 'sk-stored')
    await tick()
    expect(calls.map((call) => [call.baseUrl, call.apiKey])).toEqual([[BASE_URL, 'sk-stored']])
  })

  it('leaves every other provider entry alone', async () => {
    installApi(storedSettings())
    await openAiPane()
    await tick()
    await answer(0, ['alpha', 'beta'])
    expect(modelBox()).toBeNull()

    await pickProvider('Genspark')
    expect(modelBox()).not.toBeNull() // genspark still has no model list
  })
})

// ── 3. when it probes at all ──────────────────────────────

describe('scope', () => {
  it('does not probe a saved endpoint while another provider is selected', async () => {
    installApi(storedSettings({ provider: 'genspark' }))
    await openAiPane()
    await tick()
    expect(calls).toEqual([])
  })

  it('probes as soon as the provider dropdown switches to the endpoint', async () => {
    installApi(storedSettings({ provider: 'genspark' }))
    await openAiPane()
    await tick()
    expect(calls).toEqual([])

    await pickProvider('Custom')
    await tick()
    await answer(0, ['alpha', 'beta'])
    expect(modelBox()).toBeNull()
  })

  it('does nothing without an address', async () => {
    installApi(storedSettings({ baseUrl: '' }))
    await openAiPane()
    await tick()
    expect(calls).toEqual([])
    expect(modelBox()).not.toBeNull()
  })
})

// ── 4. request discipline ─────────────────────────────────

describe('request discipline', () => {
  it('collapses a burst of keystrokes in the address field into one request', async () => {
    installApi(storedSettings({ baseUrl: '' }))
    await openAiPane()
    for (const partial of ['http://127.0.0.1:8', 'http://127.0.0.1:88', BASE_URL]) {
      await type(field('set-ai-base-url'), partial)
      await tick(100) // each keystroke well inside the debounce window
    }
    expect(calls).toEqual([])
    await tick()
    expect(calls.map((call) => call.baseUrl)).toEqual([BASE_URL])
  })

  it('an edit that changes nothing it depends on costs no request', async () => {
    installApi(storedSettings())
    await openAiPane()
    await tick()
    await answer(0, ['alpha', 'beta'])
    for (const value of ['9216', '8192', '9216']) {
      await type(field('set-ai-max-tokens'), value)
      await tick()
    }
    expect(calls).toHaveLength(1)
    expect(modelBox()).toBeNull()
  })

  it("drops a superseded reply — B's list wins and A never lands", async () => {
    installApi(storedSettings())
    await openAiPane()
    await tick()
    expect(calls).toHaveLength(1) // A is out, holding

    await type(field('set-ai-base-url'), OTHER_URL)
    await tick()
    expect(calls).toHaveLength(2) // B is out too

    await answer(1, ['from-B'])
    expect(await modelOptions()).toEqual(['from-B'])

    await answer(0, ['from-A']) // A finally answers, far too late
    expect(await modelOptions()).toEqual(['from-B'])
  })

  it('clears the picker the moment the address changes, before the new server answers', async () => {
    installApi(storedSettings())
    await openAiPane()
    await tick()
    await answer(0, ['alpha', 'beta'])
    expect(modelBox()).toBeNull()

    await type(field('set-ai-base-url'), OTHER_URL)
    expect(modelBox()).not.toBeNull() // the text box is back immediately
  })

  it('clears the picker when the address is deleted outright', async () => {
    // there is no server to ask any more, so a model must not be pickable —
    // and must certainly not be saved — against an address that is gone
    installApi(storedSettings())
    await openAiPane()
    await tick()
    await answer(0, ['alpha', 'beta'])
    expect(modelBox()).toBeNull()

    await type(field('set-ai-base-url'), '')
    expect(modelBox()).not.toBeNull()
    await tick()
    expect(calls).toHaveLength(1) // and nothing is asked of the empty address
  })

  it('keeps the list when the same address fails — a blip must not wipe the picker', async () => {
    installApi(storedSettings())
    await openAiPane()
    await tick()
    await answer(0, ['alpha', 'beta'])

    // only the key changes, so the address is unchanged; the server refuses
    await type(field('set-ai-key'), 'sk-wrong')
    expect(modelBox()).toBeNull()
    await tick()
    await answer(1, [])
    expect(modelBox()).toBeNull()
    expect(await modelOptions()).toEqual(['alpha', 'beta'])
  })

  it('choosing a model from the picker does not send another request', async () => {
    installApi(storedSettings())
    await openAiPane()
    await tick()
    await answer(0, ['alpha', 'beta'])
    await pickModel('beta')
    await tick()
    expect(calls).toHaveLength(1)
    expect(picker('Model')!.textContent).toBe('beta')
  })

  it('drops a reply that lands after the provider has been switched away', async () => {
    installApi(storedSettings())
    await openAiPane()
    await tick()
    expect(calls).toHaveLength(1) // out, and holding

    await pickProvider('Genspark')
    await answer(0, ['alpha', 'beta']) // the endpoint answers too late to matter

    await pickProvider('Custom')
    expect(modelBox()).not.toBeNull() // the abandoned reply folded nothing in
  })
})
