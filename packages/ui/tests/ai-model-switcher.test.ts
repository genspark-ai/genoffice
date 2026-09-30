/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { AiModelSwitcher, modelSwitcherEntries } from '../src/AiModelSwitcher'
import { AI_MODEL_SWITCHER_LABELS } from '../src/strings-ai-model-switcher'
import { AI_PROVIDERS, defaultAiSettings } from '@genoffice/ai-provider/browser'

let host: HTMLDivElement
let root: ReturnType<typeof createRoot>
beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})
afterEach(() => {
  act(() => root.unmount())
  host.remove()
  vi.restoreAllMocks()
})

describe('modelSwitcherEntries', () => {
  it('lists genspark on the gsk login and every catalog model', () => {
    const entries = modelSwitcherEntries(defaultAiSettings(), true)
    expect(entries.map((entry) => entry.meta.id)).toEqual(['genspark'])
    const genspark = AI_PROVIDERS.find((provider) => provider.id === 'genspark')!
    expect(entries[0]!.models).toEqual([...genspark.models])
  })

  it('keeps only the active row visible once the gsk session is gone', () => {
    const settings = defaultAiSettings()
    settings.providers.genspark.model = 'gpt-6-sol'
    const entries = modelSwitcherEntries(settings, false)
    expect(entries.map((entry) => entry.meta.id)).toEqual(['genspark'])
    expect(entries[0]!.models).toEqual(['gpt-6-sol'])
  })

  it('lists a configured BYOK provider next to genspark, with an off-catalog saved model kept pickable', () => {
    const settings = defaultAiSettings()
    settings.providers.kimi.apiKey = 'sk-user'
    settings.providers.kimi.model = 'kimi-k2-custom'
    const entries = modelSwitcherEntries(settings, true)
    expect(entries.map((entry) => entry.meta.id)).toEqual(['genspark', 'kimi'])
    expect(entries[1]!.models[0]).toBe('kimi-k2-custom')
    // the vendor's current catalog follows the stored id
    expect(entries[1]!.models).toContain('kimi-k3')
  })

  it('offers custom as its single saved model and skips an untouched codex', () => {
    const settings = defaultAiSettings()
    settings.providers.custom.baseUrl = 'http://localhost:11434/v1'
    settings.providers.custom.model = 'llama3'
    settings.providers.codex.cliPath = '/usr/local/bin/codex'
    const entries = modelSwitcherEntries(settings, true)
    const byId = new Map(entries.map((entry) => [entry.meta.id, entry]))
    expect(byId.get('custom')!.models).toEqual(['llama3'])
    expect(byId.get('codex')!.models).toEqual([''])
    // an untouched codex (no CLI path, no saved model) stays off the list
    settings.providers.codex.cliPath = ''
    expect(modelSwitcherEntries(settings, true).some((entry) => entry.meta.id === 'codex')).toBe(
      false,
    )
  })
})

describe('AiModelSwitcher', () => {
  it('shows the active provider and model on the trigger and switches from the fresh read', async () => {
    const settings = defaultAiSettings()
    settings.providers.anthropic.apiKey = 'sk-user'
    settings.provider = 'anthropic'
    settings.providers.anthropic.model = 'claude-sonnet-5'
    const fresh = { ...settings, gskToolsEnabled: false }
    const reload = vi.fn(async () => fresh)
    const onSwitch = vi.fn()
    act(() =>
      root.render(
        createElement(AiModelSwitcher, {
          lang: 'en',
          settings,
          gskLoggedIn: true,
          reload,
          onSwitch,
        }),
      ),
    )
    const trigger = host.querySelector<HTMLButtonElement>('.gs-dd-btn')!
    expect(trigger.getAttribute('aria-label')).toBe(AI_MODEL_SWITCHER_LABELS.en.switch)
    expect(trigger.textContent).toContain('Claude')
    expect(trigger.textContent).toContain('claude-sonnet-5')

    act(() => trigger.click())
    const items = [...host.querySelectorAll<HTMLButtonElement>('.gs-dd-item')]
    // one row per model: the genspark catalog plus the configured Claude catalog
    const genspark = AI_PROVIDERS.find((provider) => provider.id === 'genspark')!
    const anthropic = AI_PROVIDERS.find((provider) => provider.id === 'anthropic')!
    expect(items.length).toBe(genspark.models.length + anthropic.models.length)
    const target = items.find((item) => item.dataset.value === 'anthropic|claude-opus-4-7')!
    await act(async () => target.click())
    expect(onSwitch).toHaveBeenCalledTimes(1)
    const next = onSwitch.mock.calls[0]![0] as typeof settings
    expect(next.provider).toBe('anthropic')
    expect(next.providers.anthropic.model).toBe('claude-opus-4-7')
    // the write base is the fresh read, not the stale prop: the marker the prop
    // does not carry proves onSwitch received the reloaded settings
    expect(next.gskToolsEnabled).toBe(false)
  })

  it('re-picking the active row is a no-op', async () => {
    const settings = defaultAiSettings()
    const onSwitch = vi.fn()
    const reload = vi.fn(async () => settings)
    act(() =>
      root.render(
        createElement(AiModelSwitcher, {
          lang: 'en',
          settings,
          gskLoggedIn: true,
          reload,
          onSwitch,
        }),
      ),
    )
    const trigger = host.querySelector<HTMLButtonElement>('.gs-dd-btn')!
    act(() => trigger.click())
    const active = host.querySelector<HTMLButtonElement>('.gs-dd-item.selected')!
    await act(async () => active.click())
    expect(onSwitch).not.toHaveBeenCalled()
  })

  it('says when nothing besides the active selection is configured', () => {
    const settings = defaultAiSettings()
    settings.providers.genspark.model = 'gpt-6-sol'
    const entries = modelSwitcherEntries(settings, false)
    expect(entries).toHaveLength(1)
    // and the empty-state copy stays available to the header for that case
    expect(AI_MODEL_SWITCHER_LABELS.en.empty).toBeTruthy()
  })
})
