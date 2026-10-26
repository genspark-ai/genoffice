import { describe, expect, it } from 'vitest'
import { defaultAiSettings } from '@genoffice/ai-provider'
import { aiSetupNeeded } from '../src/renderer/src/ai-setup-state'

describe('aiSetupNeeded', () => {
  it('shows the entry while signed out with no provider configured', () => {
    expect(aiSetupNeeded(defaultAiSettings(), false)).toBe(true)
  })

  it('hides the entry once Genspark is signed in', () => {
    expect(aiSetupNeeded(defaultAiSettings(), true)).toBe(false)
  })

  it('hides the entry once a BYOK provider is usable', () => {
    const settings = defaultAiSettings()
    settings.providers.anthropic = { apiKey: 'sk-ant-...', model: 'claude-sonnet-5' }
    expect(aiSetupNeeded(settings, false)).toBe(false)
  })

  it('keeps the entry for a half-filled BYOK selection', () => {
    const settings = defaultAiSettings()
    // stored selection but no key: activeProvider falls back to genspark
    settings.provider = 'anthropic'
    settings.providers.anthropic = { apiKey: '', model: 'claude-sonnet-5' }
    expect(aiSetupNeeded(settings, false)).toBe(true)
  })

  it('hides the entry for a usable custom endpoint', () => {
    const settings = defaultAiSettings()
    settings.customEndpoints = [
      {
        id: 'ep1',
        name: 'Ollama',
        baseUrl: 'http://localhost:11434/v1',
        apiKey: '',
        model: 'llama',
      },
    ]
    expect(aiSetupNeeded(settings, false)).toBe(false)
  })

  it('treats an unknown settings load as not needing the entry', () => {
    expect(aiSetupNeeded(null, false)).toBe(false)
  })
})
