import { describe, expect, it } from 'vitest'
import { AI_PROVIDERS, defaultAiSettings } from '@genoffice/ai-provider'
import {
  aiModelPickerGroups,
  aiModelPickerSelection,
  withAiModelSelection,
} from '../src/ai-model-picker-options'

const anthropicModels = AI_PROVIDERS.find((m) => m.id === 'anthropic')!.models

describe('aiModelPickerGroups', () => {
  it('lists nothing while signed out with no keys', () => {
    expect(aiModelPickerGroups(defaultAiSettings(), false)).toEqual([])
  })

  it('lists Genspark only when signed in', () => {
    const ids = aiModelPickerGroups(defaultAiSettings(), true).map((g) => g.id)
    expect(ids).toEqual(['genspark'])
  })

  it('adds a vendor once its key is set and puts an off-catalog model first', () => {
    const settings = defaultAiSettings()
    settings.providers.anthropic = { apiKey: 'k', model: 'claude-next' }
    const groups = aiModelPickerGroups(settings, false)
    expect(groups.map((g) => g.id)).toEqual(['anthropic'])
    expect(groups[0]!.models).toEqual(['claude-next', ...anthropicModels])
  })

  it('needs a base URL for custom endpoints', () => {
    const settings = defaultAiSettings()
    settings.providers.custom = { apiKey: '', model: 'llama', baseUrl: '' }
    expect(aiModelPickerGroups(settings, false)).toEqual([])
    settings.providers.custom = { apiKey: '', model: 'llama', baseUrl: 'http://localhost:11434/v1' }
    expect(aiModelPickerGroups(settings, false).map((g) => g.id)).toEqual(['custom'])
  })
})

describe('selection round trip', () => {
  it('switches provider and model in one write and reads back', () => {
    const settings = defaultAiSettings()
    settings.providers.anthropic = { apiKey: 'k', model: anthropicModels[0]! }
    const next = withAiModelSelection(settings, {
      provider: 'anthropic',
      model: anthropicModels[1]!,
    })
    expect(next.provider).toBe('anthropic')
    expect(next.providers.anthropic.apiKey).toBe('k')
    expect(aiModelPickerSelection(next)).toEqual({
      provider: 'anthropic',
      model: anthropicModels[1],
    })
    expect(settings.providers.anthropic.model).toBe(anthropicModels[0])
  })

  it('falls back to Genspark when the stored provider is unusable', () => {
    const settings = defaultAiSettings()
    settings.provider = 'anthropic'
    expect(aiModelPickerSelection(settings).provider).toBe('genspark')
  })
})

describe('CLI vendors', () => {
  it('lists Codex only after the settings page stored a path or model', () => {
    const settings = defaultAiSettings()
    expect(aiModelPickerGroups(settings, false)).toEqual([])
    settings.providers.codex = { apiKey: '', model: '', cliPath: '/usr/local/bin/codex' }
    expect(aiModelPickerGroups(settings, false)).toEqual([
      { id: 'codex', label: 'Codex CLI', models: [] },
    ])
  })
})
