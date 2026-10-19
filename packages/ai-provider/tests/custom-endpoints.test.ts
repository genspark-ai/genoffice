import { describe, expect, it } from 'vitest'
import {
  activeCustomEndpoint,
  customEndpointLabel,
  removeCustomEndpoint,
  resolveCustomEndpoints,
  selectCustomEndpoint,
  upsertCustomEndpoint,
} from '../src/custom-endpoints'
import { sanitizeAiSettings } from '../src/ai-settings-guard'
import { defaultAiSettings, resolveAiSettings } from '../src/providers'
import type { AiCustomEndpoint } from '../src/types'

const ollama: AiCustomEndpoint = {
  id: 'a',
  name: 'Ollama',
  baseUrl: 'http://localhost:11434/v1',
  apiKey: '',
  model: 'llama3',
}
const gateway: AiCustomEndpoint = {
  id: 'b',
  name: '',
  baseUrl: 'https://gw.example.com/v1',
  apiKey: 'k',
  model: 'gpt-x',
  models: ['gpt-x', 'gpt-y'],
}

describe('resolveCustomEndpoints', () => {
  it('seeds the list from a pre-list custom slot and mirrors it back', () => {
    const settings = defaultAiSettings()
    settings.providers.custom = { apiKey: 'k', model: 'm', baseUrl: 'http://h/v1' }
    const next = resolveCustomEndpoints(settings)
    expect(next.customEndpoints).toHaveLength(1)
    expect(next.customEndpoints![0]).toMatchObject({
      apiKey: 'k',
      model: 'm',
      baseUrl: 'http://h/v1',
    })
    expect(next.customEndpoint).toBe(next.customEndpoints![0]!.id)
    expect(next.providers.custom).toEqual({ apiKey: 'k', model: 'm', baseUrl: 'http://h/v1' })
  })

  it('leaves an unconfigured file without a list', () => {
    const next = resolveCustomEndpoints(defaultAiSettings())
    expect(next.customEndpoints).toBeUndefined()
    expect(next.customEndpoint).toBeUndefined()
  })

  it('falls back to the first entry when the active id is stale', () => {
    const next = resolveCustomEndpoints({
      ...defaultAiSettings(),
      customEndpoints: [ollama, gateway],
      customEndpoint: 'gone',
    })
    expect(next.customEndpoint).toBe('a')
    expect(next.providers.custom.baseUrl).toBe(ollama.baseUrl)
  })

  it('drops duplicate and id-less entries', () => {
    const next = resolveCustomEndpoints({
      ...defaultAiSettings(),
      customEndpoints: [ollama, { ...gateway, id: 'a' }, { ...gateway, id: '' }],
    })
    expect(next.customEndpoints!.map((e) => e.id)).toEqual(['a'])
  })

  it('runs as part of resolveAiSettings', () => {
    const stored = {
      provider: 'custom',
      providers: { custom: { apiKey: '', model: 'old', baseUrl: 'http://old/v1' } },
      customEndpoints: [ollama, gateway],
      customEndpoint: 'b',
    }
    const settings = resolveAiSettings(stored, defaultAiSettings())
    expect(settings.providers.custom).toEqual({
      apiKey: 'k',
      model: 'gpt-x',
      baseUrl: gateway.baseUrl,
    })
    expect(activeCustomEndpoint(settings)?.id).toBe('b')
  })
})

describe('select / upsert / remove', () => {
  const base = resolveCustomEndpoints({
    ...defaultAiSettings(),
    customEndpoints: [ollama, gateway],
    customEndpoint: 'a',
  })

  it('switching the active endpoint swaps the slot', () => {
    const next = selectCustomEndpoint(base, 'b')
    expect(next.providers.custom.model).toBe('gpt-x')
    expect(base.providers.custom.model).toBe('llama3')
  })

  it('upsert edits in place and appends new ids as active', () => {
    const edited = upsertCustomEndpoint(base, { ...ollama, model: 'mistral' })
    expect(edited.customEndpoints!.map((e) => e.model)).toEqual(['mistral', 'gpt-x'])
    expect(edited.providers.custom.model).toBe('mistral')
    const added = upsertCustomEndpoint(base, { ...gateway, id: 'c', name: 'Third' })
    expect(added.customEndpoints).toHaveLength(3)
    expect(added.customEndpoint).toBe('c')
  })

  it('removing the active entry activates the next one, removing the last clears the slot', () => {
    const one = removeCustomEndpoint(base, 'a')
    expect(one.customEndpoint).toBe('b')
    expect(one.providers.custom.baseUrl).toBe(gateway.baseUrl)
    const none = removeCustomEndpoint(one, 'b')
    expect(none.customEndpoints).toBeUndefined()
    expect(none.providers.custom).toEqual({ apiKey: '', model: '', baseUrl: '' })
  })

  it('labels fall back to the host', () => {
    expect(customEndpointLabel(ollama, 'Custom')).toBe('Ollama')
    expect(customEndpointLabel(gateway, 'Custom')).toBe('gw.example.com')
    expect(customEndpointLabel({ ...gateway, baseUrl: '' }, 'Custom')).toBe('Custom')
  })
})

describe('sanitizeAiSettings', () => {
  it('keeps the endpoint list, drops bad base URLs and oversized ids', () => {
    const out = sanitizeAiSettings({
      provider: 'custom',
      providers: {},
      customEndpoints: [
        gateway,
        { ...ollama, id: 'x'.repeat(100), baseUrl: 'javascript:alert(1)' },
        { name: 'no id' },
      ],
      customEndpoint: 'b',
    })
    expect(out?.customEndpoints).toHaveLength(2)
    expect(out?.customEndpoints![0]).toEqual(gateway)
    expect(out?.customEndpoints![1]!.id).toHaveLength(64)
    expect(out?.customEndpoints![1]!.baseUrl).toBe('')
    expect(out?.customEndpoint).toBe('b')
  })

  it('omits the list when nothing survives', () => {
    const out = sanitizeAiSettings({ provider: 'custom', providers: {}, customEndpoints: 'nope' })
    expect(out?.customEndpoints).toBeUndefined()
  })
})
