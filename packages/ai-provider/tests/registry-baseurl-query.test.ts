import { describe, expect, it } from 'vitest'
import { AI_PROVIDER_ADAPTERS } from '../src/registry'
import type { AiProviderConfig } from '../src/types'

function config(model: string, baseUrl?: string): AiProviderConfig {
  return { apiKey: 'k', model, baseUrl }
}

describe('opencode endpoint base URL with query string', () => {
  const zen = AI_PROVIDER_ADAPTERS['opencode-zen']

  it('joins /v1 before the query string (openai-compatible route)', () => {
    expect(
      zen.resolveEndpoint(config('glm-5.2', 'https://gateway.example/v1?api-version=2024-01-01')),
    ).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://gateway.example/v1?api-version=2024-01-01',
    })
  })

  it('joins /v1 before the query string (gemini route)', () => {
    expect(
      zen.resolveEndpoint(
        config('gemini-3.7-flash', 'https://gateway.example?api-version=2024-01-01'),
      ),
    ).toEqual({
      protocol: 'gemini',
      baseUrl: 'https://gateway.example/v1?api-version=2024-01-01',
      omitTemperature: true,
    })
  })

  it('strips a trailing /v1 from the path, not the query (anthropic route)', () => {
    expect(
      zen.resolveEndpoint(
        config('claude-sonnet-5', 'https://gateway.example/v1?api-version=2024-01-01'),
      ),
    ).toEqual({
      protocol: 'anthropic',
      baseUrl: 'https://gateway.example?api-version=2024-01-01',
    })
  })

  it('still composes correctly without a query string', () => {
    expect(zen.resolveEndpoint(config('glm-5.2', 'https://mirror.example/zen/v1/'))).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://mirror.example/zen/v1',
    })
    expect(
      zen.resolveEndpoint(config('claude-sonnet-5', 'https://mirror.example/zen/v1/')),
    ).toEqual({
      protocol: 'anthropic',
      baseUrl: 'https://mirror.example/zen',
    })
  })
})
