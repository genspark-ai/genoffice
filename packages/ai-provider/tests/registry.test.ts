import { describe, expect, it } from 'vitest'
import { AI_PROVIDER_ADAPTERS, getProviderAdapter } from '../src/registry'
import { AI_PROVIDERS, GENSPARK_LLM_BASE_URLS } from '../src/providers'
import type { AiProviderConfig, AiProviderId } from '../src/types'

function config(model: string, baseUrl?: string): AiProviderConfig {
  return { apiKey: 'k', model, baseUrl }
}

describe('provider registry', () => {
  it('covers every provider in AI_PROVIDERS with matching meta', () => {
    for (const meta of AI_PROVIDERS) {
      expect(AI_PROVIDER_ADAPTERS[meta.id].meta).toBe(meta)
    }
    expect(Object.keys(AI_PROVIDER_ADAPTERS).sort()).toEqual(AI_PROVIDERS.map((m) => m.id).sort())
  })

  it('routes genspark by model id prefix onto the three proxy endpoints', () => {
    const resolve = (model: string) => AI_PROVIDER_ADAPTERS.genspark.resolveEndpoint(config(model))
    expect(resolve('claude-opus-4-7')).toEqual({
      protocol: 'anthropic',
      baseUrl: GENSPARK_LLM_BASE_URLS.anthropic,
    })
    expect(resolve('gemini-3.1-pro-preview')).toEqual({
      protocol: 'gemini',
      baseUrl: GENSPARK_LLM_BASE_URLS.gemini,
    })
    // gpt-5.x fixes sampling, so the proxy's OpenAI route also drops temperature
    expect(resolve('gpt-5.2')).toEqual({
      protocol: 'openai-compatible',
      baseUrl: GENSPARK_LLM_BASE_URLS.openai,
      omitTemperature: true,
    })
  })

  it('resolves direct providers to their official endpoints', () => {
    expect(AI_PROVIDER_ADAPTERS.anthropic.resolveEndpoint(config('claude-sonnet-5'))).toEqual({
      protocol: 'anthropic',
      baseUrl: 'https://api.anthropic.com',
    })
    expect(AI_PROVIDER_ADAPTERS.gemini.resolveEndpoint(config('gemini-2.5-flash'))).toEqual({
      protocol: 'gemini',
      baseUrl: 'https://generativelanguage.googleapis.com/v1beta',
    })
    expect(AI_PROVIDER_ADAPTERS.deepseek.resolveEndpoint(config('deepseek-chat'))).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://api.deepseek.com/v1',
    })
    expect(AI_PROVIDER_ADAPTERS.openai.resolveEndpoint(config('gpt-4.1-mini'))).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://api.openai.com/v1',
    })
  })

  it('marks the GPT-5 family as fixed-sampling (rejects any non-default temperature)', () => {
    for (const model of ['gpt-5.6', 'gpt-5.6-terra', 'gpt-5.5', 'gpt-5.4-mini']) {
      expect(AI_PROVIDER_ADAPTERS.openai.resolveEndpoint(config(model))).toEqual({
        protocol: 'openai-compatible',
        baseUrl: 'https://api.openai.com/v1',
        omitTemperature: true,
      })
    }
  })

  it('resolves the catalog additions to their OpenAI-compatible endpoints', () => {
    const cases: Array<[AiProviderId, string, string]> = [
      ['glm', 'glm-5.3', 'https://open.bigmodel.cn/api/paas/v4'],
      ['qwen', 'qwen-max', 'https://dashscope.aliyuncs.com/compatible-mode/v1'],
      ['doubao', 'doubao-seed-1-6-251015', 'https://ark.cn-beijing.volces.com/api/v3'],
      ['minimax', 'MiniMax-M2.5', 'https://api.minimax.io/v1'],
      ['xai', 'grok-4.6', 'https://api.x.ai/v1'],
      ['mistral', 'mistral-large-latest', 'https://api.mistral.ai/v1'],
      ['openrouter', 'openrouter/auto', 'https://openrouter.ai/api/v1'],
    ]
    for (const [id, model, baseUrl] of cases) {
      expect(AI_PROVIDER_ADAPTERS[id].resolveEndpoint(config(model))).toEqual({
        protocol: 'openai-compatible',
        baseUrl,
      })
    }
  })

  it('marks Kimi as fixed-sampling (K3 rejects any temperature but 1)', () => {
    expect(AI_PROVIDER_ADAPTERS.kimi.resolveEndpoint(config('kimi-k3'))).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://api.moonshot.ai/v1',
      omitTemperature: true,
    })
  })

  it('lets a stored base URL override a fixed endpoint (regional mirrors)', () => {
    expect(
      AI_PROVIDER_ADAPTERS.kimi.resolveEndpoint(config('kimi-k3', 'https://api.moonshot.cn/v1')),
    ).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://api.moonshot.cn/v1',
      omitTemperature: true,
    })
    // empty string falls back to the default
    expect(AI_PROVIDER_ADAPTERS.anthropic.resolveEndpoint(config('claude-sonnet-5', ''))).toEqual({
      protocol: 'anthropic',
      baseUrl: 'https://api.anthropic.com',
    })
  })

  it('uses the configured base URL for custom and rejects a missing one', () => {
    expect(
      AI_PROVIDER_ADAPTERS.custom.resolveEndpoint(config('m', 'http://localhost:1234/v1')),
    ).toEqual({ protocol: 'openai-compatible', baseUrl: 'http://localhost:1234/v1' })
    expect(() => AI_PROVIDER_ADAPTERS.custom.resolveEndpoint(config('m'))).toThrow(
      'A custom provider requires a Base URL',
    )
  })

  it('only genspark authenticates through the gsk login', () => {
    for (const [id, adapter] of Object.entries(AI_PROVIDER_ADAPTERS)) {
      expect(adapter.capabilities.auth).toBe(id === 'genspark' ? 'gsk-login' : 'api-key')
    }
  })

  it('throws a typed error for ids outside the registry', () => {
    expect(() => getProviderAdapter('nonsense' as AiProviderId)).toThrow(
      'Unknown provider: nonsense',
    )
  })
})

describe('fixed-sampling models on indirect routes', () => {
  it('omits temperature for kimi-k3 via OpenRouter and via a custom endpoint', () => {
    expect(AI_PROVIDER_ADAPTERS.openrouter.resolveEndpoint(config('moonshotai/kimi-k3'))).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://openrouter.ai/api/v1',
      omitTemperature: true,
    })
    expect(
      AI_PROVIDER_ADAPTERS.custom.resolveEndpoint(config('kimi-k3', 'https://api.moonshot.cn/v1')),
    ).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://api.moonshot.cn/v1',
      omitTemperature: true,
    })
    expect(AI_PROVIDER_ADAPTERS.openrouter.resolveEndpoint(config('openrouter/auto'))).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://openrouter.ai/api/v1',
    })
  })

  it('omits temperature for gpt-5 via OpenRouter and via a custom endpoint', () => {
    expect(AI_PROVIDER_ADAPTERS.openrouter.resolveEndpoint(config('openai/gpt-5.6'))).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://openrouter.ai/api/v1',
      omitTemperature: true,
    })
    expect(
      AI_PROVIDER_ADAPTERS.custom.resolveEndpoint(config('gpt-5.6-terra', 'https://mirror/v1')),
    ).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://mirror/v1',
      omitTemperature: true,
    })
  })
})
