import { describe, expect, it } from 'vitest'
import {
  AI_PROVIDER_ADAPTERS,
  getProviderAdapter,
  modelEchoesReasoning,
  modelLacksVision,
} from '../src/registry'
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

  it('routes genspark by model id prefix onto the two proxy endpoints', () => {
    const resolve = (model: string) => AI_PROVIDER_ADAPTERS.genspark.resolveEndpoint(config(model))
    expect(resolve('claude-opus-4-7')).toEqual({
      protocol: 'anthropic',
      baseUrl: GENSPARK_LLM_BASE_URLS.anthropic,
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
    expect(AI_PROVIDER_ADAPTERS.gemini.resolveEndpoint(config('gemini-3.7-flash'))).toEqual({
      protocol: 'gemini',
      baseUrl: 'https://generativelanguage.googleapis.com/v1beta',
      omitTemperature: true,
    })
    // thinking stays off: once tools are in play DeepSeek 400s any turn that
    // does not echo back the reasoning_content our transcript cannot carry
    expect(AI_PROVIDER_ADAPTERS.deepseek.resolveEndpoint(config('deepseek-v4-pro'))).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://api.deepseek.com/v1',
      bodyExtras: { thinking: { type: 'disabled' } },
    })
    expect(AI_PROVIDER_ADAPTERS.openai.resolveEndpoint(config('gpt-4.1-mini'))).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://api.openai.com/v1',
      useMaxCompletionTokens: true,
    })
  })

  it('marks the GPT-5 family as fixed-sampling (rejects any non-default temperature)', () => {
    for (const model of ['gpt-5.6', 'gpt-5.6-terra', 'gpt-5.5', 'gpt-5.4-mini']) {
      expect(AI_PROVIDER_ADAPTERS.openai.resolveEndpoint(config(model))).toEqual({
        protocol: 'openai-compatible',
        baseUrl: 'https://api.openai.com/v1',
        omitTemperature: true,
        useMaxCompletionTokens: true,
      })
    }
  })

  it('marks the OpenAI o-series reasoning models as fixed-sampling', () => {
    for (const model of ['o1', 'o1-mini', 'o1-preview', 'o3', 'o3-mini', 'o4-mini']) {
      expect(AI_PROVIDER_ADAPTERS.openai.resolveEndpoint(config(model))).toEqual({
        protocol: 'openai-compatible',
        baseUrl: 'https://api.openai.com/v1',
        omitTemperature: true,
        useMaxCompletionTokens: true,
      })
    }
    // Non-reasoning models still carry the configured temperature.
    expect(AI_PROVIDER_ADAPTERS.openai.resolveEndpoint(config('gpt-4o'))).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://api.openai.com/v1',
      useMaxCompletionTokens: true,
    })
  })

  it('resolves the catalog additions to their OpenAI-compatible endpoints', () => {
    const cases: Array<[AiProviderId, string, string]> = [
      ['glm', 'glm-5.3', 'https://open.bigmodel.cn/api/paas/v4'],
      ['qwen', 'qwen3.8-max', 'https://dashscope.aliyuncs.com/compatible-mode/v1'],
      ['doubao', 'doubao-seed-2-1-pro-260628', 'https://ark.cn-beijing.volces.com/api/v3'],
      ['minimax', 'MiniMax-M3', 'https://api.minimax.io/v1'],
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

  it('routes OpenCode Zen per model onto the protocol the gateway serves it on', () => {
    const resolve = (model: string, baseUrl?: string) =>
      AI_PROVIDER_ADAPTERS['opencode-zen'].resolveEndpoint(config(model, baseUrl))
    expect(resolve('claude-sonnet-5')).toEqual({
      protocol: 'anthropic',
      baseUrl: 'https://opencode.ai/zen',
    })
    expect(resolve('qwen3.6-plus')).toEqual({
      protocol: 'anthropic',
      baseUrl: 'https://opencode.ai/zen',
    })
    // the Gemini 3 family keeps Google's default sampling on this route too
    expect(resolve('gemini-3.7-flash')).toEqual({
      protocol: 'gemini',
      baseUrl: 'https://opencode.ai/zen/v1',
      omitTemperature: true,
    })
    expect(resolve('gemini-2.5-flash')).toEqual({
      protocol: 'gemini',
      baseUrl: 'https://opencode.ai/zen/v1',
    })
    for (const model of ['deepseek-v4-pro', 'glm-5.2', 'minimax-m3']) {
      expect(resolve(model)).toEqual({
        protocol: 'openai-compatible',
        baseUrl: 'https://opencode.ai/zen/v1',
      })
    }
    // Kimi fixes sampling on the K2 line as well as K3
    for (const model of ['kimi-k3', 'kimi-k2.7-code']) {
      expect(resolve(model)).toEqual({
        protocol: 'openai-compatible',
        baseUrl: 'https://opencode.ai/zen/v1',
        omitTemperature: true,
      })
    }
    // a stored base URL replaces the gateway root, with or without the documented /v1
    expect(resolve('claude-sonnet-5', 'https://mirror.example/zen/v1/')).toEqual({
      protocol: 'anthropic',
      baseUrl: 'https://mirror.example/zen',
    })
    expect(resolve('glm-5.2', 'https://mirror.example/zen')).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://mirror.example/zen/v1',
    })
  })

  it('routes OpenCode Go with its own table (MiniMax rides Messages there, not chat-completions)', () => {
    const resolve = (model: string) =>
      AI_PROVIDER_ADAPTERS['opencode-go'].resolveEndpoint(config(model))
    for (const model of ['minimax-m3', 'qwen3.8-flash']) {
      expect(resolve(model)).toEqual({
        protocol: 'anthropic',
        baseUrl: 'https://opencode.ai/zen/go',
      })
    }
    for (const model of ['glm-5.3', 'deepseek-v4-flash', 'qwen3.8-max', 'longcat-2.0']) {
      expect(resolve(model)).toEqual({
        protocol: 'openai-compatible',
        baseUrl: 'https://opencode.ai/zen/go/v1',
      })
    }
    expect(resolve('kimi-k2.7-code')).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://opencode.ai/zen/go/v1',
      omitTemperature: true,
    })
    // Go has no Gemini route
    expect(resolve('gemini-3.7-flash').protocol).toBe('openai-compatible')
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
      expect(adapter.capabilities.auth).toBe(
        id === 'genspark' ? 'gsk-login' : id === 'codex' ? 'codex-chatgpt' : 'api-key',
      )
    }
  })

  it('routes Codex to the auto-discovered local process bridge', () => {
    expect(
      AI_PROVIDER_ADAPTERS.codex.resolveEndpoint({
        apiKey: '',
        model: 'gpt-5.6-terra',
        cliPath: 'C:\\Tools\\codex.exe',
      }),
    ).toEqual({ protocol: 'codex-app-server', baseUrl: '' })
    expect(AI_PROVIDER_ADAPTERS.codex.resolveEndpoint(config('gpt-5.6-terra'))).toEqual({
      protocol: 'codex-app-server',
      baseUrl: '',
    })
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
    expect(AI_PROVIDER_ADAPTERS.openrouter.resolveEndpoint(config('openai/gpt-5.6-sol'))).toEqual({
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

describe('modelLacksVision', () => {
  it('flags text-only DeepSeek V4 models but not the vision branch', () => {
    expect(modelLacksVision('deep-seek-v4-flash')).toBe(true)
    expect(modelLacksVision('deep-seek-v4-flash-baseten')).toBe(true)
    expect(modelLacksVision('deepseek-v4-pro')).toBe(true)
    expect(modelLacksVision('deepseek-v4-flash')).toBe(true)
    expect(modelLacksVision('deepseek-v4-flash-vision-exp')).toBe(false)
    expect(modelLacksVision('deep-seek-v4-flash-vision-exp-openrouter')).toBe(false)
    expect(modelLacksVision('claude-opus-4-7')).toBe(false)
  })
})

describe('modelEchoesReasoning', () => {
  it('flags interleaved-thinking families on any route, case-insensitively', () => {
    expect(modelEchoesReasoning('MiniMax-M3')).toBe(true)
    expect(modelEchoesReasoning('minimax-m2p7')).toBe(true)
    expect(modelEchoesReasoning('deep-seek-v4-flash')).toBe(true)
    expect(modelEchoesReasoning('deepseek-v4-pro')).toBe(true)
    expect(modelEchoesReasoning('gpt-5.6-luna')).toBe(false)
    expect(modelEchoesReasoning('kimi-k3')).toBe(false)
  })
})
