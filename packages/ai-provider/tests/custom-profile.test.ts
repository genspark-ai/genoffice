import { describe, expect, it } from 'vitest'
import {
  AI_CUSTOM_PROTOCOLS,
  normalizeCustomProfile,
  resolveCustomProfile,
  customProtocolBaseUrl,
} from '../src/custom-profile'
import { AI_PROVIDER_ADAPTERS, endpointSupportsVision } from '../src/registry'
import { resolveCustomEndpoints, upsertCustomEndpoint } from '../src/custom-endpoints'
import { resolveProfileCredential } from '../src/credential-env'
import type { AiCustomEndpoint, AiSettings } from '../src/types'

const endpoint = (over: Partial<AiCustomEndpoint> = {}): AiCustomEndpoint => ({
  id: 'ep-test',
  name: 'test',
  baseUrl: 'https://gw.internal/v1',
  apiKey: 'sk-test',
  model: 'glm-4.6',
  ...over,
})

const settingsOf = (ep: AiCustomEndpoint): AiSettings =>
  ({
    provider: 'custom',
    providers: {},
    customEndpoints: [ep],
    customEndpoint: ep.id,
  }) as unknown as AiSettings

describe('normalizeCustomProfile', () => {
  it('accepts each protocol the surface documents', () => {
    for (const api of AI_CUSTOM_PROTOCOLS) {
      expect(normalizeCustomProfile({ api })?.api).toBe(api)
    }
  })

  it('drops an unknown protocol and a non-object payload', () => {
    expect(normalizeCustomProfile({ api: 'ollama-native' })).toBeUndefined()
    expect(normalizeCustomProfile('nope')).toBeUndefined()
    expect(normalizeCustomProfile(undefined)).toBeUndefined()
  })

  it('bounds header counts and drops pairs HTTP cannot represent', () => {
    const many = Object.fromEntries(Array.from({ length: 40 }, (_, i) => [`X-K${i}`, 'v']))
    const out = normalizeCustomProfile({ headers: many })
    expect(Object.keys(out!.headers!)).toHaveLength(16)
    // a CRLF value is a header-injection attempt; the pair must not survive
    expect(
      normalizeCustomProfile({ headers: { 'X-Evil': 'a\r\nX-Forged: 1' } })?.headers,
    ).toBeUndefined()
    expect(normalizeCustomProfile({ headers: { 'bad name': 'v' } })?.headers).toBeUndefined()
  })

  it('keeps only real booleans in compat', () => {
    expect(
      normalizeCustomProfile({ compat: { omitTemperature: 'yes', vision: false } })?.compat,
    ).toEqual({
      vision: false,
    })
  })

  it('refuses an oversized or circular bodyExtras blob', () => {
    const huge = { blob: 'x'.repeat(70_000) }
    expect(normalizeCustomProfile({ bodyExtras: huge })?.bodyExtras).toBeUndefined()
    const circular: Record<string, unknown> = {}
    circular.self = circular
    expect(normalizeCustomProfile({ bodyExtras: circular })?.bodyExtras).toBeUndefined()
  })

  it('normalizes model profiles and drops entries without an id', () => {
    const out = normalizeCustomProfile({
      modelProfiles: [
        { id: 'glm-4.6', alias: 'glm-4.6-air', vision: true, maxTokens: 8192 },
        { alias: 'no-id' },
      ],
    })
    expect(out?.modelProfiles).toEqual([
      { id: 'glm-4.6', alias: 'glm-4.6-air', vision: true, maxTokens: 8192 },
    ])
  })
})

describe('resolveCustomProfile', () => {
  it('an absent profile keeps every pre-profile answer', () => {
    const r = resolveCustomProfile(undefined, 'deepseek-v4-pro')
    expect(r).toEqual({
      protocol: 'openai-compatible',
      omitTemperature: false,
      useMaxCompletionTokens: false,
    })
    expect(r.protocol).toBe('openai-compatible')
  })

  it('falls through model > endpoint > heuristic', () => {
    const profile = normalizeCustomProfile({
      compat: { omitTemperature: true },
      modelProfiles: [{ id: 'fast-model', omitTemperature: false }],
    })!
    // heuristic says the family fixes sampling…
    expect(resolveCustomProfile(profile, 'kimi-k3').omitTemperature).toBe(true)
    // endpoint switch says so for everything…
    expect(resolveCustomProfile(profile, 'some-new-model').omitTemperature).toBe(true)
    // and the model declaration outranks both
    expect(resolveCustomProfile(profile, 'fast-model').omitTemperature).toBe(false)
  })

  it('echoReasoning is tri-state so a declared false silences an echoing family', () => {
    const off = normalizeCustomProfile({ compat: { echoReasoning: false } })!
    expect(resolveCustomProfile(off, 'minimax-m3').echoReasoning).toBe(false)
    // undeclared: the field is absent and the protocol applies its heuristic
    expect('echoReasoning' in resolveCustomProfile(undefined, 'minimax-m3')).toBe(false)
  })

  it('a declared alias rewrites the wire id', () => {
    const profile = normalizeCustomProfile({
      modelProfiles: [{ id: 'deep-seek-v4.1-flash', alias: 'deepseek-flash' }],
    })!
    expect(resolveCustomProfile(profile, 'deep-seek-v4.1-flash').wireModel).toBe('deepseek-flash')
  })

  it('vision is only present when declared', () => {
    const r = normalizeCustomProfile({ compat: { vision: false } })!
    expect(resolveCustomProfile(r, 'glm-4.6').vision).toBe(false)
    expect('vision' in resolveCustomProfile(undefined, 'glm-4.6')).toBe(false)
  })
})

describe('custom adapter protocol dispatch', () => {
  const config = (over: Record<string, unknown> = {}) =>
    ({ apiKey: 'k', model: 'm', baseUrl: 'https://gw.example.com/v1', ...over }) as never

  it('routes an Anthropic-shaped gateway and folds a pasted /v1 once', () => {
    const e = AI_PROVIDER_ADAPTERS.custom.resolveEndpoint(config({ profile: { api: 'anthropic' } }))
    expect(e.protocol).toBe('anthropic')
    expect(e.baseUrl).toBe('https://gw.example.com')
    // a bare host keeps its shape; the protocol appends v1/messages
    const bare = AI_PROVIDER_ADAPTERS.custom.resolveEndpoint(
      config({ baseUrl: 'https://gw.example.com', profile: { api: 'anthropic' } }),
    )
    expect(bare.baseUrl).toBe('https://gw.example.com')
  })

  it('routes a Gemini-shaped gateway, adding /v1beta only to a bare host', () => {
    const e = AI_PROVIDER_ADAPTERS.custom.resolveEndpoint(config({ profile: { api: 'gemini' } }))
    expect(e.protocol).toBe('gemini')
    expect(e.baseUrl).toBe('https://gw.example.com/v1') // path already present: use as stored
    const bare = AI_PROVIDER_ADAPTERS.custom.resolveEndpoint(
      config({ baseUrl: 'https://gw.example.com', profile: { api: 'gemini' } }),
    )
    expect(bare.baseUrl).toBe('https://gw.example.com/v1beta')
  })

  it('keeps the historic exact contract when no profile is stored', () => {
    expect(AI_PROVIDER_ADAPTERS.custom.resolveEndpoint(config())).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://gw.example.com/v1',
    })
  })

  it('carries headers, idle budget and body extras onto the resolved endpoint', () => {
    const e = AI_PROVIDER_ADAPTERS.custom.resolveEndpoint(
      config({
        profile: {
          api: 'anthropic',
          headers: { 'X-Tenant': 'office' },
          idleTimeoutMs: 300000,
          bodyExtras: { chat_template_kwargs: { thinking: false } },
        },
      }),
    )
    expect(e.headers).toEqual({ 'X-Tenant': 'office' })
    expect(e.idleTimeoutMs).toBe(300000)
    expect(e.bodyExtras).toEqual({ chat_template_kwargs: { thinking: false } })
  })

  it('the request base and the probe base are the same address', () => {
    // regression: the probe once appended v1/models to the raw stored URL while
    // the request path stripped a trailing /v1 first, so the settings pane listed
    // models from a route no chat turn ever used.
    for (const [api, stored] of [
      ['anthropic', 'https://gw.example.com/v1'],
      ['anthropic', 'https://gw.example.com'],
      ['gemini', 'https://gw.example.com'],
      ['gemini', 'https://gw.example.com/v1beta'],
      ['openai-compatible', 'https://gw.example.com/v1'],
    ] as const) {
      const endpoint = AI_PROVIDER_ADAPTERS.custom.resolveEndpoint(
        config({ baseUrl: stored, profile: { api } }),
      )
      expect(endpoint.baseUrl).toBe(customProtocolBaseUrl(api, stored))
    }
  })
})

describe('endpointSupportsVision', () => {
  it('answers from the declared endpoint compat', () => {
    expect(
      endpointSupportsVision({
        provider: 'custom',
        providers: {
          custom: {
            apiKey: 'k',
            model: 'some-new-model',
            baseUrl: 'http://x/v1',
            profile: { compat: { vision: false } },
          },
        },
      }),
    ).toBe(false)
  })

  it('a model declaration outranks the endpoint switch', () => {
    const settings: Parameters<typeof endpointSupportsVision>[0] = {
      provider: 'custom',
      providers: {
        custom: {
          apiKey: 'k',
          model: 'vl-model',
          baseUrl: 'http://x/v1',
          profile: { compat: { vision: false }, modelProfiles: [{ id: 'vl-model', vision: true }] },
        },
      },
    }
    expect(endpointSupportsVision(settings)).toBe(true)
  })

  it('undeclared endpoints keep the shipped fallbacks', () => {
    expect(
      endpointSupportsVision({
        provider: 'custom',
        providers: { custom: { apiKey: 'k', model: 'deepseek-v4-pro', baseUrl: 'http://x/v1' } },
      }),
    ).toBe(false) // shipped text-only list
    expect(
      endpointSupportsVision({
        provider: 'custom',
        providers: { custom: { apiKey: 'k', model: 'mystery-model', baseUrl: 'http://x/v1' } },
      }),
    ).toBe(true) // assume vision, as before
  })

  it('a broken settings payload is text-only, not a crash', () => {
    expect(endpointSupportsVision({ provider: 'not-a-provider' as never })).toBe(false)
  })
})

describe('custom endpoint list carries profiles', () => {
  it('mirrors the active endpoint profile into the slot', () => {
    const ep = endpoint({ api: 'anthropic', headers: { 'X-Tenant': 'o' }, idleTimeoutMs: 300_000 })
    const out = resolveCustomEndpoints(settingsOf(ep))
    expect(out.providers.custom?.profile).toEqual({
      api: 'anthropic',
      headers: { 'X-Tenant': 'o' },
      idleTimeoutMs: 300000,
    })
  })

  it('a profile stored nested in the slot seeds the entry flat', () => {
    const settings = {
      provider: 'custom',
      providers: {
        custom: {
          apiKey: 'k',
          model: 'm',
          baseUrl: 'http://x/v1',
          profile: { api: 'gemini' },
        },
      },
    } as unknown as AiSettings
    const out = resolveCustomEndpoints(settings)
    expect(out.customEndpoints?.[0]?.api).toBe('gemini')
    expect(out.providers.custom?.profile).toEqual({ api: 'gemini' })
  })

  it('upsert keeps profile edits', () => {
    const base = resolveCustomEndpoints(settingsOf(endpoint()))
    const edited = upsertCustomEndpoint(
      base,
      { ...endpoint(), compat: { omitTemperature: true } },
      true,
    )
    expect(edited.providers.custom?.profile).toEqual({ compat: { omitTemperature: true } })
  })

  it('an endpoint without profile fields mirrors none (no shape drift)', () => {
    const out = resolveCustomEndpoints(settingsOf(endpoint()))
    expect(out.providers.custom).toEqual({
      apiKey: 'sk-test',
      model: 'glm-4.6',
      baseUrl: 'https://gw.internal/v1',
    })
  })
})

describe('resolveProfileCredential', () => {
  it('fills an empty key from the named environment variable', () => {
    const config = { apiKey: '', model: 'm', profile: { apiKeyEnv: 'GW_KEY' } }
    const out = resolveProfileCredential(config, (name) =>
      name === 'GW_KEY' ? ' env-key ' : undefined,
    )
    expect(out.apiKey).toBe('env-key')
  })

  it('the inline key always wins and an unset variable changes nothing', () => {
    const inline = { apiKey: 'sk-inline', model: 'm', profile: { apiKeyEnv: 'GW_KEY' } }
    expect(resolveProfileCredential(inline, () => 'env-key').apiKey).toBe('sk-inline')
    const missing = { apiKey: '', model: 'm', profile: { apiKeyEnv: 'GW_KEY' } }
    expect(resolveProfileCredential(missing, () => undefined).apiKey).toBe('')
  })

  it('a config without a profile is returned as the same object', () => {
    const plain = { apiKey: '', model: 'm' }
    expect(resolveProfileCredential(plain, () => 'x')).toBe(plain)
  })
})
