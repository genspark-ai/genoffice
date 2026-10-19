import { describe, expect, it } from 'vitest'
import {
  defaultAiSettings,
  resolveCustomEndpoints,
  resolveProfileCredential,
  type AiCustomProfile,
} from '@genoffice/ai-provider'
import {
  aiChatRequestSchema,
  aiSettingsInputSchema,
  aiStreamRequestSchema,
} from '../src/shared/desktop-api'

const profile: AiCustomProfile = {
  api: 'anthropic',
  apiKeyEnv: 'TEST_GATEWAY_KEY',
  headers: { 'X-Tenant': 'office' },
  idleTimeoutMs: 300_000,
  compat: {
    omitTemperature: true,
    echoReasoning: false,
    vision: false,
    useMaxCompletionTokens: true,
  },
  bodyExtras: { thinking: { enabled: false } },
  modelProfiles: [{ id: 'm', alias: 'wire-m', name: 'Model M', maxTokens: 8192 }],
  probedProfiles: [{ id: 'm', vision: true, contextWindow: 128_000 }],
}

function settingsWithProfile() {
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = {
    apiKey: '',
    model: 'm',
    baseUrl: 'https://gateway.example/v1',
    profile,
  }
  return resolveCustomEndpoints(settings)
}

describe('Sheets AI settings IPC profiles', () => {
  it.each(['chat', 'stream'] as const)(
    'preserves an env-only credential and the endpoint profile on the %s path',
    (path) => {
      const settings = settingsWithProfile()
      const request =
        path === 'chat'
          ? aiChatRequestSchema.parse({ settings, system: 'sys', user: 'ping' })
          : aiStreamRequestSchema.parse({
              requestId: 'test-request',
              settings,
              system: 'sys',
              messages: [],
            })
      const config = request.settings.providers.custom!
      expect(config.profile).toEqual(profile)
      expect(request.settings.customEndpoints?.[0]).toMatchObject(profile)
      expect(resolveProfileCredential(config, () => 'test-key').apiKey).toBe('test-key')
      expect(config.apiKey).toBe('')
    },
  )

  it('preserves flat declarations and the separate probe cache through a settings save', () => {
    const settings = settingsWithProfile()
    const parsed = aiSettingsInputSchema.parse(settings)
    expect(parsed).toEqual(settings)
    expect(parsed.customEndpoints?.[0]?.modelProfiles).toEqual(profile.modelProfiles)
    expect(parsed.customEndpoints?.[0]?.probedProfiles).toEqual(profile.probedProfiles)
  })

  it('keeps legacy settings without a profile unchanged', () => {
    const settings = defaultAiSettings()
    expect(aiSettingsInputSchema.parse(settings)).toEqual(settings)
  })

  it('normalizes unsafe headers and oversized fields in both profile representations', () => {
    const settings = settingsWithProfile()
    const oversized: AiCustomProfile = {
      ...profile,
      apiKeyEnv: 'X'.repeat(200),
      headers: { 'X-Tenant': 'office', 'X-Injected': 'a\r\nX-Forged: yes' },
      bodyExtras: { blob: 'x'.repeat(70_000) },
      probedProfiles: Array.from({ length: 250 }, (_, i) => ({ id: `m${i}` })),
    }
    settings.providers.custom.profile = oversized
    settings.customEndpoints![0] = { ...settings.customEndpoints![0]!, ...oversized }
    const parsed = aiSettingsInputSchema.parse(settings)
    for (const normalized of [parsed.providers.custom!.profile, parsed.customEndpoints?.[0]]) {
      expect(normalized?.apiKeyEnv).toHaveLength(128)
      expect(normalized?.headers).toEqual({ 'X-Tenant': 'office' })
      expect(normalized?.bodyExtras).toBeUndefined()
      expect(normalized?.probedProfiles).toHaveLength(200)
    }
  })

  it.each([
    { api: 'unsupported-protocol' },
    { compat: { vision: 'yes' } },
    { modelProfiles: [{ id: 'm', maxTokens: -1 }] },
    { probedProfiles: [{ id: 'm', unexpected: true }] },
    { unexpected: true },
  ])('rejects invalid or unknown profile fields: %j', (invalid) => {
    const settings = settingsWithProfile()
    const nested = {
      ...settings,
      providers: {
        ...settings.providers,
        custom: { ...settings.providers.custom, profile: invalid },
      },
    }
    expect(aiSettingsInputSchema.safeParse(nested).success).toBe(false)
    const flat = {
      ...settings,
      customEndpoints: [{ ...settings.customEndpoints![0]!, ...invalid }],
    }
    expect(aiSettingsInputSchema.safeParse(flat).success).toBe(false)
  })
})
