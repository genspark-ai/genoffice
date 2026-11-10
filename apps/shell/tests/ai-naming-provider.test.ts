import { describe, expect, it } from 'vitest'
import { resolveNamingProvider } from '../src/main/ai-naming-provider'

/**
 * The naming service used to resolve its provider on its own: a hardcoded list
 * of three ids, and the key read out of the settings file. That list did not
 * know every provider the build can talk to, and it never saw the Genspark
 * key — which lives in the login state, not in the file — so for the default
 * account the call failed and the feature quietly did nothing.
 *
 * These assertions are the evidence that the provider is now resolved the way
 * the rest of the app resolves it.
 */

describe('resolveNamingProvider', () => {
  it('takes the Genspark key from the login state, not from the settings file', () => {
    const resolved = resolveNamingProvider({}, 'gsk-live-key')
    expect(resolved?.provider).toBe('genspark')
    expect(resolved?.config.apiKey).toBe('gsk-live-key')
  })

  it('declines when Genspark is selected and nobody is signed in', () => {
    expect(resolveNamingProvider({}, '')).toBeNull()
  })

  it('declines when the settings file could not be read at all', () => {
    expect(resolveNamingProvider(null, '')).toBeNull()
  })

  it('honors a configured provider', () => {
    const resolved = resolveNamingProvider(
      {
        provider: 'anthropic',
        providers: { anthropic: { apiKey: 'sk-ant', model: 'claude-sonnet-4-6' } },
      },
      '',
    )
    expect(resolved?.provider).toBe('anthropic')
    expect(resolved?.config.apiKey).toBe('sk-ant')
    expect(resolved?.config.model).toBe('claude-sonnet-4-6')
  })

  it('falls back to Genspark when the configured provider is half-filled', () => {
    const resolved = resolveNamingProvider(
      {
        provider: 'anthropic',
        providers: { anthropic: { apiKey: 'sk-ant', model: '' } },
      },
      'gsk-live-key',
    )
    expect(resolved?.provider).toBe('genspark')
    expect(resolved?.config.apiKey).toBe('gsk-live-key')
  })

  it('falls back to Genspark for a provider this build does not have', () => {
    const resolved = resolveNamingProvider(
      { provider: 'not-a-provider', providers: {} },
      'gsk-live-key',
    )
    expect(resolved?.provider).toBe('genspark')
  })

  it('does not need a key for codex', () => {
    const resolved = resolveNamingProvider({ provider: 'codex', providers: {} }, '')
    expect(resolved?.provider).toBe('codex')
  })

  it("resolves a custom endpoint's key from the environment", () => {
    process.env.GENOFFICE_NAMING_TEST_KEY = 'from-env'
    try {
      const resolved = resolveNamingProvider(
        {
          provider: 'custom',
          providers: {
            custom: {
              apiKey: '',
              model: 'local-model',
              baseUrl: 'http://127.0.0.1:11434/v1',
              profile: { apiKeyEnv: 'GENOFFICE_NAMING_TEST_KEY' },
            },
          },
        },
        '',
      )
      expect(resolved?.provider).toBe('custom')
      expect(resolved?.config.apiKey).toBe('from-env')
    } finally {
      delete process.env.GENOFFICE_NAMING_TEST_KEY
    }
  })
})
