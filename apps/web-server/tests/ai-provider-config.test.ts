/**
 * Provider config resolution regression test.
 *
 * Genspark signs in through the shared gsk login (`~/.genoffice/auth.json` /
 * the gsk CLI config), not through a key pasted into Settings, so its stored
 * config legitimately has `apiKey: ''`. Before `resolveProviderConfig` the
 * empty key was passed straight through, the request hit genspark.ai's website
 * instead of the LLM endpoint, and every translate/dictionary call failed with
 * an HTML 403 that was surfaced to the user as "the service returned a web
 * page".
 *
 * The agent side (@genoffice/agent-skills translate-skill) resolves the same
 * file, which is why this must stay a single shared resolver rather than an
 * ad-hoc patch at one call site.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const TMP_DATA = mkdtempSync(join(tmpdir(), 'genoffice-provider-config-'))
process.env.DATA_DIR = TMP_DATA
process.env.GENOFFICE_DATA_DIR = TMP_DATA
process.env.GSK_API_KEY = 'gsk-from-env-test-key'

const { resolveProviderConfig } = await import('../src/ai/chat')
const { defaultAiSettings } = await import('@genoffice/ai-provider')

afterAll(() => {
  rmSync(TMP_DATA, { recursive: true, force: true })
  delete process.env.GSK_API_KEY
})

describe('resolveProviderConfig', () => {
  it('injects the shared gsk login key when genspark has no stored key', () => {
    const settings = defaultAiSettings()
    settings.provider = 'genspark'
    settings.providers.genspark = { ...settings.providers.genspark, apiKey: '' }
    const resolved = resolveProviderConfig(settings, 'genspark')
    expect(resolved?.apiKey).toBe('gsk-from-env-test-key')
  })

  it('never overrides a key the user explicitly stored', () => {
    const settings = defaultAiSettings()
    settings.providers.genspark = { ...settings.providers.genspark, apiKey: 'user-key' }
    expect(resolveProviderConfig(settings, 'genspark')?.apiKey).toBe('user-key')
  })

  it('does not rewrite other providers', () => {
    const settings = defaultAiSettings()
    settings.providers.minimax = { ...settings.providers.minimax, apiKey: '' }
    expect(resolveProviderConfig(settings, 'minimax')?.apiKey).toBe('')
  })

  it('returns undefined for a provider with no config at all', () => {
    const settings = defaultAiSettings()
    delete (settings.providers as Record<string, unknown>).genspark
    expect(resolveProviderConfig(settings, 'genspark')).toBeUndefined()
  })
})
