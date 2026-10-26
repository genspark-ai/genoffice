/**
 * Which provider answers a file-naming request.
 *
 * The naming service used to carry its own list of provider ids and read the
 * key out of the settings file. That was the wrong answer twice over: the list
 * did not know every provider the build can talk to, and it never saw the
 * Genspark key, which lives in the login state rather than in the file — so
 * for the default account the call failed and the feature quietly did nothing.
 *
 * This module resolves the provider the way the rest of the app does: the
 * shared `resolveAiSettings` + `activeProvider` from `@genoffice/ai-provider`,
 * the Genspark key from the gsk login state, and a custom endpoint's
 * credential reference resolved from the environment.
 */
import type { AiProviderConfig, AiProviderId } from '@genoffice/ai-provider'
import {
  activeProvider,
  defaultAiSettings,
  resolveAiSettings,
  resolveProfileCredential,
} from '@genoffice/ai-provider'

/** The provider to call, and the credentials to call it with. */
export interface NamingProvider {
  provider: AiProviderId
  config: AiProviderConfig
}

/**
 * The provider a naming request goes to, or null when no provider can answer.
 *
 * `stored` is the parsed `ai-settings.json`; `gskKey` is the Genspark key from
 * the login state, empty when signed out. Null rather than a broken provider:
 * a request with no usable provider is declined, and the caller keeps the name
 * the document already has.
 */
export function resolveNamingProvider(stored: unknown, gskKey: string): NamingProvider | null {
  const settings = resolveAiSettings(stored, defaultAiSettings())
  // a stored BYOK provider is honored when usable; a half-filled one falls back
  // to genspark rather than silently disabling AI
  settings.provider = activeProvider(settings)
  const provider = settings.provider
  const storedConfig = settings.providers?.[provider]
  if (!storedConfig) return null
  let config = storedConfig
  // the genspark key never enters the settings file; requests take it from the
  // gsk login state
  if (provider === 'genspark' && !config.apiKey?.trim()) {
    config = { ...config, apiKey: gskKey }
  }
  // a custom endpoint may name an environment variable instead of storing a key
  config = resolveProfileCredential(config)
  // codex drives its own executable and needs neither a key nor a model id
  if (provider !== 'codex' && !config.apiKey?.trim()) return null
  if (provider !== 'codex' && !config.model?.trim()) return null
  return { provider, config }
}
