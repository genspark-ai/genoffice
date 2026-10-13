// Resolve a custom endpoint's `apiKeyEnv` credential reference. Main-process
// only: the point of the indirection is that the settings file stores the
// *name* of an environment variable, not a secret, so the file can be synced
// or shared while the key lives in the environment the app was launched with.
// The inline apiKey always wins, so enabling this never changes the behavior
// of a settings file that was written without it.
import { normalizeCustomProfile } from './custom-profile'
import type { AiProviderConfig } from './types'

/** Environment lookup seam, so tests and alternative credential stores can substitute. */
export type EnvLookup = (name: string) => string | undefined

/**
 * A copy of `config` whose apiKey is filled from the profile's `apiKeyEnv`
 * when the inline key is empty. Anything without a resolvable reference —
 * no profile, no env var set, a non-custom provider — returns the input
 * unchanged, so callers can apply it blindly to every config.
 */
export function resolveProfileCredential(
  config: AiProviderConfig,
  lookup: EnvLookup = (name) => process.env[name],
): AiProviderConfig {
  if (config.apiKey?.trim()) return config
  const apiKeyEnv = normalizeCustomProfile(config.profile)?.apiKeyEnv
  if (!apiKeyEnv) return config
  const value = lookup(apiKeyEnv)
  if (!value?.trim()) return config
  return { ...config, apiKey: value.trim() }
}
