import type { AiSettings } from './types'
import { getProviderAdapter } from './registry'

/**
 * Determines whether the configured AI provider is ready to attempt a chat/agent run.
 *
 * Auth capability rules:
 * - 'agy-cli': local Antigravity CLI does not require an API key or explicit model in config
 * - 'codex-chatgpt': Codex app-server bridge uses the local Codex login session
 * - 'gsk-login': Genspark proxy requires a configured model
 * - 'api-key': direct providers require both a non-empty API key (or, for a custom endpoint,
 *   a named apiKeyEnv credential reference) and a model
 */
export function canAttemptChat(settings: AiSettings | null | undefined): boolean {
  if (!settings) return false
  const provider = settings.provider
  const config = settings.providers?.[provider]
  if (!config) return false

  let auth: ReturnType<typeof getProviderAdapter>['capabilities']['auth']
  try {
    auth = getProviderAdapter(provider).capabilities.auth
  } catch {
    return false
  }

  switch (auth) {
    case 'agy-cli':
    case 'codex-chatgpt':
    case 'local-vlm':
      return true
    case 'gsk-login':
      return Boolean(config.model)
    case 'api-key':
      // A custom endpoint may name an environment variable (profile.apiKeyEnv) instead of carrying
      // the key: the main process resolves it at request time, so the renderer only sees the name.
      // Naming one counts as configured; if the variable is unset the request reports that itself.
      return Boolean(
        (config.apiKey?.trim() || config.profile?.apiKeyEnv?.trim()) && config.model?.trim(),
      )
    default:
      return false
  }
}
