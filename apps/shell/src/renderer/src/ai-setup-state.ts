import { activeProvider } from '@genoffice/ai-provider/browser'
import type { AiSettings } from '@genoffice/ai-provider/browser'

/**
 * Whether the home "set up AI" entry should show: no usable provider is
 * configured AND the default Genspark backend is not signed in. Mirrors the
 * composer chip's usability test (`activeProvider` on each vendor's config,
 * with genspark judged by the gsk login state instead of a key).
 */
export function aiSetupNeeded(settings: AiSettings | null, gskLoggedIn: boolean): boolean {
  if (settings === null) return false
  if (gskLoggedIn) return false
  // a half-filled BYOK selection falls back to genspark in activeProvider(),
  // so the entry hides only once some vendor's config is actually usable
  return activeProvider(settings) === 'genspark' && !usableByokProvider(settings)
}

/**
 * Any non-genspark provider whose saved config passes the chat path's own
 * usability test (key + model for direct vendors, base URL + model for
 * user-hosted endpoints, any cliPath or model for Codex).
 */
function usableByokProvider(settings: AiSettings): boolean {
  const providers = settings.providers ?? {}
  for (const id of Object.keys(providers) as Array<keyof typeof providers>) {
    if (id === 'genspark') continue
    if (id === 'codex') {
      const codex = providers[id]
      if ((codex?.model ?? '').trim() || (codex?.cliPath ?? '').trim()) return true
      continue
    }
    if (activeProvider({ ...settings, provider: id }) === id) return true
  }
  // saved custom endpoints count too (the picker lists them as usable groups)
  for (const endpoint of settings.customEndpoints ?? []) {
    if (endpoint.baseUrl?.trim() && endpoint.model?.trim()) return true
  }
  return false
}
