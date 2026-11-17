import type { AiProviderId, AiProviderMeta } from './types'
import { AGY_CHAT_DEFAULT_MODEL } from './agy-chat'
import { isLocalVlmProvider } from './local-vlm-meta'

/**
 * Browser-safe description of the Antigravity CLI provider. It lives apart from agy-cli.ts, which
 * imports Node-only modules (child_process, fs), so renderer bundles that need the provider list
 * (registry.ts, providers.ts) never pull those in.
 */

export const AGY_PROVIDER_ID = 'agy' satisfies AiProviderId

/**
 * Base Antigravity provider model for media analysis, image understanding and search fallbacks.
 * Chat defaults to AGY_CHAT_DEFAULT_MODEL ('gemini-3.8-flash-high').
 */
export const AGY_DEFAULT_MODEL = 'gemini-3.8-flash-low'

export const AGY_PROVIDER_META: AiProviderMeta = {
  id: AGY_PROVIDER_ID,
  label: 'Antigravity CLI',
  // Populated at runtime from `agy models`; a short seed keeps the picker usable offline.
  models: [AGY_CHAT_DEFAULT_MODEL, AGY_DEFAULT_MODEL],
  defaultModel: AGY_CHAT_DEFAULT_MODEL,
  keyPlaceholder: '',
  needsCliPath: true,
}

/** Capability flags of this provider (consumed by registry.ts and the chat UIs). */
export const AGY_CAPABILITIES = { auth: 'agy-cli', vision: true, tools: true } as const

/** CLI-backed providers authenticate through their own login: no API key is stored or required. */
export function isCliProvider(provider: string): boolean {
  return provider === 'codex' || provider === AGY_PROVIDER_ID
}

/** Providers that do not require an API key (local CLIs and local on-device models). */
export function isKeylessProvider(provider: string): boolean {
  return isCliProvider(provider) || isLocalVlmProvider(provider)
}
