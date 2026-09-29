/**
 * Schema check for renderer-supplied AI settings, shared by 'ai:set-settings'
 * (what gets persisted), 'ai:stream' / 'ai:chat' (what the main process acts
 * on), and 'ai:codex-models' (the direct cliPath probe). SECURITY.md promises
 * payloads are schema-checked in the main process; a compromised renderer must
 * not be able to plant an arbitrary cliPath (later spawn()ed by the Codex
 * app-server) or a baseUrl that would receive the user's gsk bearer token.
 * Electron-free so unit tests can import it directly.
 *
 * Coercion follows the file's own house style (String() narrowing like the
 * web-search handler, filename sanitizing like the style-template handler).
 * Mirrors apps/slides/src/main/ai-settings-guard.ts (PR #1492); move both into
 * @genoffice/electron-utils once that package grows an AI home.
 */
import { statSync } from 'node:fs'

import { AI_PROVIDER_ADAPTERS, normalizeBaseUrl } from '@genoffice/ai-provider'
import type { AiProviderConfig, AiProviderId, AiSettings } from '@genoffice/ai-provider'

const PROVIDER_IDS = Object.keys(AI_PROVIDER_ADAPTERS) as AiProviderId[]

const MAX_API_KEY_LENGTH = 8192
const MAX_MODEL_LENGTH = 256
const MAX_CLI_PATH_LENGTH = 1024

/**
 * Executable path accepted from the renderer: plain path characters only (no
 * shell metacharacters — it is handed to spawn() by the Codex app-server), and
 * when it names a path it must already exist as a file. A bare command name
 * ("codex") is kept: it resolves via PATH at spawn time and a miss is a
 * handled ENOENT. Empty means auto-detect.
 */
export function validCliPath(raw: unknown): raw is string {
  if (typeof raw !== 'string') return false
  const value = raw.trim()
  if (value === '') return false
  if (value.length > MAX_CLI_PATH_LENGTH) return false
  if (/[^\w./:\\ -]/.test(value)) return false
  if (/[./\\]/.test(value)) {
    // a path, not a bare command: it must be an existing file
    try {
      return statSync(value).isFile()
    } catch {
      return false
    }
  }
  return true
}

/** http(s) URL with no embedded credentials, normalized like request time. */
function sanitizeBaseUrl(raw: unknown): string | undefined {
  try {
    return normalizeBaseUrl(typeof raw === 'string' ? raw : '', '')
  } catch {
    return undefined
  }
}

function sanitizeProviderConfig(input: unknown): AiProviderConfig {
  const raw = (typeof input === 'object' && input !== null ? input : {}) as Record<string, unknown>
  const config: AiProviderConfig = {
    apiKey: String(raw.apiKey ?? '').slice(0, MAX_API_KEY_LENGTH),
    model: String(raw.model ?? '').slice(0, MAX_MODEL_LENGTH),
  }
  if (raw.baseUrl !== undefined && raw.baseUrl !== null && raw.baseUrl !== '') {
    const baseUrl = sanitizeBaseUrl(raw.baseUrl)
    if (baseUrl !== undefined) config.baseUrl = baseUrl
  }
  if (typeof raw.cliPath === 'string' && validCliPath(raw.cliPath)) {
    config.cliPath = raw.cliPath.trim()
  }
  return config
}

/**
 * Returns a sanitized copy of `input`, or null when the payload is not even
 * the right shape (non-object, unknown provider). Invalid individual fields
 * are dropped rather than rejecting the whole payload, so a half-filled
 * provider config cannot brick the rest of the settings.
 */
export function sanitizeAiSettings(input: unknown): AiSettings | null {
  if (typeof input !== 'object' || input === null || Array.isArray(input)) return null
  const raw = input as Record<string, unknown>
  const provider = raw.provider
  if (typeof provider !== 'string' || !PROVIDER_IDS.includes(provider as AiProviderId)) return null
  if (typeof raw.providers !== 'object' || raw.providers === null || Array.isArray(raw.providers)) {
    return null
  }
  const providers = {} as AiSettings['providers']
  for (const [id, config] of Object.entries(raw.providers as Record<string, unknown>)) {
    if (!PROVIDER_IDS.includes(id as AiProviderId)) continue
    ;(providers as Record<string, AiProviderConfig>)[id] = sanitizeProviderConfig(config)
  }
  const settings: AiSettings = {
    provider: provider as AiProviderId,
    providers,
  }
  if (typeof raw.gskToolsEnabled === 'boolean') settings.gskToolsEnabled = raw.gskToolsEnabled
  if (
    typeof raw.maxOutputTokens === 'number' &&
    Number.isFinite(raw.maxOutputTokens) &&
    raw.maxOutputTokens > 0
  ) {
    settings.maxOutputTokens = Math.floor(raw.maxOutputTokens)
  }
  // media/search carry only enums and api keys today; keep them when they are
  // plain objects, drop anything else
  if (typeof raw.media === 'object' && raw.media !== null && !Array.isArray(raw.media)) {
    settings.media = raw.media as AiSettings['media']
  }
  if (typeof raw.search === 'object' && raw.search !== null && !Array.isArray(raw.search)) {
    settings.search = raw.search as AiSettings['search']
  }
  return settings
}
