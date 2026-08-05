/** The tasks that can select different AI providers. */
export type AiTask = 'chat' | 'image' | 'vision' | 'slides-generation'

export const AI_SETTINGS_SCHEMA_VERSION = 1 as const

export interface AiSelection {
  providerId: string
  model: string
}

export interface AiTaskSelections {
  chat: AiSelection
  image: AiSelection
  vision?: AiSelection
  'slides-generation'?: AiSelection
}

/**
 * The only settings shape that may cross the renderer/main public boundary.
 * It intentionally contains a boolean instead of a credential value.
 */
export interface PublicProviderSettings {
  providerId: string
  model: string
  baseUrl?: string
  credentialConfigured: boolean
  /** Masked renderer-safe suffix, never the complete credential. */
  credentialHint?: string
  enabled: boolean
}

export interface PublicAiSettings {
  version: typeof AI_SETTINGS_SCHEMA_VERSION
  active: AiTaskSelections
  providers: Record<string, PublicProviderSettings>
}

/** Data written to the settings JSON file. It has credential ids, never keys. */
export interface PersistedProviderSettings {
  providerId: string
  model: string
  baseUrl?: string
  credentialId?: string
  enabled: boolean
}

export interface PersistedAiSettings {
  version: typeof AI_SETTINGS_SCHEMA_VERSION
  active: AiTaskSelections
  providers: Record<string, PersistedProviderSettings>
  updatedAt: string
}

/** Ephemeral input accepted by the main process; it is never returned. */
export interface UpdateProviderInput {
  providerId: string
  model?: string
  baseUrl?: string
  enabled?: boolean
  /** undefined keeps the existing credential; null removes it. */
  credential?: string | null
}

export interface SetActiveProviderInput {
  task: AiTask
  providerId: string
  model?: string
}

/** Internal request configuration, available only in the main process. */
export interface ResolvedAiProvider {
  task: AiTask
  providerId: string
  model: string
  baseUrl?: string
  apiKey?: string
}

/** A deliberately small legacy shape used only for one-time migration. */
export interface LegacyAiSettings {
  provider?: unknown
  model?: unknown
  baseUrl?: unknown
  apiKey?: unknown
  providers?: unknown
  /** Older shell builds stored task defaults outside the provider map. */
  imageProvider?: unknown
  imageModel?: unknown
}
