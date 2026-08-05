import type { CredentialStore } from './credentials.js'
import { readJsonFile, writeAtomicJson } from './storage.js'
import {
  AI_SETTINGS_SCHEMA_VERSION,
  type AiTask,
  type AiTaskSelections,
  type LegacyAiSettings,
  type PersistedAiSettings,
  type PersistedProviderSettings,
  type PublicAiSettings,
  type PublicProviderSettings,
  type SetActiveProviderInput,
  type UpdateProviderInput,
} from './types.js'

export interface AiSettingsStoreOptions {
  settingsPath: string
  credentialStore: CredentialStore
  defaults: PersistedAiSettings
}

const TASKS: AiTask[] = ['chat', 'image', 'vision', 'slides-generation']

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function cloneDefaults(defaults: PersistedAiSettings): PersistedAiSettings {
  return JSON.parse(JSON.stringify(defaults)) as PersistedAiSettings
}

function credentialId(providerId: string): string {
  return `ai-provider:${providerId}`
}

/** Provider aliases that were retired upstream but may still exist in user data. */
const LEGACY_MODEL_ALIASES: Record<string, Record<string, string>> = {
  deepseek: {
    'deepseek-chat': 'deepseek-v4-flash',
    'deepseek-reasoner': 'deepseek-v4-flash',
  },
}

function canonicalModel(providerId: string, model: string): string {
  return LEGACY_MODEL_ALIASES[providerId]?.[model] ?? model
}

function canonicalTaskModel(task: AiTask, providerId: string, model: string): string {
  if (task === 'image' && providerId === 'genspark' && model === 'claude-opus-4-7') {
    return 'nano-banana-2'
  }
  return canonicalModel(providerId, model)
}

function normalizeProvider(
  providerId: string,
  input: unknown,
  fallback?: PersistedProviderSettings,
): PersistedProviderSettings {
  const value = isRecord(input) ? input : {}
  const model = canonicalModel(
    providerId,
    typeof value.model === 'string' && value.model ? value.model : (fallback?.model ?? ''),
  )
  const baseUrl =
    typeof value.baseUrl === 'string' && value.baseUrl ? value.baseUrl : fallback?.baseUrl
  const enabled = typeof value.enabled === 'boolean' ? value.enabled : (fallback?.enabled ?? true)
  const id = typeof value.credentialId === 'string' ? value.credentialId : fallback?.credentialId
  return {
    providerId,
    model,
    ...(baseUrl ? { baseUrl } : {}),
    ...(id ? { credentialId: id } : {}),
    enabled,
  }
}

function providerApiKey(input: unknown): string | undefined {
  if (!isRecord(input)) return undefined
  return typeof input.apiKey === 'string' && input.apiKey ? input.apiKey : undefined
}

/** A schema-versioned file may still contain keys written by an older build. */
function containsLegacyCredentials(raw: unknown): boolean {
  if (!isRecord(raw)) return false
  if (typeof raw.apiKey === 'string' && raw.apiKey.length > 0) return true
  if (!isRecord(raw.providers)) return false
  return Object.values(raw.providers).some((value) => providerApiKey(value) !== undefined)
}

function containsLegacyModelAliases(raw: unknown): boolean {
  if (!isRecord(raw)) return false
  const providers = isRecord(raw.providers) ? Object.entries(raw.providers) : []
  if (
    providers.some(([providerId, value]) => {
      const model = isRecord(value) && typeof value.model === 'string' ? value.model : ''
      return canonicalModel(providerId, model) !== model
    })
  )
    return true
  const active = isRecord(raw.active) ? Object.entries(raw.active) : []
  return active.some(([task, value]) => {
    if (!isRecord(value) || typeof value.providerId !== 'string' || typeof value.model !== 'string')
      return false
    return canonicalTaskModel(task as AiTask, value.providerId, value.model) !== value.model
  })
}

function normalizePersisted(
  raw: unknown,
  defaults: PersistedAiSettings,
  credentials: CredentialStore,
): PersistedAiSettings {
  if (isRecord(raw) && raw.version === AI_SETTINGS_SCHEMA_VERSION && isRecord(raw.providers)) {
    const providers: Record<string, PersistedProviderSettings> = {}
    for (const [providerId, value] of Object.entries(raw.providers)) {
      const normalized = normalizeProvider(providerId, value, defaults.providers[providerId])
      const oldKey = providerApiKey(value)
      if (oldKey) {
        const id = normalized.credentialId ?? credentialId(providerId)
        credentials.set(id, oldKey)
        normalized.credentialId = id
      }
      providers[providerId] = normalized
    }
    for (const [providerId, fallback] of Object.entries(defaults.providers)) {
      if (!providers[providerId]) providers[providerId] = { ...fallback }
    }
    const active: AiTaskSelections = { ...defaults.active }
    if (isRecord(raw.active)) {
      for (const task of TASKS) {
        const selection = raw.active[task]
        if (
          isRecord(selection) &&
          typeof selection.providerId === 'string' &&
          typeof selection.model === 'string'
        ) {
          active[task] = {
            providerId: selection.providerId,
            model: canonicalTaskModel(task, selection.providerId, selection.model),
          }
        }
      }
    }
    const topLevelKey = providerApiKey(raw)
    if (topLevelKey) {
      const providerId = active.chat.providerId
      const provider = providers[providerId]
      if (provider) {
        const id = provider.credentialId ?? credentialId(providerId)
        credentials.set(id, topLevelKey)
        provider.credentialId = id
      }
    }
    return { version: 1, active, providers, updatedAt: new Date().toISOString() }
  }

  // Migrate the old ai-settings.json shape once. Keys are moved into the
  // credential store before the sanitized version is ever written back.
  const legacy = (isRecord(raw) ? raw : {}) as LegacyAiSettings
  const providers: Record<string, PersistedProviderSettings> = {}
  const legacyProviders = isRecord(legacy.providers) ? legacy.providers : {}
  const providerIds = new Set([...Object.keys(defaults.providers), ...Object.keys(legacyProviders)])
  for (const providerId of providerIds) {
    const fallback = defaults.providers[providerId]
    const old = isRecord(legacyProviders[providerId]) ? legacyProviders[providerId] : {}
    const normalized = normalizeProvider(providerId, old, fallback)
    const oldKey = providerApiKey(old)
    if (oldKey) {
      const id = normalized.credentialId ?? credentialId(providerId)
      credentials.set(id, oldKey)
      normalized.credentialId = id
    }
    providers[providerId] = normalized
  }
  const activeProvider =
    typeof legacy.provider === 'string' ? legacy.provider : defaults.active.chat.providerId
  const activeModel =
    typeof legacy.model === 'string' && legacy.model
      ? legacy.model
      : (providers[activeProvider]?.model ?? defaults.active.chat.model)
  if (!providers[activeProvider]) {
    providers[activeProvider] = {
      providerId: activeProvider,
      model: activeModel,
      ...(typeof legacy.baseUrl === 'string' && legacy.baseUrl ? { baseUrl: legacy.baseUrl } : {}),
      enabled: true,
    }
  }
  if (typeof legacy.apiKey === 'string' && legacy.apiKey) {
    const id = credentialId(activeProvider)
    credentials.set(id, legacy.apiKey)
    providers[activeProvider].credentialId = id
  }
  const imageProvider =
    typeof legacy.imageProvider === 'string' && legacy.imageProvider
      ? legacy.imageProvider
      : defaults.active.image.providerId
  const imageModel = canonicalTaskModel(
    'image',
    imageProvider,
    typeof legacy.imageModel === 'string' && legacy.imageModel
      ? legacy.imageModel
      : (providers[imageProvider]?.model ?? defaults.active.image.model),
  )
  if (!providers[imageProvider]) {
    providers[imageProvider] = {
      providerId: imageProvider,
      model: imageModel,
      enabled: true,
    }
  }
  return {
    version: 1,
    active: {
      chat: { providerId: activeProvider, model: activeModel },
      image: { providerId: imageProvider, model: imageModel },
    },
    providers,
    updatedAt: new Date().toISOString(),
  }
}

export function toPublicAiSettings(
  settings: PersistedAiSettings,
  credentials: CredentialStore,
): PublicAiSettings {
  const providers: Record<string, PublicProviderSettings> = {}
  for (const [providerId, value] of Object.entries(settings.providers)) {
    const credential = value.credentialId ? credentials.get(value.credentialId) : undefined
    providers[providerId] = {
      providerId,
      model: value.model,
      ...(value.baseUrl ? { baseUrl: value.baseUrl } : {}),
      credentialConfigured: Boolean(credential),
      ...(credential ? { credentialHint: `••••${credential.slice(-4)}` } : {}),
      enabled: value.enabled,
    }
  }
  return { version: 1, active: { ...settings.active }, providers }
}

/** Main-process owner of persisted settings, public DTOs, and the safe cache. */
export class VersionedAiSettingsStore {
  private cached: PersistedAiSettings | undefined

  constructor(private readonly options: AiSettingsStoreOptions) {}

  load(): PublicAiSettings {
    return toPublicAiSettings(this.loadInternal(), this.options.credentialStore)
  }

  updateProvider(input: UpdateProviderInput): PublicAiSettings {
    const settings = this.loadInternal()
    const previous = settings.providers[input.providerId]
    const provider: PersistedProviderSettings = {
      providerId: input.providerId,
      model: canonicalModel(input.providerId, input.model ?? previous?.model ?? ''),
      ...(input.baseUrl !== undefined
        ? input.baseUrl
          ? { baseUrl: input.baseUrl }
          : {}
        : previous?.baseUrl
          ? { baseUrl: previous.baseUrl }
          : {}),
      ...(input.enabled !== undefined
        ? { enabled: input.enabled }
        : { enabled: previous?.enabled ?? true }),
      ...(previous?.credentialId ? { credentialId: previous.credentialId } : {}),
    }
    if (input.credential !== undefined) {
      if (input.credential === null || input.credential === '') {
        if (provider.credentialId) this.options.credentialStore.delete(provider.credentialId)
        delete provider.credentialId
      } else {
        const id = provider.credentialId ?? credentialId(input.providerId)
        this.options.credentialStore.set(id, input.credential)
        provider.credentialId = id
      }
    }
    settings.providers[input.providerId] = provider
    this.persist(settings)
    return toPublicAiSettings(settings, this.options.credentialStore)
  }

  setActive(input: SetActiveProviderInput): PublicAiSettings {
    const settings = this.loadInternal()
    const provider = settings.providers[input.providerId]
    if (!provider) throw new Error(`Unknown AI provider: ${input.providerId}`)
    settings.active[input.task] = {
      providerId: input.providerId,
      model: canonicalTaskModel(input.task, input.providerId, input.model ?? provider.model),
    }
    this.persist(settings)
    return toPublicAiSettings(settings, this.options.credentialStore)
  }

  /** Resolve a request using only main-process state; renderer values are not accepted. */
  resolve(task: AiTask): import('./types.js').ResolvedAiProvider {
    const settings = this.loadInternal()
    const selection = settings.active[task] ?? settings.active.chat
    const provider = settings.providers[selection.providerId]
    if (!provider || !provider.enabled)
      throw new Error(`AI provider is unavailable: ${selection.providerId}`)
    const apiKey = provider.credentialId
      ? this.options.credentialStore.get(provider.credentialId)
      : undefined
    return {
      task,
      providerId: provider.providerId,
      model: selection.model || provider.model,
      ...(provider.baseUrl ? { baseUrl: provider.baseUrl } : {}),
      ...(apiKey ? { apiKey } : {}),
    }
  }

  /** Resolve one configured provider for connection checks without changing the active task. */
  resolveProvider(
    providerId: string,
    task: AiTask = 'chat',
  ): import('./types.js').ResolvedAiProvider {
    const settings = this.loadInternal()
    const provider = settings.providers[providerId]
    if (!provider || !provider.enabled) throw new Error(`AI provider is unavailable: ${providerId}`)
    const apiKey = provider.credentialId
      ? this.options.credentialStore.get(provider.credentialId)
      : undefined
    return {
      task,
      providerId: provider.providerId,
      model: provider.model,
      ...(provider.baseUrl ? { baseUrl: provider.baseUrl } : {}),
      ...(apiKey ? { apiKey } : {}),
    }
  }

  private loadInternal(): PersistedAiSettings {
    if (this.cached) return this.cached
    const raw = readJsonFile<unknown>(this.options.settingsPath)
    const normalized = normalizePersisted(
      raw,
      cloneDefaults(this.options.defaults),
      this.options.credentialStore,
    )
    this.cached = normalized
    // Always rewrite legacy/corrupt/old documents in the sanitized schema.
    if (
      !raw ||
      !isRecord(raw) ||
      raw.version !== AI_SETTINGS_SCHEMA_VERSION ||
      containsLegacyCredentials(raw) ||
      containsLegacyModelAliases(raw)
    )
      this.persist(normalized)
    return normalized
  }

  private persist(settings: PersistedAiSettings): void {
    const next = {
      ...settings,
      version: AI_SETTINGS_SCHEMA_VERSION,
      updatedAt: new Date().toISOString(),
    }
    writeAtomicJson(this.options.settingsPath, next)
    this.cached = next
  }
}
