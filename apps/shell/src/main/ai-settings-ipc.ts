import { ipcMain, safeStorage } from 'electron'
import { dirname, join } from 'node:path'
import {
  EncryptedCredentialStore,
  VersionedAiSettingsStore,
  configureAiSettingsStore,
  type PersistedAiSettings,
  type SafeStorageAdapter,
} from '@genoffice/ai-electron'
import {
  discoverFalImageModels,
  discoverModels,
  discoverOpenRouterImageModels,
  discoverRunwareImageModels,
  validateAiEndpoint,
  type AiProviderId,
} from '@genoffice/ai-provider'
import {
  AI_PROVIDER_DEFINITIONS,
  AI_SETTINGS_CHANNELS,
  type AiProviderConfigView,
  type AiProviderConnectionResult,
  type AiProviderDefinition,
  type AiSettingsSnapshot,
  type SaveAiProviderInput,
  type SaveAiSettingsInput,
} from '../shared/ai-settings-api'

function definitionFor(providerId: string): AiProviderDefinition | undefined {
  return AI_PROVIDER_DEFINITIONS.find((definition) => definition.id === providerId)
}

function defaults(): PersistedAiSettings {
  const providers = Object.fromEntries(
    AI_PROVIDER_DEFINITIONS.map((definition) => [
      definition.id,
      {
        providerId: definition.id,
        model: definition.models[0] ?? definition.imageModels?.[0] ?? '',
        ...(definition.defaultBaseUrl ? { baseUrl: definition.defaultBaseUrl } : {}),
        enabled: true,
      },
    ]),
  )
  return {
    version: 1,
    active: {
      chat: { providerId: 'genspark', model: AI_PROVIDER_DEFINITIONS[0]?.models[0] ?? '' },
      image: {
        providerId: 'genspark',
        model:
          AI_PROVIDER_DEFINITIONS[0]?.imageModels?.[0] ??
          AI_PROVIDER_DEFINITIONS[0]?.models[0] ??
          '',
      },
    },
    providers,
    updatedAt: new Date().toISOString(),
  }
}

function createStore(settingsPath: string): VersionedAiSettingsStore {
  const adapter: SafeStorageAdapter = {
    isEncryptionAvailable: () => safeStorage.isEncryptionAvailable(),
    encryptString: (value) => safeStorage.encryptString(value),
    decryptString: (value) => safeStorage.decryptString(Buffer.from(value)),
  }
  const root = dirname(settingsPath)
  return new VersionedAiSettingsStore({
    settingsPath: join(root, 'ai-settings.json'),
    credentialStore: new EncryptedCredentialStore(join(root, 'ai-credentials.json'), adapter),
    defaults: defaults(),
  })
}

function toSnapshot(store: VersionedAiSettingsStore): AiSettingsSnapshot {
  const publicSettings = store.load()
  const providers: AiProviderConfigView[] = AI_PROVIDER_DEFINITIONS.map((definition) => {
    const config = publicSettings.providers[definition.id]
    return {
      providerId: definition.id,
      model: config?.model ?? definition.models[0] ?? '',
      baseUrl: config?.baseUrl ?? definition.defaultBaseUrl ?? '',
      credentialSet: config?.credentialConfigured === true,
      ...(config?.credentialHint ? { credentialHint: config.credentialHint } : {}),
      enabled: config?.enabled !== false,
    }
  })
  return {
    activeProvider: publicSettings.active.chat.providerId,
    activeModel: publicSettings.active.chat.model,
    imageProvider: publicSettings.active.image.providerId,
    imageModel: publicSettings.active.image.model,
    providers,
    definitions: AI_PROVIDER_DEFINITIONS.map((definition) => ({
      ...definition,
      models: [...definition.models],
    })),
  }
}

function cleanInput(input: unknown): SaveAiSettingsInput {
  if (!input || typeof input !== 'object' || Array.isArray(input)) return {}
  const value = input as Record<string, unknown>
  const providerValue = value.provider
  const provider =
    providerValue && typeof providerValue === 'object' && !Array.isArray(providerValue)
      ? (providerValue as Record<string, unknown>)
      : undefined
  return {
    activeProvider: typeof value.activeProvider === 'string' ? value.activeProvider : undefined,
    activeModel: typeof value.activeModel === 'string' ? value.activeModel.trim() : undefined,
    imageProvider: typeof value.imageProvider === 'string' ? value.imageProvider : undefined,
    imageModel: typeof value.imageModel === 'string' ? value.imageModel.trim() : undefined,
    provider: provider
      ? {
          providerId: typeof provider.providerId === 'string' ? provider.providerId : '',
          capability:
            provider.capability === 'image' || provider.capability === 'text'
              ? provider.capability
              : undefined,
          operation:
            provider.operation === 'discover' || provider.operation === 'test'
              ? provider.operation
              : undefined,
          model: typeof provider.model === 'string' ? provider.model.trim() : '',
          baseUrl: typeof provider.baseUrl === 'string' ? provider.baseUrl.trim() : undefined,
          apiKey: typeof provider.apiKey === 'string' ? provider.apiKey : undefined,
          clearCredential: provider.clearCredential === true,
          enabled: provider.enabled !== false,
        }
      : undefined,
  }
}

function validateProvider(
  input: SaveAiProviderInput,
  credentialAlreadySet = false,
  requireModel = true,
): AiProviderConnectionResult {
  const definition = definitionFor(input.providerId)
  if (!definition) return { ok: false, message: 'Unknown provider.' }
  if (requireModel && !input.model.trim()) return { ok: false, message: 'Choose a model.' }
  if (definition.needsBaseUrl || input.baseUrl) {
    const url = input.baseUrl?.trim() ?? definition.defaultBaseUrl
    if (!url) return { ok: false, message: 'Enter an endpoint URL.' }
    const validation = validateAiEndpoint(url)
    if (!validation.ok) return { ok: false, message: validation.reason ?? 'Invalid endpoint URL.' }
  }
  if (definition.requiresApiKey && !input.apiKey?.trim() && !credentialAlreadySet) {
    return { ok: false, message: 'Enter an API key or save one first.' }
  }
  return { ok: true, message: 'Provider settings are valid.' }
}

/** Registers the public settings API; credentials are encrypted in main only. */
export function registerAiSettingsIpc(settingsPath: () => string): void {
  const store = createStore(settingsPath())
  const activeTests = new Map<number, AbortController>()
  configureAiSettingsStore(store)
  ipcMain.handle(AI_SETTINGS_CHANNELS.get, () => toSnapshot(store))

  ipcMain.handle(AI_SETTINGS_CHANNELS.save, (_event, raw: unknown) => {
    const input = cleanInput(raw)
    if (input.provider) {
      const validation = validateProvider(
        input.provider,
        store.load().providers[input.provider.providerId]?.credentialConfigured === true,
      )
      if (!validation.ok) throw new Error(validation.message)
      store.updateProvider({
        providerId: input.provider.providerId,
        model: input.provider.model,
        baseUrl: input.provider.baseUrl,
        enabled: input.provider.enabled,
        credential: input.provider.clearCredential ? null : input.provider.apiKey,
      })
    }
    if (input.activeProvider) {
      const current = store.load()
      const provider = current.providers[input.activeProvider]
      if (!provider) throw new Error('Unknown active provider.')
      store.setActive({
        task: 'chat',
        providerId: input.activeProvider,
        model: input.activeModel || provider.model,
      })
    }
    if (input.imageProvider) {
      const current = store.load()
      const provider = current.providers[input.imageProvider]
      if (!provider) throw new Error('Unknown image provider.')
      store.setActive({
        task: 'image',
        providerId: input.imageProvider,
        model: input.imageModel || provider.model,
      })
    }
    return toSnapshot(store)
  })

  ipcMain.handle(AI_SETTINGS_CHANNELS.cancelTest, (event) => {
    activeTests.get(event.sender.id)?.abort()
  })

  ipcMain.handle(AI_SETTINGS_CHANNELS.test, async (event, raw: unknown) => {
    const input = cleanInput({ provider: raw }).provider
    if (!input) return { ok: false, message: 'Invalid provider settings.' }
    const validation = validateProvider(
      input,
      store.load().providers[input.providerId]?.credentialConfigured === true,
      false,
    )
    if (!validation.ok) return validation
    const definition = definitionFor(input.providerId)
    if (!definition || input.providerId === 'genspark') return validation
    if (!definition.supportsModelDiscovery) {
      return { ok: true, message: 'Provider configuration is ready.' }
    }
    const apiKey = input.apiKey || store.resolveProvider(input.providerId).apiKey || ''
    const config = {
      apiKey,
      model: input.model,
      ...(input.baseUrl || definition.defaultBaseUrl
        ? { baseUrl: input.baseUrl || definition.defaultBaseUrl }
        : {}),
    }
    const isConnectionTest = input.operation === 'test'
    const senderId = event.sender.id
    activeTests.get(senderId)?.abort()
    const controller = new AbortController()
    activeTests.set(senderId, controller)
    const abortOnDestroyed = () => controller.abort()
    event.sender.once('destroyed', abortOnDestroyed)
    const discoveryOptions = { timeoutMs: 20_000, signal: controller.signal }
    const runwareOptions = isConnectionTest
      ? { timeoutMs: 30_000, signal: controller.signal, search: input.model, maxPages: 1 }
      : { timeoutMs: 60_000, signal: controller.signal }
    try {
      const catalog = await (input.capability === 'image' && input.providerId === 'openrouter'
        ? discoverOpenRouterImageModels(config, discoveryOptions)
        : input.capability === 'image' && input.providerId === 'fal'
          ? discoverFalImageModels(config, discoveryOptions)
          : input.capability === 'image' && input.providerId === 'runware'
            ? discoverRunwareImageModels(config, runwareOptions)
            : discoverModels(input.providerId as AiProviderId, config, discoveryOptions))
      const compatibleModels = catalog.models.filter((model) =>
        input.capability === 'image'
          ? model.capabilities.imageGeneration === true
          : model.capabilities.chat === true,
      )
      if (isConnectionTest && input.model) {
        const selectedModel = compatibleModels.find((model) => model.id === input.model)
        if (!selectedModel) {
          return {
            ok: false,
            message: `Connection succeeded, but “${input.model}” was not found in the provider catalog. Check the model ID.`,
          }
        }
        return {
          ok: true,
          message: `Connection succeeded and “${input.model}” is available.`,
        }
      }
      return {
        ok: true,
        message: `Connection succeeded; ${compatibleModels.length} compatible models found.`,
        models: compatibleModels.map((model) => model.id),
      }
    } catch (error: unknown) {
      return {
        ok: false,
        message: error instanceof Error ? error.message : String(error),
      }
    } finally {
      event.sender.removeListener('destroyed', abortOnDestroyed)
      if (activeTests.get(senderId) === controller) activeTests.delete(senderId)
    }
  })
}
