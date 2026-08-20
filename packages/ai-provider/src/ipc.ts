/**
 * Main-process IPC registration shared by the docs/shell, sheets and slides
 * apps (this module is only imported from main processes — never from a
 * renderer bundle; it is exported as a separate `@genoffice/ai-provider/ipc`
 * subpath so the electron import never leaks into renderer code).
 *
 * Registers the four provider-configuration channels exactly once per app:
 * - ai:get-settings / ai:set-settings  — read/persist the AI settings file
 * - ai:ollama-models                  — /api/tags discovery (auto-detected model list)
 * - ai:test-connection                — least-invasive reachability/auth probe
 */
import { ipcMain, type IpcMainInvokeEvent } from 'electron'
import { listOllamaModels, type OllamaModelsResult } from './ollama'
import { testProviderConnection, type AiConnectionTestInput } from './connection'
import { defaultAiSettings, resolveAiSettings } from './providers'
import type { AiProviderConfig, AiProviderId, AiSettings, LegacyAiSettings } from './types'

export interface AiSettingsIpcDeps {
  settingsPath: () => string
  readJson: <T>(path: string, fallback: T) => T
  writeJson: (path: string, value: unknown) => void
  /**
   * Genspark's key lives in the gsk login state, never the settings file; the
   * connection test injects it for the genspark provider. Omit only if the
   * app has no gsk login (then the test reports auth failure as-is).
   */
  gensparkApiKey?: () => string
  /** optional per-app guard before settings access (e.g. sheets session check) */
  beforeAccess?: (event: IpcMainInvokeEvent) => void
  /** optional input validation before persisting (sheets validates with a zod schema) */
  validateSettings?: (input: unknown) => AiSettings
}

/**
 * Register the provider-configuration IPC channels. Call once per main
 * process; the channel names are the shared app-wide contract (the shell
 * registers them for all editors, sheets/slides register their own copy in
 * standalone mode).
 */
export function registerAiSettingsIpc(
  deps: AiSettingsIpcDeps,
  ipc: Pick<typeof ipcMain, 'handle'> = ipcMain,
): void {
  ipc.handle('ai:get-settings', (event): AiSettings => {
    deps.beforeAccess?.(event)
    const stored = deps.readJson<Partial<AiSettings> & LegacyAiSettings>(deps.settingsPath(), {})
    return resolveAiSettings(stored, defaultAiSettings())
  })

  ipc.handle('ai:set-settings', (event, input: unknown): void => {
    deps.beforeAccess?.(event)
    const settings = deps.validateSettings ? deps.validateSettings(input) : (input as AiSettings)
    deps.writeJson(deps.settingsPath(), settings)
  })

  ipc.handle(
    'ai:ollama-models',
    async (_event, baseUrl?: string): Promise<OllamaModelsResult> => {
      try {
        return await listOllamaModels(baseUrl)
      } catch (err) {
        return { models: [], error: String(err) }
      }
    },
  )

  ipc.handle('ai:test-connection', async (_event, input: unknown) => {
    const raw = (input ?? {}) as Partial<AiConnectionTestInput>
    const provider = raw.provider as AiProviderId | undefined
    if (!provider) return { ok: false as const, status: 'unknown' as const }
    let config: AiProviderConfig = {
      apiKey: typeof raw.apiKey === 'string' ? raw.apiKey : '',
      model: typeof raw.model === 'string' ? raw.model : '',
      baseUrl: typeof raw.baseUrl === 'string' ? raw.baseUrl : undefined,
    }
    // the genspark key lives in the gsk login state, never the settings file
    if (provider === 'genspark' && !config.apiKey && deps.gensparkApiKey) {
      config = { ...config, apiKey: deps.gensparkApiKey() }
    }
    return testProviderConnection(provider, config)
  })
}
