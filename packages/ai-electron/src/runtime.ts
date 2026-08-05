import type { AiTask, ResolvedAiProvider } from './types.js'
import type { VersionedAiSettingsStore } from './settings.js'

let configuredStore: VersionedAiSettingsStore | undefined

/** Configure the process-wide main-process settings owner once Electron is ready. */
export function configureAiSettingsStore(store: VersionedAiSettingsStore): void {
  configuredStore = store
}

/** Resolve credentials only inside the main process. */
export function resolveConfiguredAiProvider(task: AiTask): ResolvedAiProvider | undefined {
  return configuredStore?.resolve(task)
}

export function getConfiguredAiSettingsStore(): VersionedAiSettingsStore | undefined {
  return configuredStore
}
