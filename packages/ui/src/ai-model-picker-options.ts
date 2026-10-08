import {
  AI_PROVIDERS,
  activeProvider,
  type AiProviderId,
  type AiSettings,
} from '@genoffice/ai-provider/browser'

export interface AiModelPickerGroup {
  readonly id: AiProviderId
  readonly label: string
  readonly models: readonly string[]
}

export interface AiModelPickerSelection {
  readonly provider: AiProviderId
  readonly model: string
}

/**
 * Providers the composer chip may switch to: Genspark while signed in, every
 * other vendor once its settings pass the same usability test the chat path
 * applies (`activeProvider`), so picking a row never lands on a 401. A model
 * typed into the settings page that is not in the catalog is listed first.
 */
export function aiModelPickerGroups(
  settings: AiSettings,
  gskLoggedIn: boolean,
): AiModelPickerGroup[] {
  const groups: AiModelPickerGroup[] = []
  for (const meta of AI_PROVIDERS) {
    const config = settings.providers?.[meta.id]
    const stored = config?.model?.trim() ?? ''
    // CLI vendors pass activeProvider unconfigured (the binary auto-discovers);
    // list them only once the settings page holds a model or a path
    const usable =
      meta.id === 'genspark'
        ? gskLoggedIn
        : meta.needsCliPath
          ? Boolean(stored || config?.cliPath?.trim())
          : activeProvider({ ...settings, provider: meta.id }) === meta.id
    if (!usable) continue
    const models = stored && !meta.models.includes(stored) ? [stored, ...meta.models] : meta.models
    if (models.length === 0 && !meta.needsCliPath) continue
    groups.push({ id: meta.id, label: meta.label, models })
  }
  return groups
}

export function aiModelPickerSelection(settings: AiSettings): AiModelPickerSelection {
  const provider = activeProvider(settings)
  const meta = AI_PROVIDERS.find((m) => m.id === provider)
  const model = settings.providers?.[provider]?.model?.trim() || meta?.defaultModel || ''
  return { provider, model }
}

export function withAiModelSelection(
  settings: AiSettings,
  pick: AiModelPickerSelection,
): AiSettings {
  const config = settings.providers[pick.provider] ?? { apiKey: '', model: '' }
  return {
    ...settings,
    provider: pick.provider,
    providers: { ...settings.providers, [pick.provider]: { ...config, model: pick.model } },
  }
}
