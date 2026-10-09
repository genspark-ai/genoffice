import {
  AI_PROVIDERS,
  activeProvider,
  customEndpointLabel,
  resolveCustomEndpoints,
  upsertCustomEndpoint,
  type AiProviderId,
  type AiSettings,
} from '@genoffice/ai-provider/browser'

export interface AiModelPickerGroup {
  readonly id: AiProviderId
  /** set on custom groups: one group per saved endpoint */
  readonly endpoint?: string | undefined
  readonly label: string
  readonly models: readonly string[]
}

export interface AiModelPickerSelection {
  readonly provider: AiProviderId
  readonly endpoint?: string | undefined
  readonly model: string
}

/**
 * Providers the composer chip may switch to: Genspark while signed in, every
 * other vendor once its settings pass the same usability test the chat path
 * applies (`activeProvider`), so picking a row never lands on a 401. A model
 * typed into the settings page that is not in the catalog is listed first.
 */
export function aiModelPickerGroups(input: AiSettings, gskLoggedIn: boolean): AiModelPickerGroup[] {
  const settings = resolveCustomEndpoints(input)
  const groups: AiModelPickerGroup[] = []
  for (const meta of AI_PROVIDERS) {
    if (meta.id === 'custom') {
      for (const ep of settings.customEndpoints ?? []) {
        if (!ep.baseUrl || !ep.model) continue
        const models = [ep.model, ...(ep.models ?? []).filter((m) => m !== ep.model)]
        groups.push({
          id: meta.id,
          endpoint: ep.id,
          label: customEndpointLabel(ep, meta.label),
          models,
        })
      }
      continue
    }
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

export function aiModelPickerSelection(input: AiSettings): AiModelPickerSelection {
  const settings = resolveCustomEndpoints(input)
  const provider = activeProvider(settings)
  const meta = AI_PROVIDERS.find((m) => m.id === provider)
  const model = settings.providers?.[provider]?.model?.trim() || meta?.defaultModel || ''
  if (provider === 'custom' && settings.customEndpoint) {
    return { provider, endpoint: settings.customEndpoint, model }
  }
  return { provider, model }
}

export function withAiModelSelection(input: AiSettings, pick: AiModelPickerSelection): AiSettings {
  const settings = resolveCustomEndpoints(input)
  if (pick.provider === 'custom' && pick.endpoint) {
    const ep = settings.customEndpoints?.find((e) => e.id === pick.endpoint)
    if (ep) {
      return {
        ...upsertCustomEndpoint(settings, { ...ep, model: pick.model }, true),
        provider: 'custom',
      }
    }
  }
  const config = settings.providers[pick.provider] ?? { apiKey: '', model: '' }
  return {
    ...settings,
    provider: pick.provider,
    providers: { ...settings.providers, [pick.provider]: { ...config, model: pick.model } },
  }
}
