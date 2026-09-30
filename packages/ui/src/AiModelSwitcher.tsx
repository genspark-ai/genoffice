/**
 * Quick model switcher for every editor's AI panel header (genoffice#692):
 * one control that lists every model of every provider that would actually
 * work — genspark with a gsk login, BYOK providers with key + model, custom
 * with a base URL, codex once a CLI or model was set — and switches with a
 * single pick instead of a detour through the shell's settings screen.
 *
 * The pick is written through the host (`onSwitch`) on top of a fresh
 * `reload()` read, so switching here never clobbers a change another window
 * made to ai-settings.json after this panel mounted. The host owns its
 * settings state: handing `next` back keeps the trigger and the per-turn
 * transport on the selection the user just made.
 */
import React, { useMemo, useState } from 'react'
import {
  AI_PROVIDERS,
  isProviderConfigured,
  type AiProviderMeta,
  type AiSettings,
} from '@genoffice/ai-provider/browser'
import type { Lang } from '@genoffice/i18n'
import { Dropdown, type DropdownOption } from './dropdown'
import { AI_MODEL_SWITCHER_LABELS } from './strings-ai-model-switcher'

/** one configured provider and the model ids the switcher offers for it */
export interface AiModelSwitcherEntry {
  readonly meta: AiProviderMeta
  readonly models: readonly string[]
}

/**
 * The switcher's list from a settings snapshot: every provider that passes
 * isProviderConfigured, in catalog order, one row per model. Codex and custom
 * keep their single saved id (their catalogs need a live round trip / are free
 * text; '' is codex's "the CLI picks its default"); static catalogs list every
 * vendor model, and a stored id that has fallen off the current catalog stays
 * pickable at the front. The selection currently in use always gets a row,
 * even when it no longer passes the configured bar — it is what the app is
 * running on, so it must stay visible and re-pickable.
 */
export function modelSwitcherEntries(
  settings: AiSettings,
  gskLoggedIn: boolean,
): AiModelSwitcherEntry[] {
  const entries: { meta: AiProviderMeta; models: string[] }[] = []
  for (const meta of AI_PROVIDERS) {
    if (!isProviderConfigured(settings, meta.id, gskLoggedIn)) continue
    const stored = settings.providers[meta.id]?.model?.trim()
    if (meta.needsCliPath || meta.needsBaseUrl) {
      entries.push({ meta, models: [stored ?? ''] })
    } else {
      const models = [...meta.models]
      if (stored && !models.includes(stored)) models.unshift(stored)
      entries.push({ meta, models })
    }
  }
  const activeMeta = AI_PROVIDERS.find((m) => m.id === settings.provider)
  const activeModel = settings.providers[settings.provider]?.model?.trim() ?? ''
  const active = entries.find((entry) => entry.meta.id === settings.provider)
  if (!active) {
    entries.push({
      meta:
        activeMeta ?? {
          id: settings.provider,
          label: settings.provider,
          models: [],
          defaultModel: '',
          keyPlaceholder: '',
        },
      models: [activeModel],
    })
  } else if (!active.models.includes(activeModel)) {
    active.models.unshift(activeModel)
  }
  return entries
}

/** provider and model in one key; split at the first '|' so model ids containing one survive */
const rowKey = (providerId: string, model: string): string => `${providerId}|${model}`

export function AiModelSwitcher({
  lang,
  settings,
  gskLoggedIn,
  reload,
  onSwitch,
}: {
  readonly lang: Lang
  /** the host's current settings; the trigger reads its active row from here */
  readonly settings: AiSettings
  readonly gskLoggedIn: boolean
  /** freshest settings for the pick base — the file may have moved since mount */
  readonly reload: () => Promise<AiSettings>
  /** persist AND apply; the host owns its settings state */
  readonly onSwitch: (next: AiSettings) => void | Promise<void>
}): React.JSX.Element {
  const label = AI_MODEL_SWITCHER_LABELS[lang]
  const [busy, setBusy] = useState(false)
  const entries = useMemo(
    () => modelSwitcherEntries(settings, gskLoggedIn),
    [settings, gskLoggedIn],
  )
  const activeModel = settings.providers[settings.provider]?.model?.trim() ?? ''
  const activeKey = rowKey(settings.provider, activeModel)
  const options = useMemo<ReadonlyArray<DropdownOption<string>>>(
    () =>
      entries.flatMap((entry) =>
        entry.models.map((model): DropdownOption<string> => {
          const key = rowKey(entry.meta.id, model)
          const shown = model || label.auto
          return {
            value: key,
            label: `${entry.meta.label} · ${shown}`,
            render: (
              <span className="ai-model-switcher-row">
                <span className="ai-model-switcher-provider">{entry.meta.label}</span>
                <span className="ai-model-switcher-model">{shown}</span>
              </span>
            ),
          }
        }),
      ),
    [entries, label.auto],
  )
  const pick = async (key: string): Promise<void> => {
    if (busy || key === activeKey) return
    const split = key.indexOf('|')
    const providerId = key.slice(0, split)
    const model = key.slice(split + 1)
    setBusy(true)
    try {
      const fresh = await reload().catch(() => settings)
      const config = fresh.providers[providerId as keyof AiSettings['providers']] ?? {
        apiKey: '',
        model: '',
      }
      await onSwitch({
        ...fresh,
        provider: providerId as AiSettings['provider'],
        providers: {
          ...fresh.providers,
          [providerId]: { ...config, model },
        },
      })
    } finally {
      setBusy(false)
    }
  }
  return (
    <Dropdown
      className="ai-model-switcher"
      value={activeKey}
      ariaLabel={label.switch}
      tip={label.switch}
      disabled={busy}
      options={options}
      onPick={(key) => void pick(key)}
    />
  )
}
