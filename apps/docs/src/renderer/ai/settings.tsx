import { useState } from 'react'
import type { AiProviderConfig, AiProviderId, AiSettings } from '../../shared/ipc'
import { AI_PROVIDERS } from '../../shared/ipc'
import { useI18n } from '../i18n/locale'
import { IconClose } from '../components/icons'

/**
 * AI provider/model settings dialog. Editing happens on a local draft; the
 * save button persists the result (ai:set-settings → ai-settings.json) and
 * hands it back so the app state and the streaming transport see the change.
 */
export function AiSettingsDialog({
  settings,
  onSave,
  onClose,
}: {
  settings: AiSettings
  onSave: (next: AiSettings) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [provider, setProvider] = useState<AiProviderId>(settings.provider)
  const [providers, setProviders] = useState<Record<AiProviderId, AiProviderConfig>>(
    settings.providers,
  )

  const meta = AI_PROVIDERS.find((p) => p.id === provider) ?? AI_PROVIDERS[0]!
  const config = providers[provider] ?? { apiKey: '', model: '', baseUrl: '' }

  const update = (patch: Partial<AiProviderConfig>) => {
    setProviders((prev) => ({
      ...prev,
      [provider]: { ...(prev[provider] ?? { apiKey: '', model: '', baseUrl: '' }), ...patch },
    }))
  }

  const noKeyProvider = provider === 'genspark' || provider === 'opencode'

  const save = () => {
    onSave({ provider, providers })
  }

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" role="dialog" aria-label={t('aiSettingsTitle')} onClick={(e) => e.stopPropagation()}>
        <h2>
          {t('aiSettingsTitle')}
          <button className="modal-close" onClick={onClose} aria-label={t('aiSettingsCancel')}>
            <IconClose size={16} />
          </button>
        </h2>

        <div className="provider-tabs" role="tablist">
          {AI_PROVIDERS.map((p) => (
            <button
              key={p.id}
              role="tab"
              aria-selected={p.id === provider}
              className={`provider-tab${p.id === provider ? ' provider-tab-active' : ''}`}
              onClick={() => setProvider(p.id)}
            >
              {p.label}
            </button>
          ))}
        </div>

        {noKeyProvider && (
          <p className="provider-key-hint">
            {provider === 'genspark' ? t('aiSettingsGensparkHint') : t('aiSettingsOpenCodeHint')}
          </p>
        )}

        {meta.models.length > 0 ? (
          <label>
            {t('aiSettingsModel')}
            <select value={config.model} onChange={(e) => update({ model: e.target.value })}>
              {meta.models.map((m) => (
                <option key={m} value={m}>
                  {m}
                </option>
              ))}
            </select>
          </label>
        ) : (
          <label>
            {t('aiSettingsModel')}
            <input
              value={config.model}
              onChange={(e) => update({ model: e.target.value })}
              placeholder={provider === 'opencode' ? t('aiSettingsModelPlaceholder') : t('aiSettingsCustomModelPlaceholder')}
              spellCheck={false}
            />
          </label>
        )}

        {!noKeyProvider && (
          <label>
            {t('aiSettingsApiKey')}
            <input
              type="password"
              value={config.apiKey}
              onChange={(e) => update({ apiKey: e.target.value })}
              placeholder={meta.keyPlaceholder}
              autoComplete="off"
              spellCheck={false}
            />
          </label>
        )}

        {meta.needsBaseUrl && (
          <label>
            {t('aiSettingsBaseUrl')}
            <input
              value={config.baseUrl ?? ''}
              onChange={(e) => update({ baseUrl: e.target.value })}
              placeholder={provider === 'opencode' ? 'http://127.0.0.1:3456' : 'https://api.example.com/v1'}
              spellCheck={false}
            />
          </label>
        )}

        <div className="modal-actions">
          <button onClick={onClose}>{t('aiSettingsCancel')}</button>
          <button className="btn-primary" onClick={save}>
            {t('aiSettingsSave')}
          </button>
        </div>
      </div>
    </div>
  )
}
