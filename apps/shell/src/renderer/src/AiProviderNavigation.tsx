import type {
  AiProviderCapability,
  AiProviderDefinition,
  AiSettingsSnapshot,
} from '../../shared/ai-settings-api'
import { useI18n } from './locale'
import { ProviderIcon } from './ProviderIcon'

interface ProviderModeTabsProps {
  capability: AiProviderCapability
  onChange: (capability: AiProviderCapability) => void
}

export function ProviderModeTabs({ capability, onChange }: ProviderModeTabsProps) {
  const { t } = useI18n()
  return (
    <div className="provider-mode-tabs" role="tablist" aria-label={t('aiProviderModeLabel')}>
      <button
        type="button"
        role="tab"
        aria-selected={capability === 'text'}
        className={capability === 'text' ? 'active' : ''}
        onClick={() => onChange('text')}
      >
        <span className="provider-mode-icon" aria-hidden="true">
          Aa
        </span>
        <span>
          <strong>{t('aiTextModelsTab')}</strong>
          <small>{t('aiTextModelsTabHint')}</small>
        </span>
      </button>
      <button
        type="button"
        role="tab"
        aria-selected={capability === 'image'}
        className={capability === 'image' ? 'active' : ''}
        onClick={() => onChange('image')}
      >
        <span className="provider-mode-icon image" aria-hidden="true">
          ▧
        </span>
        <span>
          <strong>{t('aiImageModelsTab')}</strong>
          <small>{t('aiImageModelsTabHint')}</small>
        </span>
      </button>
    </div>
  )
}

interface ProviderListProps {
  providers: AiProviderDefinition[]
  snapshot: AiSettingsSnapshot
  activeProviderId: string
  selectedId: string
  query: string
  onQueryChange: (query: string) => void
  onSelect: (providerId: string) => void
}

export function ProviderList({
  providers,
  snapshot,
  activeProviderId,
  selectedId,
  query,
  onQueryChange,
  onSelect,
}: ProviderListProps) {
  const { t } = useI18n()
  const configured = new Set(
    snapshot.providers
      .filter((provider) => provider.credentialSet)
      .map((provider) => provider.providerId),
  )
  return (
    <aside className="provider-list-panel">
      <div className="provider-list-header">
        <div>
          <p className="settings-eyebrow">{t('aiProviderListLabel')}</p>
          <span>{providers.length}</span>
        </div>
        <label className="provider-search">
          <span className="sr-only">{t('aiSearchProviders')}</span>
          <svg width="15" height="15" viewBox="0 0 20 20" aria-hidden="true">
            <circle cx="8.5" cy="8.5" r="5.5" fill="none" stroke="currentColor" strokeWidth="1.6" />
            <path
              d="m12.5 12.5 4 4"
              stroke="currentColor"
              strokeWidth="1.6"
              strokeLinecap="round"
            />
          </svg>
          <input
            value={query}
            onChange={(event) => onQueryChange(event.target.value)}
            placeholder={t('aiSearchProviders')}
          />
        </label>
      </div>
      <nav className="provider-list-scroll" aria-label={t('aiProviderListLabel')}>
        {providers.map((provider) => {
          const isRuntimeActive = provider.id === activeProviderId
          return (
            <button
              className={`provider-list-item${provider.id === selectedId ? ' active' : ''}`}
              key={provider.id}
              type="button"
              aria-pressed={provider.id === selectedId}
              onClick={() => onSelect(provider.id)}
            >
              <ProviderIcon provider={provider} size={34} />
              <span className="provider-list-copy">
                <span className="provider-list-name">{provider.label}</span>
                <span className="provider-list-protocol">
                  {isRuntimeActive
                    ? t('aiProviderActive')
                    : provider.protocol === 'local'
                      ? t('aiProviderLocal')
                      : provider.protocol === 'openai-compatible'
                        ? t('aiProviderOpenAiCompatible')
                        : t('aiProviderCloud')}
                </span>
              </span>
              {configured.has(provider.id) && (
                <span className="provider-configured-dot" aria-label={t('aiProviderConfigured')} />
              )}
            </button>
          )
        })}
      </nav>
    </aside>
  )
}
