import { useEffect, useMemo, useRef, useState } from 'react'
import type {
  AiProviderCapability,
  AiProviderConfigView,
  AiProviderConnectionResult,
  AiProviderDefinition,
  AiSettingsSnapshot,
} from '../../shared/ai-settings-api'
import { AI_PROVIDER_DEFINITIONS } from '../../shared/ai-settings-api'
import { ProviderList, ProviderModeTabs } from './AiProviderNavigation'
import { useI18n } from './locale'
import { ProviderIcon } from './ProviderIcon'

const CUSTOM_MODEL_VALUE = '__genoffice_custom_model__'

const FALLBACK_SNAPSHOT: AiSettingsSnapshot = {
  activeProvider: 'genspark',
  activeModel: AI_PROVIDER_DEFINITIONS[0]?.models[0] ?? '',
  imageProvider: 'genspark',
  imageModel:
    AI_PROVIDER_DEFINITIONS[0]?.imageModels?.[0] ?? AI_PROVIDER_DEFINITIONS[0]?.models[0] ?? '',
  providers: AI_PROVIDER_DEFINITIONS.map((definition) => ({
    providerId: definition.id,
    model: definition.models[0] ?? definition.imageModels?.[0] ?? '',
    baseUrl: definition.defaultBaseUrl ?? '',
    credentialSet: false,
    enabled: true,
  })),
  definitions: AI_PROVIDER_DEFINITIONS.map((definition) => ({
    ...definition,
    models: [...definition.models],
    imageModels: [...(definition.imageModels ?? [])],
  })),
}

function configFor(snapshot: AiSettingsSnapshot, providerId: string): AiProviderConfigView {
  return (
    snapshot.providers.find((config) => config.providerId === providerId) ?? {
      providerId,
      model: '',
      baseUrl: '',
      credentialSet: false,
      enabled: true,
    }
  )
}

function definitionFor(snapshot: AiSettingsSnapshot, providerId: string): AiProviderDefinition {
  return (
    snapshot.definitions.find((definition) => definition.id === providerId) ??
    AI_PROVIDER_DEFINITIONS.find((definition) => definition.id === providerId) ??
    AI_PROVIDER_DEFINITIONS[0]
  )
}

function modelsFor(definition: AiProviderDefinition, capability: AiProviderCapability): string[] {
  return capability === 'image' ? (definition.imageModels ?? []) : definition.models
}

function supportsCapability(
  definition: AiProviderDefinition,
  capability: AiProviderCapability,
): boolean {
  return capability === 'image' ? definition.supportsImages : definition.supportsText !== false
}

export function AiProvidersPage() {
  const { t } = useI18n()
  const [snapshot, setSnapshot] = useState<AiSettingsSnapshot>(FALLBACK_SNAPSHOT)
  const [capability, setCapability] = useState<AiProviderCapability>('text')
  const [selectedId, setSelectedId] = useState('genspark')
  const [model, setModel] = useState('')
  const [customModelEnabled, setCustomModelEnabled] = useState(false)
  const [customModel, setCustomModel] = useState('')
  const [baseUrl, setBaseUrl] = useState('')
  const [apiKey, setApiKey] = useState('')
  const [providerQuery, setProviderQuery] = useState('')
  const [saving, setSaving] = useState(false)
  const [testing, setTesting] = useState(false)
  const [activeOperation, setActiveOperation] = useState<'discover' | 'test' | null>(null)
  const [status, setStatus] = useState<AiProviderConnectionResult | null>(null)
  const [loadError, setLoadError] = useState<string | null>(null)
  const [loaded, setLoaded] = useState(false)
  const autoDiscovered = useRef(new Set<string>())

  const definition = useMemo(() => definitionFor(snapshot, selectedId), [snapshot, selectedId])
  const selectedConfig = useMemo(() => configFor(snapshot, selectedId), [snapshot, selectedId])
  const availableModels = useMemo(() => modelsFor(definition, capability), [definition, capability])
  const effectiveModel = customModelEnabled ? customModel.trim() : model.trim()
  const providers = useMemo(() => {
    const query = providerQuery.trim().toLowerCase()
    return snapshot.definitions.filter(
      (provider) =>
        supportsCapability(provider, capability) &&
        (!query ||
          provider.label.toLowerCase().includes(query) ||
          provider.description.toLowerCase().includes(query)),
    )
  }, [snapshot.definitions, capability, providerQuery])

  const selectedModelFor = (
    providerId: string,
    nextSnapshot: AiSettingsSnapshot,
    nextCapability: AiProviderCapability,
  ) => {
    const nextDefinition = definitionFor(nextSnapshot, providerId)
    const curated = modelsFor(nextDefinition, nextCapability)
    if (nextCapability === 'text' && nextSnapshot.activeProvider === providerId) {
      const active = nextSnapshot.activeModel
      return active || configFor(nextSnapshot, providerId).model || curated[0] || ''
    }
    if (nextCapability === 'image' && nextSnapshot.imageProvider === providerId) {
      const active = nextSnapshot.imageModel
      return active || configFor(nextSnapshot, providerId).model || curated[0] || ''
    }
    const configured = configFor(nextSnapshot, providerId).model
    return configured || curated[0] || ''
  }

  const syncForm = (providerId: string, nextSnapshot = snapshot, nextCapability = capability) => {
    const config = configFor(nextSnapshot, providerId)
    const nextDefinition = definitionFor(nextSnapshot, providerId)
    const curated = modelsFor(nextDefinition, nextCapability)
    const nextModel = selectedModelFor(providerId, nextSnapshot, nextCapability)
    const isCustomModel = Boolean(curated.length && nextModel && !curated.includes(nextModel))
    setSelectedId(providerId)
    setModel(nextModel)
    setCustomModelEnabled(isCustomModel)
    setCustomModel(isCustomModel ? nextModel : '')
    setBaseUrl(config.baseUrl || nextDefinition.defaultBaseUrl || '')
    setApiKey('')
    setStatus(null)
  }

  useEffect(() => {
    let active = true
    void window.aiOfficeAiSettings
      .get()
      .then((next) => {
        if (!active) return
        setSnapshot(next)
        syncForm(next.activeProvider, next, 'text')
        setLoaded(true)
      })
      .catch(() => {
        if (active) setLoadError(t('aiSettingsLoadFailed'))
      })
    return () => {
      active = false
    }
    // Initial read only; locale changes do not change persisted settings.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const applyDiscoveredModels = (
    providerId: string,
    nextCapability: AiProviderCapability,
    models: string[],
  ) => {
    const uniqueModels = [...new Set(models)].sort((a, b) => a.localeCompare(b))
    setSnapshot((current) => ({
      ...current,
      definitions: current.definitions.map((provider) =>
        provider.id === providerId
          ? nextCapability === 'image'
            ? { ...provider, imageModels: uniqueModels }
            : { ...provider, models: uniqueModels }
          : provider,
      ),
    }))
    setModel((current) => current || uniqueModels[0] || '')
    if (effectiveModel && !uniqueModels.includes(effectiveModel)) {
      setCustomModelEnabled(true)
      setCustomModel(effectiveModel)
    }
  }

  const discover = async (quiet = false, operation: 'discover' | 'test' = 'discover') => {
    setTesting(true)
    setActiveOperation(operation)
    if (!quiet) setStatus(null)
    try {
      const result = await window.aiOfficeAiSettings.test({
        providerId: selectedId,
        capability,
        operation,
        model: effectiveModel,
        baseUrl: baseUrl.trim() || undefined,
        apiKey: apiKey.trim() || undefined,
      })
      if (operation === 'discover' && result.ok && result.models?.length) {
        applyDiscoveredModels(selectedId, capability, result.models)
      }
      if (!quiet || !result.ok) setStatus(result)
    } catch {
      if (!quiet) setStatus({ ok: false, message: t('aiSettingsTestFailed') })
    } finally {
      setTesting(false)
      setActiveOperation(null)
    }
  }

  const cancelTest = async () => {
    try {
      await window.aiOfficeAiSettings.cancelTest()
    } catch {
      setStatus({ ok: false, message: t('aiCancelTestFailed') })
    }
  }

  useEffect(() => {
    if (!loaded || !definition.supportsModelDiscovery || !selectedConfig.credentialSet) return
    // Runware's merged catalog is large. Keep it manual so opening settings never starts a crawl.
    if (selectedId === 'runware') return
    const key = `${capability}:${selectedId}`
    if (autoDiscovered.current.has(key)) return
    autoDiscovered.current.add(key)
    void discover(true)
    // Discovery intentionally runs once per provider and capability in a dialog session.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [loaded, capability, selectedId, selectedConfig.credentialSet])

  const switchCapability = (nextCapability: AiProviderCapability) => {
    setCapability(nextCapability)
    setProviderQuery('')
    const providerId = nextCapability === 'text' ? snapshot.activeProvider : snapshot.imageProvider
    syncForm(providerId, snapshot, nextCapability)
  }

  const save = async () => {
    setSaving(true)
    setStatus(null)
    try {
      const selection =
        capability === 'text'
          ? { activeProvider: selectedId, activeModel: effectiveModel }
          : { imageProvider: selectedId, imageModel: effectiveModel }
      const next = await window.aiOfficeAiSettings.save({
        ...selection,
        provider: {
          providerId: selectedId,
          capability,
          model: effectiveModel,
          baseUrl: baseUrl.trim() || undefined,
          apiKey: apiKey || undefined,
        },
      })
      setSnapshot(next)
      setApiKey('')
      setStatus({ ok: true, message: t('aiSettingsSaved') })
    } catch (error) {
      setStatus({
        ok: false,
        message: error instanceof Error ? error.message : t('aiSettingsSaveFailed'),
      })
    } finally {
      setSaving(false)
    }
  }

  const clearCredential = async () => {
    setSaving(true)
    try {
      const next = await window.aiOfficeAiSettings.save({
        provider: {
          providerId: selectedId,
          capability,
          model: effectiveModel,
          baseUrl: baseUrl.trim() || undefined,
          clearCredential: true,
        },
      })
      setSnapshot(next)
      setApiKey('')
      setStatus({ ok: true, message: t('aiCredentialRemoved') })
    } catch {
      setStatus({ ok: false, message: t('aiSettingsSaveFailed') })
    } finally {
      setSaving(false)
    }
  }

  const requiresNewCredential = definition.requiresApiKey && !selectedConfig.credentialSet
  const saveDisabled =
    saving || testing || !effectiveModel || (requiresNewCredential && !apiKey.trim())
  const activeProviderId = capability === 'text' ? snapshot.activeProvider : snapshot.imageProvider

  return (
    <div className="ai-providers-page">
      <header className="settings-page-intro">
        <div>
          <h2>{t('aiProvidersTitle')}</h2>
          <p>{t('aiProvidersDescription')}</p>
        </div>
        {loadError && (
          <p className="settings-alert" role="alert">
            {loadError}
          </p>
        )}
      </header>

      <ProviderModeTabs capability={capability} onChange={switchCapability} />

      <div className="provider-settings-layout" role="tabpanel">
        <ProviderList
          providers={providers}
          snapshot={snapshot}
          activeProviderId={activeProviderId}
          selectedId={selectedId}
          query={providerQuery}
          onQueryChange={setProviderQuery}
          onSelect={syncForm}
        />

        <section className="provider-editor" aria-labelledby="provider-editor-heading">
          <div className="provider-editor-heading">
            <ProviderIcon provider={definition} size={46} />
            <div>
              <div className="provider-heading-line">
                <h3 id="provider-editor-heading">{definition.label}</h3>
                {selectedId === activeProviderId && (
                  <span className="active-provider-pill">{t('aiProviderActive')}</span>
                )}
              </div>
              <p>{definition.description}</p>
            </div>
          </div>

          <div className="provider-editor-scroll">
            <label className="settings-field">
              <span className="model-field-label">
                <span>{t('aiModelLabel')}</span>
                {availableModels.length > 0 && (
                  <span className="model-count">
                    {t('aiModelsAvailable').replace('{count}', String(availableModels.length))}
                  </span>
                )}
              </span>
              <div className="model-control-row">
                {availableModels.length > 0 ? (
                  <select
                    value={customModelEnabled ? CUSTOM_MODEL_VALUE : model}
                    onChange={(event) => {
                      const value = event.target.value
                      if (value === CUSTOM_MODEL_VALUE) {
                        setCustomModelEnabled(true)
                        setCustomModel(availableModels.includes(model) ? '' : model)
                      } else {
                        setCustomModelEnabled(false)
                        setCustomModel('')
                        setModel(value)
                      }
                      setStatus(null)
                    }}
                    aria-describedby={`model-help-${selectedId}-${capability}`}
                  >
                    {availableModels.map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))}
                    <option value={CUSTOM_MODEL_VALUE}>{t('aiCustomModelOption')}</option>
                  </select>
                ) : (
                  <input
                    value={model}
                    onChange={(event) => setModel(event.target.value)}
                    placeholder={t('aiModelPlaceholder')}
                  />
                )}
                {definition.supportsModelDiscovery && (
                  <button
                    type="button"
                    className="model-refresh-button"
                    onClick={() => void discover(false)}
                    disabled={testing || (!selectedConfig.credentialSet && !apiKey.trim())}
                    aria-label={t('aiRefreshModels')}
                    title={t('aiRefreshModels')}
                  >
                    <svg viewBox="0 0 20 20" aria-hidden="true">
                      <path
                        d="M15.7 7A6 6 0 1 0 16 12"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="1.7"
                        strokeLinecap="round"
                      />
                      <path
                        d="m13.2 4.8 2.7 2.4 1.8-3.1"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="1.7"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      />
                    </svg>
                  </button>
                )}
              </div>
              <small id={`model-help-${selectedId}-${capability}`}>
                {testing
                  ? activeOperation === 'test'
                    ? t('aiTestingModel')
                    : t('aiLoadingModels')
                  : definition.supportsModelDiscovery
                    ? selectedId === 'runware'
                      ? t('aiRunwareDiscoveryHint')
                      : t('aiModelDiscoveryHint')
                    : t('aiCuratedModelsHint')}
              </small>
            </label>

            {availableModels.length > 0 && customModelEnabled && (
              <label className="settings-field custom-model-field">
                <span>{t('aiCustomModelLabel')}</span>
                <input
                  value={customModel}
                  onChange={(event) => {
                    setCustomModel(event.target.value)
                    setStatus(null)
                  }}
                  placeholder={t('aiCustomModelPlaceholder')}
                  autoFocus
                  autoComplete="off"
                  spellCheck={false}
                />
                <small>{t('aiCustomModelHint')}</small>
              </label>
            )}

            {definition.needsBaseUrl && (
              <label className="settings-field">
                <span>{t('aiEndpointLabel')}</span>
                <input
                  value={baseUrl}
                  onChange={(event) => setBaseUrl(event.target.value)}
                  placeholder="https://…/v1"
                  inputMode="url"
                />
                <small>{t('aiEndpointHint')}</small>
              </label>
            )}

            {definition.requiresApiKey && (
              <div className="settings-field">
                <label htmlFor={`provider-key-${selectedId}`}>{t('aiApiKeyLabel')}</label>
                <div className="credential-control">
                  <input
                    id={`provider-key-${selectedId}`}
                    type="password"
                    value={apiKey}
                    onChange={(event) => setApiKey(event.target.value)}
                    placeholder={
                      selectedConfig.credentialHint ||
                      (selectedConfig.credentialSet
                        ? t('aiApiKeySavedPlaceholder')
                        : t('aiApiKeyPlaceholder'))
                    }
                    autoComplete="off"
                    spellCheck={false}
                  />
                  {selectedConfig.credentialSet && (
                    <button type="button" onClick={() => void clearCredential()} disabled={saving}>
                      {t('aiRemoveKey')}
                    </button>
                  )}
                </div>
                <small>
                  {selectedConfig.credentialSet
                    ? `${t('aiApiKeySaved')} ${selectedConfig.credentialHint ?? ''}`
                    : t('aiApiKeyHint')}
                </small>
              </div>
            )}

            {!definition.requiresApiKey && definition.protocol === 'local' && (
              <div className="local-provider-note">
                <span aria-hidden="true">⌁</span>
                <p>{t('aiLocalProviderHint')}</p>
              </div>
            )}
          </div>

          <footer className="provider-editor-footer">
            <div className="provider-runtime-summary">
              <span>{t('aiCurrentSelection')}</span>
              <strong>{effectiveModel || t('aiNoModelSelected')}</strong>
            </div>
            <div className="provider-editor-actions">
              <button
                type="button"
                className="settings-button secondary"
                onClick={() => (testing ? void cancelTest() : void discover(false, 'test'))}
                disabled={
                  saving ||
                  (!testing &&
                    (!effectiveModel || (!selectedConfig.credentialSet && !apiKey.trim())))
                }
              >
                {testing ? t('aiCancelTest') : t('aiTestConnection')}
              </button>
              <button
                type="button"
                className="settings-button primary"
                onClick={() => void save()}
                disabled={saveDisabled}
              >
                {saving ? t('aiSaving') : t('aiSaveProvider')}
              </button>
            </div>
            {status && (
              <p
                className={`settings-status ${status.ok ? 'success' : 'error'}`}
                role={status.ok ? 'status' : 'alert'}
                aria-live={status.ok ? 'polite' : 'assertive'}
              >
                {status.message}
              </p>
            )}
          </footer>
        </section>
      </div>
    </div>
  )
}
