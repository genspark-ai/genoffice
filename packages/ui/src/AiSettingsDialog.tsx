import { useEffect, useState, useCallback } from 'react'
import type { AiSettings, AiProviderId, AiProviderMeta } from '@genoffice/ai-provider'
import { AI_PROVIDERS } from '@genoffice/ai-provider'

export interface OllamaModel {
  name: string
  parameterSize?: string
  modifiedAt?: string
}

export interface AiSettingsDialogStrings {
  aiSettingsProvider: string
  aiSettingsApiKey: string
  aiSettingsApiKeyHint: string
  aiSettingsBaseUrl: string
  aiSettingsDetectedModels: string
  aiSettingsRefresh: string
  aiSettingsNoModel: string
  aiSettingsTestFail: string
  aiSettingsCancel: string
  aiSettingsSave: string
  aiSettingsModel: string
  aiSettingsGensparkLogin: string
  aiSettingsGensparkConnected: string
  aiSettingsGensparkDisconnected: string
  aiSettingsOllamaBaseUrlHint: string
}

export interface AiSettingsDialogProps {
  settings: AiSettings
  strings: AiSettingsDialogStrings
  gskStatus?: { loggedIn: boolean; email?: string } | null
  onGskLogin?: () => void
  listOllamaModels?: (baseUrl: string) => Promise<OllamaModel[]>
  onSettingsChange: (next: AiSettings) => void
  onClose: () => void
}

const OLLAMA_DEFAULT_BASE = 'http://localhost:11434/v1'
const KEY_OPTIONAL_PROVIDERS = new Set(['ollama'])

export function AiSettingsDialog({
  settings,
  strings,
  gskStatus,
  onGskLogin,
  listOllamaModels,
  onSettingsChange,
  onClose,
}: AiSettingsDialogProps) {
  const [provider, setProvider] = useState<AiProviderId>(settings.provider)
  const [apiKey, setApiKey] = useState(() => settings.providers[settings.provider]?.apiKey ?? '')
  const [baseUrl, setBaseUrl] = useState(() => settings.providers[settings.provider]?.baseUrl ?? '')
  const [model, setModel] = useState(() => settings.providers[settings.provider]?.model ?? '')
  const [ollamaModels, setOllamaModels] = useState<OllamaModel[]>([])
  const [ollamaLoading, setOllamaLoading] = useState(false)
  const [ollamaError, setOllamaError] = useState<string | null>(null)

  const meta = AI_PROVIDERS.find((p) => p.id === provider) ?? null

  const fetchOllamaModels = useCallback(
    async (url?: string) => {
      if (!listOllamaModels) return
      const target = url ?? baseUrl ?? OLLAMA_DEFAULT_BASE
      setOllamaLoading(true)
      setOllamaError(null)
      try {
        const models = await listOllamaModels(target)
        setOllamaModels(models)
        const first = models[0]
        if (first && !model) {
          setModel(first.name)
        }
      } catch {
        setOllamaError(strings.aiSettingsTestFail)
      } finally {
        setOllamaLoading(false)
      }
    },
    [listOllamaModels, baseUrl, model, strings.aiSettingsTestFail],
  )

  useEffect(() => {
    if (provider === 'ollama') {
      fetchOllamaModels()
    }
  }, []) // eslint-disable-line react-hooks/exhaustive-deps

  const handleProviderChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const next = e.target.value as AiProviderId
    setProvider(next)
    const nextMeta = AI_PROVIDERS.find((p) => p.id === next)
    if (nextMeta?.defaultBaseUrl) {
      setBaseUrl(nextMeta.defaultBaseUrl)
    } else {
      setBaseUrl('')
    }
    setApiKey('')
    setModel(nextMeta?.defaultModel ?? '')
    setOllamaModels([])
    setOllamaError(null)
  }

  const handleSave = () => {
    const nextProviders = {
      ...settings.providers,
      [provider]: {
        apiKey: apiKey || '',
        model: model || '',
        baseUrl: baseUrl || '',
      },
    }
    onSettingsChange({
      provider,
      providers: nextProviders,
    })
    onClose()
  }

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  const keyOptional = KEY_OPTIONAL_PROVIDERS.has(provider)
  const isOllama = provider === 'ollama'
  const isGenspark = provider === 'genspark'
  const isCustom = provider === 'custom'
  const showApiKey = !isGenspark
  const showModel = !isGenspark
  const showBaseUrl = isCustom || isOllama
  const staticModels = meta?.models ?? []

  return (
    <div
      className="ai-settings-dialog-overlay"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose()
      }}
    >
      <div className="ai-settings-dialog-panel">
        <div className="ai-settings-dialog-header">
          <span className="ai-settings-dialog-title">
            {strings.aiSettingsProvider}
          </span>
          <button
            className="ai-settings-dialog-close"
            onClick={onClose}
            aria-label="Close"
          >
            ✕
          </button>
        </div>

        <div className="ai-settings-dialog-body">
          {/* Provider */}
          <div className="ai-settings-field">
            <label>{strings.aiSettingsProvider}</label>
            <select value={provider} onChange={handleProviderChange}>
              {AI_PROVIDERS.map((p) => (
                <option key={p.id} value={p.id}>
                  {p.label}
                </option>
              ))}
            </select>
          </div>

          {/* Genspark account */}
          {isGenspark && gskStatus && (
            <div className="ai-settings-genspark-account">
              <div className="ai-settings-genspark-status">
                <span
                  className={`ai-settings-genspark-dot ${
                    gskStatus.loggedIn ? 'ai-settings-genspark-dot--connected' : ''
                  }`}
                />
                <span>
                  {gskStatus.loggedIn
                    ? `${strings.aiSettingsGensparkConnected}${gskStatus.email ? ` (${gskStatus.email})` : ''}`
                    : strings.aiSettingsGensparkDisconnected}
                </span>
              </div>
              {onGskLogin && !gskStatus.loggedIn && (
                <button className="ai-settings-genspark-login-btn" onClick={onGskLogin}>
                  {strings.aiSettingsGensparkLogin}
                </button>
              )}
            </div>
          )}

          {/* API Key */}
          {showApiKey && (
            <div className="ai-settings-field">
              <label>{strings.aiSettingsApiKey}</label>
              <input
                type="password"
                value={apiKey}
                onChange={(e) => setApiKey(e.target.value)}
                placeholder={keyOptional ? '(optional)' : ''}
              />
              {keyOptional && (
                <div className="ai-settings-field-hint">
                  {strings.aiSettingsApiKeyHint}
                </div>
              )}
            </div>
          )}

          {/* Base URL */}
          {showBaseUrl && (
            <div className="ai-settings-field">
              <label>{strings.aiSettingsBaseUrl}</label>
              <input
                type="text"
                value={baseUrl}
                onChange={(e) => setBaseUrl(e.target.value)}
                placeholder={meta?.defaultBaseUrl ?? ''}
              />
              {isOllama && (
                <div className="ai-settings-field-hint">
                  {strings.aiSettingsOllamaBaseUrlHint}
                </div>
              )}
            </div>
          )}

          {/* Model */}
          {showModel && (
            <>
              {isOllama ? (
                <div className="ai-settings-model-row">
                  <div className="ai-settings-field">
                    <label>{strings.aiSettingsDetectedModels}</label>
                    <select
                      value={model}
                      onChange={(e) => setModel(e.target.value)}
                    >
                      {ollamaLoading && (
                        <option value="">{strings.aiSettingsRefresh}…</option>
                      )}
                      {!ollamaLoading && ollamaModels.length === 0 && (
                        <option value="">{strings.aiSettingsNoModel}</option>
                      )}
                      {ollamaModels.map((m) => (
                        <option key={m.name} value={m.name}>
                          {m.name}
                          {m.parameterSize ? ` (${m.parameterSize})` : ''}
                        </option>
                      ))}
                    </select>
                    {ollamaError && (
                      <div className="ai-settings-field-hint" style={{ color: 'var(--error)' }}>
                        {ollamaError}
                      </div>
                    )}
                  </div>
                  <button
                    className="ai-settings-model-refresh"
                    onClick={() => fetchOllamaModels()}
                    disabled={ollamaLoading}
                  >
                    {strings.aiSettingsRefresh}
                  </button>
                </div>
              ) : isCustom ? (
                <div className="ai-settings-field">
                  <label>{strings.aiSettingsModel}</label>
                  <input
                    type="text"
                    value={model}
                    onChange={(e) => setModel(e.target.value)}
                    placeholder="model-name"
                  />
                </div>
              ) : (
                staticModels.length > 0 && (
                  <div className="ai-settings-field">
                    <label>{strings.aiSettingsModel}</label>
                    <select
                      value={model}
                      onChange={(e) => setModel(e.target.value)}
                    >
                      {staticModels.map((m) => (
                        <option key={m} value={m}>
                          {m}
                        </option>
                      ))}
                    </select>
                  </div>
                )
              )}
            </>
          )}
        </div>

        <div className="ai-settings-dialog-footer">
          <button className="ai-settings-dialog-cancel" onClick={onClose}>
            {strings.aiSettingsCancel}
          </button>
          <button className="ai-settings-dialog-save" onClick={handleSave}>
            {strings.aiSettingsSave}
          </button>
        </div>
      </div>
    </div>
  )
}

export type { IconSettings } from './icons'
