import { useEffect, useState, useCallback } from 'react'
import type {
  AiConnectionStatus,
  AiConnectionTestResult,
  AiSettings,
  AiProviderId,
  OllamaModelsResult,
} from '@genoffice/ai-provider'
import { AI_PROVIDERS, ollamaListStatus } from '@genoffice/ai-provider'

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
  aiSettingsModelMissing: string
  aiSettingsTestFail: string
  aiSettingsCancel: string
  aiSettingsSave: string
  aiSettingsModel: string
  aiSettingsGensparkLogin: string
  aiSettingsGensparkConnected: string
  aiSettingsGensparkDisconnected: string
  aiSettingsOllamaBaseUrlHint: string
  aiSettingsTestButton: string
  aiSettingsTestConnected: string
  aiSettingsTestNotRunning: string
  aiSettingsTestRefused: string
  aiSettingsTestInvalid: string
  aiSettingsTestAuth: string
  aiSettingsTestTimeout: string
  aiSettingsTestFailed: string
}

export interface AiSettingsDialogProps {
  settings: AiSettings
  strings: AiSettingsDialogStrings
  gskStatus?: { loggedIn: boolean; email?: string } | null
  onGskLogin?: () => void
  listOllamaModels?: (baseUrl: string) => Promise<OllamaModelsResult>
  onTestConnection?: (
    provider: AiProviderId,
    input: { baseUrl?: string; apiKey?: string; model?: string },
  ) => Promise<AiConnectionTestResult>
  onSettingsChange: (next: AiSettings) => void
  onClose: () => void
}

const OLLAMA_DEFAULT_BASE = 'http://localhost:11434/v1'
const KEY_OPTIONAL_PROVIDERS = new Set(['ollama'])

/** Model discovery is cached briefly so reopening the dialog (or re-selecting
    Ollama) doesn't hammer the local /api/tags endpoint; the refresh button
    forces a fresh probe. Errors are never cached. */
const OLLAMA_LIST_TTL_MS = 10_000
const ollamaListCache = new Map<string, { at: number; result: OllamaModelsResult }>()

export function AiSettingsDialog({
  settings,
  strings,
  gskStatus,
  onGskLogin,
  listOllamaModels,
  onTestConnection,
  onSettingsChange,
  onClose,
}: AiSettingsDialogProps) {
  const [provider, setProvider] = useState<AiProviderId>(settings.provider)
  const [apiKey, setApiKey] = useState(() => settings.providers[settings.provider]?.apiKey ?? '')
  const [baseUrl, setBaseUrl] = useState(() => settings.providers[settings.provider]?.baseUrl ?? '')
  const [model, setModel] = useState(() => settings.providers[settings.provider]?.model ?? '')
  const [ollamaModels, setOllamaModels] = useState<OllamaModel[]>([])
  const [ollamaLoading, setOllamaLoading] = useState(false)
  /** derived from the last discovery: connected / not-running / invalid / unknown */
  const [ollamaStatus, setOllamaStatus] = useState<AiConnectionStatus | null>(null)
  const [testState, setTestState] = useState<AiConnectionTestResult | null>(null)
  const [testing, setTesting] = useState(false)

  const meta = AI_PROVIDERS.find((p) => p.id === provider) ?? null

  const fetchOllamaModels = useCallback(
    async (force = false) => {
      if (!listOllamaModels) return
      const target = (baseUrl || OLLAMA_DEFAULT_BASE).replace(/\/+$/, '')
      const cached = ollamaListCache.get(target)
      if (!force && cached && Date.now() - cached.at < OLLAMA_LIST_TTL_MS) {
        const result = cached.result
        setOllamaModels(result.models)
        setOllamaStatus(ollamaListStatus(result))
        return
      }
      setOllamaLoading(true)
      setOllamaStatus(null)
      try {
        const result = await listOllamaModels(target)
        // errors are not cached: a transient failure must retry on next open
        if (!result.error) ollamaListCache.set(target, { at: Date.now(), result })
        setOllamaModels(result.models)
        setOllamaStatus(ollamaListStatus(result))
        const first = result.models[0]
        if (first && !model) {
          setModel(first.name)
        }
      } finally {
        setOllamaLoading(false)
      }
    },
    [listOllamaModels, baseUrl, model],
  )

  // discovery runs when Ollama is selected — on open and when the user switches
  // to Ollama mid-dialog (the brief TTL cache covers rapid re-selection)
  useEffect(() => {
    if (provider === 'ollama') {
      fetchOllamaModels()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [provider])

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
    setOllamaStatus(null)
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
  const statusText: Record<AiConnectionStatus, string> = {
    connected: strings.aiSettingsTestConnected,
    'not-running': strings.aiSettingsTestNotRunning,
    refused: strings.aiSettingsTestRefused,
    invalid: strings.aiSettingsTestInvalid,
    auth: strings.aiSettingsTestAuth,
    timeout: strings.aiSettingsTestTimeout,
    unknown: strings.aiSettingsTestFailed,
  }
  const handleTestConnection = async () => {
    if (!onTestConnection || testing) return
    setTesting(true)
    setTestState(null)
    try {
      const input: { baseUrl?: string; apiKey?: string; model?: string } = {}
      if (baseUrl) input.baseUrl = baseUrl
      if (apiKey) input.apiKey = apiKey
      if (model) input.model = model
      setTestState(await onTestConnection(provider, input))
    } catch {
      setTestState({ ok: false, status: 'unknown' })
    } finally {
      setTesting(false)
    }
  }
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
                <>
                  {ollamaLoading || ollamaStatus ? (
                    <div className={`ai-settings-ollama-status ai-settings-ollama-status--${ollamaStatus ?? 'loading'}`}>
                      {ollamaLoading
                        ? `${strings.aiSettingsTestButton}…`
                        : statusText[ollamaStatus!]}
                    </div>
                  ) : null}
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
                        {!ollamaLoading &&
                          model &&
                          !ollamaModels.some((m) => m.name === model) && (
                            <option value={model}>{model}</option>
                          )}
                        {ollamaModels.map((m) => (
                          <option key={m.name} value={m.name}>
                            {m.name}
                            {m.parameterSize ? ` (${m.parameterSize})` : ''}
                          </option>
                        ))}
                      </select>
                      {!ollamaLoading &&
                        model &&
                        ollamaModels.length > 0 &&
                        !ollamaModels.some((m) => m.name === model) && (
                          <div className="ai-settings-field-hint" style={{ color: 'var(--error)' }}>
                            {strings.aiSettingsModelMissing}
                          </div>
                        )}
                    </div>
                    <button
                      className="ai-settings-model-refresh"
                      onClick={() => fetchOllamaModels(true)}
                      disabled={ollamaLoading}
                    >
                      {strings.aiSettingsRefresh}
                    </button>
                  </div>
                </>
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
          {onTestConnection && (
            <div className="ai-settings-test-row">
              <button
                className="ai-settings-test-btn"
                onClick={handleTestConnection}
                disabled={testing}
              >
                {testing ? strings.aiSettingsTestButton + '…' : strings.aiSettingsTestButton}
              </button>
              {testState && (
                <span
                  className={`ai-settings-test-status ${
                    testState.ok ? 'ai-settings-test-status--ok' : 'ai-settings-test-status--fail'
                  }`}
                >
                  {statusText[testState.status]}
                </span>
              )}
            </div>
          )}
          <div className="ai-settings-dialog-actions">
            <button className="ai-settings-dialog-cancel" onClick={onClose}>
              {strings.aiSettingsCancel}
            </button>
            <button className="ai-settings-dialog-save" onClick={handleSave}>
              {strings.aiSettingsSave}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}

export type { IconSettings } from './icons'
