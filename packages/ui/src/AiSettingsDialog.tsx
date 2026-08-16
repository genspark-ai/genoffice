import { useState } from 'react'
import type { AiSettings, AiProviderId } from '@genoffice/ai-provider'
import { AI_PROVIDERS } from '@genoffice/ai-provider'

export interface AiSettingsDialogProps {
  /** current persisted AI settings */
  settings: AiSettings
  onSave(next: AiSettings): void
  onClose(): void
  /** optional per-key label overrides (defaults to English) */
  labels?: {
    title?: string
    provider?: string
    apiKey?: string
    model?: string
    baseUrl?: string
    save?: string
    cancel?: string
    gensparkNote?: string
  }
}

const DIALOG_OVERLAY_STYLE: React.CSSProperties = {
  position: 'fixed',
  inset: 0,
  zIndex: 9999,
  background: 'rgba(0,0,0,0.35)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
}

const DIALOG_STYLE: React.CSSProperties = {
  background: '#fff',
  color: '#1f2328',
  borderRadius: 10,
  boxShadow: '0 12px 40px rgba(0,0,0,0.25)',
  width: 420,
  maxWidth: 'calc(100vw - 48px)',
  maxHeight: 'calc(100vh - 96px)',
  overflow: 'auto',
  padding: 20,
  fontFamily:
    "'Segoe UI', -apple-system, BlinkMacSystemFont, 'PingFang SC', 'Microsoft YaHei', sans-serif",
  fontSize: 13,
  lineHeight: 1.5,
}

const LABEL_STYLE: React.CSSProperties = {
  display: 'block',
  margin: '10px 0 4px',
  fontWeight: 600,
}

const INPUT_STYLE: React.CSSProperties = {
  width: '100%',
  boxSizing: 'border-box',
  padding: '6px 8px',
  border: '1px solid #d0d7de',
  borderRadius: 6,
  fontSize: 13,
  background: '#fff',
  color: '#1f2328',
}

const BUTTON_ROW_STYLE: React.CSSProperties = {
  display: 'flex',
  justifyContent: 'flex-end',
  gap: 8,
  marginTop: 18,
}

const BUTTON_STYLE: React.CSSProperties = {
  padding: '6px 14px',
  borderRadius: 6,
  border: '1px solid #d0d7de',
  background: '#f6f8fa',
  color: '#1f2328',
  fontSize: 13,
  cursor: 'pointer',
}

const PRIMARY_BUTTON_STYLE: React.CSSProperties = {
  ...BUTTON_STYLE,
  background: '#1f6feb',
  borderColor: '#1f6feb',
  color: '#fff',
}

const HINT_STYLE: React.CSSProperties = {
  margin: '6px 0 0',
  color: '#57606a',
  fontSize: 12,
}

const NOTE_STYLE: React.CSSProperties = {
  marginTop: 12,
  padding: '8px 10px',
  background: '#fff8c5',
  border: '1px solid #eed888',
  borderRadius: 6,
  fontSize: 12,
  color: '#5b4b00',
}

/**
 * Modal for choosing an LLM provider (API key / model / optional base URL).
 * Reads and writes @genoffice/ai-provider's AiSettings shape directly.
 * Styling is inline so no per-app stylesheet changes are needed.
 */
export function AiSettingsDialog({ settings, onSave, onClose, labels }: AiSettingsDialogProps) {
  const L = {
    title: 'AI Settings',
    provider: 'Provider',
    apiKey: 'API Key',
    model: 'Model',
    baseUrl: 'Base URL (OpenAI-compatible)',
    save: 'Save',
    cancel: 'Cancel',
    gensparkNote:
      'Genspark uses your signed-in Genspark account — no API key needed. It routes Claude / GPT / Gemini through the Genspark service.',
    ...labels,
  }

  const [provider, setProvider] = useState<AiProviderId>(
    AI_PROVIDERS.some((p) => p.id === settings.provider) ? settings.provider : 'genspark',
  )
  const [apiKey, setApiKey] = useState(settings.providers?.[provider]?.apiKey ?? '')
  const [model, setModel] = useState(settings.providers?.[provider]?.model ?? '')
  const [baseUrl, setBaseUrl] = useState(settings.providers?.[provider]?.baseUrl ?? '')

  const meta = AI_PROVIDERS.find((p) => p.id === provider) ?? AI_PROVIDERS[0]!

  const changeProvider = (id: string) => {
    const pid = id as AiProviderId
    setProvider(pid)
    setApiKey(settings.providers?.[pid]?.apiKey ?? '')
    setModel(settings.providers?.[pid]?.model ?? '')
    setBaseUrl(settings.providers?.[pid]?.baseUrl ?? '')
  }

  const save = () => {
    const providers = { ...settings.providers }
    providers[provider] = {
      apiKey,
      model: model.trim(),
      ...(provider === 'custom' ? { baseUrl: baseUrl.trim() } : {}),
    }
    onSave({ provider, providers })
  }

  return (
    <div
      style={DIALOG_OVERLAY_STYLE}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div style={DIALOG_STYLE} role="dialog" aria-label={L.title}>
        <h3 style={{ margin: 0, fontSize: 15 }}>{L.title}</h3>

        <label style={LABEL_STYLE}>{L.provider}</label>
        <select
          style={INPUT_STYLE}
          value={provider}
          onChange={(e) => changeProvider(e.target.value)}
        >
          {AI_PROVIDERS.map((p) => (
            <option key={p.id} value={p.id}>
              {p.label}
            </option>
          ))}
        </select>

        {provider === 'genspark' ? (
          <p style={NOTE_STYLE}>{L.gensparkNote}</p>
        ) : (
          <>
            <label style={LABEL_STYLE}>{L.apiKey}</label>
            <input
              style={INPUT_STYLE}
              type="password"
              value={apiKey}
              onChange={(e) => setApiKey(e.target.value)}
              placeholder={provider === 'custom' ? 'sk-... (or any token)' : 'sk-...'}
              autoComplete="off"
            />
          </>
        )}

        <label style={LABEL_STYLE}>{L.model}</label>
        {meta.models.length > 0 ? (
          <select
            style={INPUT_STYLE}
            value={model}
            onChange={(e) => setModel(e.target.value)}
            disabled={provider === 'genspark'}
          >
            {model && !meta.models.includes(model) && <option value={model}>{model}</option>}
            {meta.models.map((m) => (
              <option key={m} value={m}>
                {m}
              </option>
            ))}
          </select>
        ) : (
          <input
            style={INPUT_STYLE}
            value={model}
            onChange={(e) => setModel(e.target.value)}
            placeholder="e.g. qwen2.5:14b, gpt-4.1-mini"
          />
        )}

        {provider === 'custom' && (
          <>
            <label style={LABEL_STYLE}>{L.baseUrl}</label>
            <input
              style={INPUT_STYLE}
              value={baseUrl}
              onChange={(e) => setBaseUrl(e.target.value)}
              placeholder="http://localhost:11434/v1"
            />
            <p style={HINT_STYLE}>Must be an OpenAI-compatible endpoint ending in /v1.</p>
          </>
        )}

        <div style={BUTTON_ROW_STYLE}>
          <button style={BUTTON_STYLE} onClick={onClose}>
            {L.cancel}
          </button>
          <button style={PRIMARY_BUTTON_STYLE} onClick={save}>
            {L.save}
          </button>
        </div>
      </div>
    </div>
  )
}
