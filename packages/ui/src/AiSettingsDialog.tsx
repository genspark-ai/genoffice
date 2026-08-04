import { useEffect, useMemo, useState } from 'react'
import type {
  AiProviderConfig,
  AiProviderId,
  AiSettings,
  ModelListEntry,
} from '@genoffice/ai-provider'
import { AI_PROVIDERS } from '@genoffice/ai-provider'

/** Local OpenAI-compatible server presets for the "Custom / Local" provider. */
const LOCAL_PRESETS: { name: string; baseUrl: string }[] = [
  { name: 'Ollama', baseUrl: 'http://localhost:11434/v1' },
  { name: 'LM Studio', baseUrl: 'http://localhost:1234/v1' },
  { name: 'llama.cpp', baseUrl: 'http://localhost:8080/v1' },
  { name: 'vLLM', baseUrl: 'http://localhost:8000/v1' },
]

export interface AiTestResult {
  ok: boolean
  content?: string
  error?: string
}

export interface AiSettingsDialogProps {
  open: boolean
  onClose: () => void
  /** Load the current settings (usually window.<app>.getAiSettings()). */
  load: () => Promise<AiSettings>
  /** Persist new settings (usually window.<app>.setAiSettings()). */
  save: (settings: AiSettings) => Promise<void>
  /** Optional connection test against the pending provider (window.<app>.aiTestSettings). */
  test?: (settings: AiSettings) => Promise<AiTestResult>
  /** Optional live model catalog fetch (window.<app>.aiListModels); shows a "Fetch live models" button. */
  listModels?: (
    provider: AiProviderId,
    config: AiProviderConfig,
    freeOnly: boolean,
  ) => Promise<ModelListEntry[]>
  /** Called after a successful save so the app can refresh its live settings state. */
  onSaved?: (settings: AiSettings) => void
}

const PROVIDER_TIPS: Record<AiProviderId, string> = {
  genspark: 'Uses your Genspark account — no API key needed here.',
  anthropic: 'Bring your Anthropic API key (starts with sk-ant-).',
  gemini: 'Bring your Google AI Studio key (starts with AIza).',
  deepseek: 'Bring your DeepSeek API key (starts with sk-).',
  openai: 'Bring your OpenAI API key (starts with sk-).',
  openrouter:
    'One OpenRouter key (sk-or-v1-) unlocks Claude, GPT, Gemini, Llama and more with a single pay-per-token bill — pick a model below.',
  custom: 'Point GenOffice at any OpenAI-compatible endpoint — including a local model.',
}

const S = {
  overlay: {
    position: 'fixed' as const,
    inset: 0,
    background: 'rgba(10,14,22,0.55)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    zIndex: 9999,
    padding: 24,
    backdropFilter: 'blur(2px)',
  },
  card: {
    width: 640,
    maxWidth: '100%',
    maxHeight: '90vh',
    background: '#ffffff',
    color: '#1f2430',
    borderRadius: 14,
    boxShadow: '0 20px 60px rgba(0,0,0,0.35)',
    display: 'flex',
    flexDirection: 'column' as const,
    overflow: 'hidden',
    fontFamily:
      '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
  },
  header: {
    padding: '18px 22px',
    borderBottom: '1px solid #ececf1',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
  },
  title: { fontSize: 16, fontWeight: 600, margin: 0 },
  closeBtn: {
    border: 0,
    background: 'transparent',
    fontSize: 22,
    lineHeight: 1,
    cursor: 'pointer',
    color: '#6b7280',
    padding: '4px 8px',
    borderRadius: 8,
  },
  body: { display: 'flex', minHeight: 320, overflow: 'hidden' },
  nav: {
    width: 190,
    background: '#f7f7fa',
    borderRight: '1px solid #ececf1',
    padding: '10px 8px',
    overflowY: 'auto' as const,
    flexShrink: 0,
  },
  navBtn: (active: boolean) =>
    ({
      display: 'block',
      width: '100%',
      textAlign: 'left',
      padding: '9px 12px',
      marginBottom: 2,
      borderRadius: 8,
      border: 0,
      cursor: 'pointer',
      fontSize: 14,
      background: active ? '#0f6fff' : 'transparent',
      color: active ? '#fff' : '#3a3f4b',
      fontWeight: active ? 600 : 400,
    }) as const,
  navHint: {
    marginTop: 8,
    padding: '10px 12px',
    fontSize: 12,
    lineHeight: 1.5,
    color: '#6b7280',
  },
  form: { flex: 1, padding: '20px 22px', overflowY: 'auto' as const },
  label: {
    display: 'block',
    fontSize: 13,
    fontWeight: 600,
    color: '#3a3f4b',
    margin: '4px 0 6px',
  },
  input: {
    width: '100%',
    boxSizing: 'border-box' as const,
    padding: '9px 12px',
    fontSize: 14,
    borderRadius: 10,
    border: '1px solid #d4d7de',
    background: '#fff',
    color: '#1f2430',
    outline: 'none',
  },
  row: { marginBottom: 16 },
  tip: { fontSize: 12.5, color: '#6b7280', lineHeight: 1.55, margin: '6px 0 0' },
  presetRow: { display: 'flex', flexWrap: 'wrap' as const, gap: 8, margin: '8px 0 4px' },
  fetchRow: {
    display: 'flex',
    alignItems: 'center',
    gap: 14,
    marginTop: 10,
    flexWrap: 'wrap' as const,
  },
  freeLabel: {
    display: 'flex',
    alignItems: 'center',
    gap: 6,
    fontSize: 12.5,
    color: '#4b5563',
    cursor: 'pointer',
  },
  presetBtn: {
    border: '1px solid #d4d7de',
    background: '#f2f3f7',
    color: '#1f2430',
    borderRadius: 999,
    padding: '6px 12px',
    fontSize: 13,
    cursor: 'pointer',
  },
  footer: {
    padding: '14px 22px',
    borderTop: '1px solid #ececf1',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: 12,
    flexWrap: 'wrap' as const,
  },
  status: { fontSize: 12.5, color: '#6b7280', flex: 1, minWidth: 200 },
  statusOk: { fontSize: 12.5, color: '#11931c', flex: 1, minWidth: 200 },
  statusErr: { fontSize: 12.5, color: '#d3322b', flex: 1, minWidth: 200 },
  btn: {
    border: 0,
    borderRadius: 10,
    padding: '9px 18px',
    fontSize: 14,
    fontWeight: 600,
    cursor: 'pointer',
  },
  ghost: {
    border: '1px solid #d4d7de',
    background: '#fff',
    color: '#3a3f4b',
    borderRadius: 10,
    padding: '9px 16px',
    fontSize: 14,
    fontWeight: 600,
    cursor: 'pointer',
    display: 'inline-flex',
    alignItems: 'center',
    gap: 8,
  },
  primary: { background: '#0f6fff', color: '#fff' },
  right: { display: 'flex', gap: 10, alignItems: 'center' },
}

function Field({
  label,
  value,
  onChange,
  placeholder,
  type = 'text',
}: {
  label: string
  value: string
  onChange: (v: string) => void
  placeholder?: string
  type?: string
}) {
  return (
    <div style={S.row}>
      <label style={S.label}>{label}</label>
      <input
        style={S.input}
        type={type}
        value={value}
        placeholder={placeholder}
        onChange={(e) => onChange(e.target.value)}
        autoComplete="off"
        spellCheck={false}
      />
    </div>
  )
}

const CUSTOM_MODEL = '__custom_model__'

/** Model dropdown: live-fetched entries (with "(free)" labels) or the provider's curated list. */
function ModelPicker({
  options,
  value,
  placeholder,
  onChange,
}: {
  options: ModelListEntry[]
  value: string
  placeholder: string
  onChange: (v: string) => void
}) {
  const [custom, setCustom] = useState(
    options.length === 0 || (value !== '' && !options.some((o) => o.id === value)),
  )
  if (options.length === 0) {
    return <Field label="Model" value={value} onChange={onChange} placeholder={placeholder} />
  }
  return (
    <div style={S.row}>
      <label style={S.label}>Model</label>
      <select
        style={{ ...S.input, cursor: 'pointer' }}
        value={custom ? CUSTOM_MODEL : value}
        onChange={(e) => {
          if (e.target.value === CUSTOM_MODEL) {
            setCustom(true)
            onChange('')
          } else {
            setCustom(false)
            onChange(e.target.value)
          }
        }}
      >
        {options.map((o) => (
          <option key={o.id} value={o.id}>
            {o.label && o.label !== o.id ? `${o.id} — ${o.label}` : o.id}
            {o.free ? ' (free)' : ''}
          </option>
        ))}
        <option value={CUSTOM_MODEL}>Custom model…</option>
      </select>
      {custom ? (
        <input
          style={{ ...S.input, marginTop: 8 }}
          value={value}
          placeholder={placeholder}
          onChange={(e) => onChange(e.target.value)}
          autoComplete="off"
          spellCheck={false}
        />
      ) : null}
    </div>
  )
}

export function AiSettingsDialog({
  open,
  onClose,
  load,
  save,
  test,
  listModels,
  onSaved,
}: AiSettingsDialogProps) {
  const [draft, setDraft] = useState<AiSettings | null>(null)
  const [selected, setSelected] = useState<AiProviderId>('genspark')
  const [saving, setSaving] = useState(false)
  const [testing, setTesting] = useState(false)
  const [fetching, setFetching] = useState(false)
  const [freeOnly, setFreeOnly] = useState(false)
  const [fetchedModels, setFetchedModels] = useState<ModelListEntry[] | null>(null)
  const [status, setStatus] = useState<{ kind: 'idle' | 'ok' | 'err'; text: string }>({
    kind: 'idle',
    text: '',
  })

  useEffect(() => {
    if (!open) return
    let cancelled = false
    setStatus({ kind: 'idle', text: '' })
    load()
      .then((s) => {
        if (cancelled) return
        setDraft(s)
        setSelected(s.provider)
      })
      .catch(() => {
        if (cancelled) return
        setStatus({ kind: 'err', text: 'Could not load current AI settings.' })
      })
    return () => {
      cancelled = true
    }
  }, [open, load])

  // a provider change invalidates the previous catalog fetch
  useEffect(() => {
    setFetchedModels(null)
    setFreeOnly(false)
  }, [selected])

  const providerMeta = useMemo(
    () => AI_PROVIDERS.find((p) => p.id === selected),
    [selected],
  )

  if (!open) return null

  const cfg = draft?.providers[selected]
  const setCfg = (patch: Partial<AiProviderConfig>) => {
    if (!draft) return
    setDraft({
      ...draft,
      providers: { ...draft.providers, [selected]: { ...cfg!, ...patch } },
    })
  }

  const doSave = async (): Promise<void> => {
    if (!draft) return
    setSaving(true)
    try {
      await save({ ...draft, provider: selected })
      onSaved?.({ ...draft, provider: selected })
      setStatus({ kind: 'ok', text: 'Saved.' })
    } catch {
      setStatus({ kind: 'err', text: 'Save failed — check the fields and try again.' })
    } finally {
      setSaving(false)
    }
  }

  const doTest = async (): Promise<void> => {
    if (!draft || !test) return
    const pending: AiSettings = { ...draft, provider: selected }
    setTesting(true)
    setStatus({ kind: 'idle', text: 'Testing connection…' })
    try {
      const res = await test(pending)
      setStatus(
        res.ok
          ? { kind: 'ok', text: 'Connection OK.' }
          : { kind: 'err', text: res.error ?? 'Connection failed.' },
      )
    } catch (err) {
      setStatus({ kind: 'err', text: err instanceof Error ? err.message : 'Connection failed.' })
    } finally {
      setTesting(false)
    }
  }

  const doFetchModels = async (): Promise<void> => {
    if (!draft || !listModels || !cfg) return
    setFetching(true)
    setStatus({ kind: 'idle', text: 'Fetching model catalog…' })
    try {
      const models = await listModels(selected, cfg, freeOnly)
      setFetchedModels(models)
      setStatus(
        models.length === 0
          ? { kind: 'err', text: 'No models returned — check the API key.' }
          : {
              kind: 'ok',
              text: `${models.length} model${models.length === 1 ? '' : 's'} fetched${
                freeOnly ? ' (free only)' : ''
              }.`,
            },
      )
    } catch (err) {
      setStatus({ kind: 'err', text: err instanceof Error ? err.message : 'Failed to fetch models.' })
    } finally {
      setFetching(false)
    }
  }

  // live-fetched catalog wins; otherwise fall back to the provider's curated list
  const modelOptions: ModelListEntry[] = fetchedModels ?? (providerMeta?.models ?? []).map((id) => ({ id }))

  return (
    <div style={S.overlay} onClick={onClose}>
      <div style={S.card} onClick={(e) => e.stopPropagation()}>
        <div style={S.header}>
          <h2 style={S.title}>AI settings</h2>
          <button style={S.closeBtn} onClick={onClose} aria-label="Close">
            ×
          </button>
        </div>

        <div style={S.body}>
          <nav style={S.nav}>
            {AI_PROVIDERS.map((p) => (
              <button key={p.id} style={S.navBtn(selected === p.id)} onClick={() => setSelected(p.id)}>
                {p.label}
              </button>
            ))}
            <div style={S.navHint}>
              Choose a model provider for the AI panel. Your key stays on this device.
            </div>
          </nav>

          <div style={S.form}>
            <p style={S.tip}>{PROVIDER_TIPS[selected]}</p>

            {selected === 'genspark' ? (
              <div style={S.row}>
                <p style={S.tip}>
                  GenOffice uses your Genspark account and routes requests through the Genspark
                  service. No model key is stored locally.
                </p>
              </div>
            ) : null}

            {selected === 'custom' ? (
              <div>
                <div style={S.presetRow}>
                  {LOCAL_PRESETS.map((preset) => (
                    <button
                      key={preset.name}
                      style={S.presetBtn}
                      onClick={() => setCfg({ baseUrl: preset.baseUrl })}
                    >
                      {preset.name}
                    </button>
                  ))}
                </div>
                <Field
                  label="Base URL"
                  value={cfg?.baseUrl ?? ''}
                  onChange={(v) => setCfg({ baseUrl: v })}
                  placeholder="http://localhost:11434/v1"
                />
                <Field
                  label="Model"
                  value={cfg?.model ?? ''}
                  onChange={(v) => setCfg({ model: v })}
                  placeholder="e.g. qwen2.5:14b, llama3.1:8b"
                />
                <Field
                  label="API key (optional for local servers)"
                  type="password"
                  value={cfg?.apiKey ?? ''}
                  onChange={(v) => setCfg({ apiKey: v })}
                  placeholder="not-needed"
                />
                <p style={S.tip}>
                  The endpoint must accept POST /chat/completions and stream OpenAI-format tool
                  calls for the built-in agents.
                </p>
              </div>
            ) : null}

            {selected !== 'genspark' && selected !== 'custom' ? (
              <div>
                <Field
                  label="API key"
                  type="password"
                  value={cfg?.apiKey ?? ''}
                  onChange={(v) => setCfg({ apiKey: v })}
                  placeholder={providerMeta?.keyPlaceholder ?? 'sk-…'}
                />
                <ModelPicker
                  key={fetchedModels ? 'live' : 'static'}
                  options={modelOptions}
                  value={cfg?.model ?? ''}
                  placeholder={providerMeta?.defaultModel ?? ''}
                  onChange={(v) => setCfg({ model: v })}
                />
                {listModels && cfg ? (
                  <div style={S.fetchRow}>
                    <button style={S.ghost} onClick={doFetchModels} disabled={fetching || !cfg.apiKey}>
                      {fetching ? 'Fetching…' : 'Fetch live models'}
                    </button>
                    {selected === 'openrouter' ? (
                      <label style={S.freeLabel}>
                        <input
                          type="checkbox"
                          checked={freeOnly}
                          onChange={(e) => setFreeOnly(e.target.checked)}
                        />
                        Free models only
                      </label>
                    ) : null}
                    {!cfg.apiKey ? (
                      <span style={S.status}>Enter an API key to fetch the catalog.</span>
                    ) : null}
                  </div>
                ) : null}
              </div>
            ) : null}
          </div>
        </div>

        <div style={S.footer}>
          <div
            style={
              status.kind === 'ok' ? S.statusOk : status.kind === 'err' ? S.statusErr : S.status
            }
          >
            {status.text}
          </div>
          <div style={S.right}>
            {test ? (
              <button style={S.ghost} onClick={doTest} disabled={testing || saving}>
                {testing ? 'Testing…' : 'Test connection'}
              </button>
            ) : null}
            <button style={S.btn} onClick={onClose} disabled={saving}>
              Cancel
            </button>
            <button style={{ ...S.btn, ...S.primary }} onClick={doSave} disabled={saving}>
              {saving ? 'Saving…' : 'Save'}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}

/** Small gear icon button that opens AI settings. */
export function AiSettingsButton({
  onClick,
  title = 'AI settings',
}: {
  onClick: () => void
  title?: string
}) {
  return (
    <button
      onClick={onClick}
      title={title}
      aria-label={title}
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'center',
        width: 30,
        height: 30,
        border: 0,
        borderRadius: 8,
        background: 'transparent',
        cursor: 'pointer',
        color: 'inherit',
      }}
    >
      <svg
        width="17"
        height="17"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      >
        <circle cx="12" cy="12" r="3" />
        <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 1 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" />
      </svg>
    </button>
  )
}

