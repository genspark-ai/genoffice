import React, { useEffect, useMemo, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { App as CapacitorApp } from '@capacitor/app'
import { Keyboard } from '@capacitor/keyboard'
import { StatusBar, Style } from '@capacitor/status-bar'
import { Capacitor } from '@capacitor/core'
import { AI_PROVIDERS, chatForProvider, defaultAiSettings } from '@genoffice/ai-provider'
import type { AiProviderId, AiSettings } from '@genoffice/ai-provider'
import { DocsEditorScreen } from './docs-editor'
import './styles.css'

type Screen = 'home' | 'docs' | 'sheets' | 'slides' | 'pdf' | 'ai'
const SETTINGS_KEY = 'genoffice.android.ai.settings'

function loadSettings(): AiSettings {
  const defaults = defaultAiSettings()
  try {
    const raw = localStorage.getItem(SETTINGS_KEY)
    if (!raw) return defaults
    const saved = JSON.parse(raw) as Partial<AiSettings>
    return {
      provider: saved.provider ?? defaults.provider,
      providers: { ...defaults.providers, ...(saved.providers ?? {}) },
    }
  } catch {
    return defaults
  }
}

function AiPanel() {
  const [settings, setSettings] = useState<AiSettings>(() => loadSettings())
  const [message, setMessage] = useState('Write a short professional leave letter.')
  const [answer, setAnswer] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')

  const meta = useMemo(() => AI_PROVIDERS.find((p) => p.id === settings.provider) ?? AI_PROVIDERS[0], [settings.provider])
  const config = settings.providers[settings.provider]

  const persist = (next: AiSettings) => {
    setSettings(next)
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(next))
  }

  const setProvider = (provider: AiProviderId) => {
    const nextMeta = AI_PROVIDERS.find((p) => p.id === provider)!
    const providers = { ...settings.providers }
    if (!providers[provider]?.model) providers[provider] = { apiKey: '', model: nextMeta.defaultModel, baseUrl: '' }
    persist({ ...settings, provider, providers })
  }

  const updateConfig = (patch: Partial<typeof config>) => {
    persist({ ...settings, providers: { ...settings.providers, [settings.provider]: { ...config, ...patch } } })
  }

  const send = async () => {
    setBusy(true)
    setError('')
    setAnswer('')
    try {
      if (!config.apiKey) throw new Error('Enter an API key first.')
      const result = await chatForProvider(settings.provider, config, 'You are a helpful office assistant. Return only the requested result.', message)
      if (!result.ok) throw new Error(result.error ?? 'AI request failed')
      setAnswer(result.content ?? '')
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="ai-panel">
      <h1>AI Assistant</h1>
      <p className="muted">Shared GenOffice provider layer, also available inside Docs.</p>
      <label>Provider</label>
      <select value={settings.provider} onChange={(e) => setProvider(e.target.value as AiProviderId)}>
        {AI_PROVIDERS.filter((p) => p.id !== 'genspark').map((p) => <option key={p.id} value={p.id}>{p.label}</option>)}
      </select>
      <label>Model</label>
      {meta.models.length ? (
        <select value={config.model} onChange={(e) => updateConfig({ model: e.target.value })}>
          {meta.models.map((model) => <option key={model} value={model}>{model}</option>)}
        </select>
      ) : <input value={config.model} onChange={(e) => updateConfig({ model: e.target.value })} placeholder="Model name" />}
      {meta.needsBaseUrl && <><label>Base URL</label><input value={config.baseUrl ?? ''} onChange={(e) => updateConfig({ baseUrl: e.target.value })} placeholder="https://your-provider.example/v1" /></>}
      <label>API key</label>
      <input type="password" value={config.apiKey} onChange={(e) => updateConfig({ apiKey: e.target.value })} placeholder={meta.keyPlaceholder} autoComplete="off" />
      <label>Request</label>
      <textarea value={message} onChange={(e) => setMessage(e.target.value)} rows={5} />
      <button className="primary" disabled={busy} onClick={() => void send()}>{busy ? 'Thinking…' : 'Send to AI'}</button>
      {error && <div className="notice warning"><strong>Request failed</strong><span>{error}</span></div>}
      {answer && <div className="answer"><strong>Response</strong><pre>{answer}</pre></div>}
    </div>
  )
}

function MobileShell() {
  const [active, setActive] = useState<Screen>('home')

  useEffect(() => {
    void StatusBar.setStyle({ style: Style.Light }).catch(() => {})
    void Keyboard.setAccessoryBarVisible({ isVisible: false }).catch(() => {})
    const back = CapacitorApp.addListener('backButton', ({ canGoBack }) => {
      if (active !== 'home') {
        setActive('home')
        return
      }
      if (canGoBack) window.history.back()
    })
    return () => { void back.then((handle) => handle.remove()) }
  }, [active])

  const open = (screen: Screen) => setActive(screen)

  return (
    <main className={`mobile-shell ${active === 'docs' ? 'editor-mode' : ''}`}>
      {active !== 'docs' && <header className="topbar">
        <div><div className="brand">GenOffice</div><div className="subtitle">Android</div></div>
        <button className="settings" aria-label="AI settings" onClick={() => open('ai')}>⚙</button>
      </header>}

      {active === 'docs' ? <DocsEditorScreen /> : <section className="content">
        {active === 'home' ? (
          <>
            <h1>Office, wherever you are.</h1>
            <p className="muted">The Android app now runs the real shared Docs editor instead of a placeholder shell.</p>
            <div className="cards">
              <button className="app-card" onClick={() => open('docs')}><span className="icon">DOCX</span><span><strong>Docs</strong><small>Open, edit and save Word documents</small></span><span className="arrow">›</span></button>
              {([['sheets', 'Sheets', 'XLSX'], ['slides', 'Slides', 'PPTX'], ['pdf', 'PDF', 'PDF']] as const).map(([id, title, ext]) => <button key={id} className="app-card" onClick={() => open(id)}><span className="icon">{ext}</span><span><strong>{title}</strong><small>Android editor adapter</small></span><span className="arrow">›</span></button>)}
            </div>
            <button className="ai-card" onClick={() => open('ai')}><strong>AI Assistant</strong><span>OpenRouter, NVIDIA Nemotron and other compatible providers</span><b>Open →</b></button>
          </>
        ) : active === 'ai' ? <AiPanel /> : (
          <>
            <button className="back" onClick={() => open('home')}>‹ Back</button>
            <h1>{active[0].toUpperCase() + active.slice(1)}</h1>
            <p className="muted">The same platform-adapter architecture is reserved for Sheets, Slides and PDF.</p>
            <div className="notice"><strong>Next editor</strong><span>This screen is intentionally kept separate from Docs so the working Android Docs port remains stable while the other engines are adapted.</span></div>
          </>
        )}
      </section>}

      {active !== 'docs' && <nav className="bottom-nav">
        <button className={active === 'home' ? 'active' : ''} onClick={() => open('home')}>Home</button>
        <button className={active === 'docs' ? 'active' : ''} onClick={() => open('docs')}>Docs</button>
        <button className={active === 'sheets' ? 'active' : ''} onClick={() => open('sheets')}>Sheets</button>
        <button className={active === 'slides' ? 'active' : ''} onClick={() => open('slides')}>Slides</button>
        <button className={active === 'ai' ? 'active' : ''} onClick={() => open('ai')}>AI</button>
      </nav>}
    </main>
  )
}

createRoot(document.getElementById('root')!).render(<React.StrictMode><MobileShell /></React.StrictMode>)
