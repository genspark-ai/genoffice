import React, { useEffect, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { App as CapacitorApp } from '@capacitor/app'
import { Keyboard } from '@capacitor/keyboard'
import { StatusBar, Style } from '@capacitor/status-bar'
import { Capacitor } from '@capacitor/core'
import './styles.css'

function MobileShell() {
  const [active, setActive] = useState<'home' | 'docs' | 'sheets' | 'slides' | 'pdf'>('home')

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

    return () => {
      void back.then((handle) => handle.remove())
    }
  }, [active])

  return (
    <main className="mobile-shell">
      <header className="topbar">
        <div>
          <div className="brand">GenOffice</div>
          <div className="subtitle">Android</div>
        </div>
        <button className="settings" aria-label="Settings">⚙</button>
      </header>

      <section className="content">
        {active === 'home' ? (
          <>
            <h1>Office, wherever you are.</h1>
            <p className="muted">The Android editor layer is being connected to the shared GenOffice engines.</p>
            <div className="cards">
              {([
                ['docs', 'Docs', 'DOCX'],
                ['sheets', 'Sheets', 'XLSX'],
                ['slides', 'Slides', 'PPTX'],
                ['pdf', 'PDF', 'PDF'],
              ] as const).map(([id, title, ext]) => (
                <button key={id} className="app-card" onClick={() => setActive(id)}>
                  <span className="icon">{ext}</span>
                  <span><strong>{title}</strong><small>Open and edit</small></span>
                  <span className="arrow">›</span>
                </button>
              ))}
            </div>
            <div className="notice">
              <strong>AI providers</strong>
              <span>OpenRouter and direct NVIDIA NIM will use the same shared provider layer as desktop.</span>
            </div>
          </>
        ) : (
          <>
            <button className="back" onClick={() => setActive('home')}>‹ Back</button>
            <h1>{active[0].toUpperCase() + active.slice(1)}</h1>
            <p className="muted">Editor integration is the next adapter stage. No desktop Electron APIs are used by this Android shell.</p>
            <div className="notice warning">
              <strong>Android adapter</strong>
              <span>This screen is intentionally not pretending to be a finished editor. The existing desktop editor depends on Electron IPC and must be adapted before it is embedded here.</span>
            </div>
          </>
        )}
      </section>

      <nav className="bottom-nav">
        <button className={active === 'home' ? 'active' : ''} onClick={() => setActive('home')}>Home</button>
        <button onClick={() => setActive('docs')}>Docs</button>
        <button onClick={() => setActive('sheets')}>Sheets</button>
        <button onClick={() => setActive('slides')}>Slides</button>
        <button onClick={() => setActive('pdf')}>PDF</button>
      </nav>
    </main>
  )
}

createRoot(document.getElementById('root')!).render(
  <React.StrictMode><MobileShell /></React.StrictMode>,
)
