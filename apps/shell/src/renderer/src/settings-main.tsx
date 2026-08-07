import React from 'react'
import { createRoot } from 'react-dom/client'
import { htmlLang, normalizeLang } from '@genoffice/i18n'
import { LocaleProvider } from './locale'
import { SettingsApp } from './Settings'
import './settings.css'

// The settings window is opened from the shell, which already resolved the
// language; read it from the snapshot so the window never flashes English
// before switching.
void window.genofficeSettings
  .load()
  .then((snapshot) => {
    const lang = normalizeLang(snapshot.language)
    document.documentElement.lang = htmlLang(lang)
    createRoot(document.getElementById('root')!).render(
      <React.StrictMode>
        <LocaleProvider initial={lang}>
          <SettingsApp initial={snapshot} />
        </LocaleProvider>
      </React.StrictMode>,
    )
  })
  .catch((err: unknown) => {
    // a settings window that cannot load is more useful showing why than blank
    const root = document.getElementById('root')
    if (root) root.textContent = `Failed to load settings: ${String(err)}`
  })
