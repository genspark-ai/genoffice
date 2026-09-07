import { createRoot } from 'react-dom/client'
import { htmlLang, type Lang } from '@genoffice/i18n'
import App from './App'
import type { UiTheme } from '../shared/ipc'
import '@genoffice/ui/tokens.css'
import './styles.css'

function applyTheme(theme: UiTheme): void {
  if (theme === 'system') document.documentElement.removeAttribute('data-theme')
  else document.documentElement.setAttribute('data-theme', theme)
}

void (async () => {
  const [lang, theme] = await Promise.all([
    window.hwpApi.getLanguage().catch(() => 'en' as const),
    window.hwpApi.getTheme().catch(() => 'system' as const),
  ])
  document.documentElement.lang = htmlLang(lang as Lang)
  applyTheme(theme)
  window.hwpApi.onLanguageChanged((next) => {
    document.documentElement.lang = htmlLang(next)
  })
  window.hwpApi.onThemeChanged(applyTheme)
  createRoot(document.getElementById('root')!).render(<App />)
})()
