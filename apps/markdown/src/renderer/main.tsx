import { createRoot } from 'react-dom/client'
import { htmlLang, type Lang } from '@genoffice/i18n'
import App from './App'
import { LocaleProvider } from './i18n/locale'
import './styles.css'

void (async () => {
  const lang: Lang = await window.markdownApi.getLanguage().catch(() => 'zh' as const)
  document.documentElement.lang = htmlLang(lang)
  createRoot(document.getElementById('root')!).render(
    <LocaleProvider initial={lang}>
      <App />
    </LocaleProvider>,
  )
})()
