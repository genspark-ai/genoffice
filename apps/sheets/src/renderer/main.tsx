import ReactDOM from 'react-dom/client'
import { htmlLang, type Lang } from '@genoffice/i18n'
import { installScreenTips } from '@genoffice/ui'

import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@univerjs/preset-sheets-core/lib/index.css'

import { App } from './App'
import { LocaleProvider, setModuleLang } from './i18n/locale'
import type { UiTheme } from '../shared/desktop-api'
import './styles.css'

if (import.meta.hot) {
  import.meta.hot.on('vite:beforeUpdate', ({ updates }) => {
    const replacesUniverRuntime = updates.some(
      ({ path }) => path.endsWith('/App.tsx') || path.endsWith('/univer-sync.ts'),
    )
    if (replacesUniverRuntime) window.location.reload()
  })
}

const root = document.getElementById('root')
if (!root) throw new Error('Missing application root.')

installScreenTips()

function applyTheme(theme: UiTheme): void {
  if (theme === 'system') document.documentElement.removeAttribute('data-theme')
  else document.documentElement.setAttribute('data-theme', theme)
}

async function bootstrap(): Promise<void> {
  let lang: Lang = 'zh'
  let theme: UiTheme = 'system'
  try {
    // per-promise catch: standalone runs have no app:get-theme handler, and
    // that rejection must not drop a resolved language
    ;[lang, theme] = await Promise.all([
      window.desktopApi.getLanguage().catch(() => 'zh' as const),
      window.desktopApi.getTheme().catch(() => 'system' as const),
    ])
  } catch {
    /* dev renderer without the preload bridge */
  }
  setModuleLang(lang)
  document.documentElement.lang = htmlLang(lang)
  applyTheme(theme)
  window.desktopApi?.onThemeChanged(applyTheme)
  ReactDOM.createRoot(root!).render(
    <LocaleProvider initial={lang}>
      <App />
    </LocaleProvider>,
  )
}

void bootstrap()
