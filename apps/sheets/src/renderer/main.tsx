import ReactDOM from 'react-dom/client'
import { htmlLang, type Lang } from '@genoffice/i18n'
import { installScreenTips } from '@genoffice/ui'

import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@genoffice/ui/color-picker.css'
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

// Canvas fillText never triggers @font-face downloads, so the bundled Carlito
// faces (Calibri/Aptos aliases in styles.css) must be loaded before Univer's
// first skeleton — MDW, wrap points, and #### overflow all measure with them.
async function loadCellFonts(): Promise<void> {
  const loads: Promise<unknown>[] = []
  for (const variant of ['', 'bold ', 'italic ', 'italic bold ']) {
    for (const family of ['Calibri', 'Aptos']) {
      loads.push(document.fonts?.load?.(`${variant}16px ${family}`)?.catch(() => {}) ?? [])
    }
  }
  // Local assets resolve in milliseconds; the timeout only guards a broken
  // bundle from blanking the app.
  await Promise.race([Promise.all(loads), new Promise((resolve) => setTimeout(resolve, 3000))])
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
  await loadCellFonts()
  window.desktopApi?.onThemeChanged(applyTheme)
  ReactDOM.createRoot(root!).render(
    <LocaleProvider initial={lang}>
      <App />
    </LocaleProvider>,
  )
}

void bootstrap()
