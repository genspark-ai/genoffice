import './browser-mock'
import React from 'react'
import { createRoot } from 'react-dom/client'
import { htmlDir, htmlLang } from '@genoffice/i18n'
import { AppFrame } from './AppFrame'
import { LocaleProvider } from './locale'
import { ToastProvider } from './toast'
import { startAppearance } from './appearance'
import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@genoffice/ui/dropdown.css'
import '@genoffice/ui/ai-model-picker.css'
import '@genoffice/ui/ai-usage.css'
import '@genoffice/ui/ai-skills.css'
import './home.css'
import './tabbar.css'
import './toast.css'
import { installScreenTips } from '@genoffice/ui'

installScreenTips()

// user appearance choices (accent, font, transparency) before first paint
startAppearance(document.documentElement)

const IS_MAC = navigator.platform.toLowerCase().includes('mac')
if (IS_MAC) document.body.classList.add('vib', 'plat-mac')
else if (navigator.platform.toLowerCase().includes('win')) document.body.classList.add('plat-win')
document.body.classList.add(IS_MAC ? 'mac' : 'overlay-title-bar')

const langPromise =
  typeof window.aiOffice?.getLanguage === 'function'
    ? window.aiOffice.getLanguage().catch(() => 'en' as const)
    : Promise.resolve('en' as const)

const onboardingPromise =
  typeof window.aiOffice?.onboardingSeen === 'function'
    ? window.aiOffice.onboardingSeen().catch(() => true)
    : Promise.resolve(true)

const materialPromise =
  typeof window.aiOffice?.windowMaterialActive === 'function'
    ? window.aiOffice.windowMaterialActive().catch(() => false)
    : Promise.resolve(false)

const themePromise =
  typeof window.aiOffice?.getTheme === 'function'
    ? window.aiOffice.getTheme().catch(() => 'system' as const)
    : Promise.resolve('system' as const)

void Promise.all([langPromise, onboardingPromise, themePromise, materialPromise]).then(
  ([lang, onboardingSeen, theme, material]) => {
    if (material) document.body.classList.add('vib')
    document.documentElement.lang = htmlLang(lang)
    document.documentElement.dir = htmlDir(lang)
    // apply theme attribute before first paint to avoid flash
    if (theme !== 'system') {
      document.documentElement.setAttribute('data-theme', theme)
    }
    if (typeof window.aiOffice?.onThemeChanged === 'function') {
      window.aiOffice.onThemeChanged((next) => {
        if (next === 'system') document.documentElement.removeAttribute('data-theme')
        else document.documentElement.setAttribute('data-theme', next)
      })
    }

    const isHelp = new URLSearchParams(location.search).get('mode') === 'help'
    const HelpScreen = React.lazy(() =>
      import('./i18n/help/HelpScreen').then((m) => ({ default: m.HelpScreen })),
    )

    createRoot(document.getElementById('root')!).render(
      <React.StrictMode>
        <LocaleProvider initial={lang}>
          <ToastProvider>
            {isHelp ? (
              <React.Suspense fallback={<div className="boot" />}>
                <HelpScreen />
              </React.Suspense>
            ) : (
              <AppFrame initialOnboardingSeen={onboardingSeen} />
            )}
          </ToastProvider>
        </LocaleProvider>
      </React.StrictMode>,
    )
  },
)
