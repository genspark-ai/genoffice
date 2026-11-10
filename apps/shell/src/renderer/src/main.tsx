import './browser-mock'
import React from 'react'
import { createRoot } from 'react-dom/client'
import { htmlLang } from '@genoffice/i18n'
import { AppFrame } from './AppFrame'
import { LocaleProvider } from './locale'
import { ToastProvider } from './toast'
import { startAppearance } from './appearance'
import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import './home.css'
import './tabbar.css'
import './toast.css'
import { installScreenTips } from '@genoffice/ui'

installScreenTips()

// user appearance choices (accent, font, transparency) before first paint
startAppearance(document.documentElement)

// macOS shell window is created with vibrancy (Windows 11 with Mica, see
// windowMaterialActive below); a transparent body lets the sidebar, tab strip
// and the editor views' translucent regions show the material behind them
const platform = navigator.platform.toLowerCase()
if (platform.includes('mac')) document.body.classList.add('vib', 'plat-mac')
else if (platform.includes('win')) document.body.classList.add('plat-win')

// resolve the persisted language, first-run flag, and theme before first paint
// so the UI never flashes (home showing briefly before the onboarding overlay)
const langPromise =
  typeof window.aiOffice?.getLanguage === 'function'
    ? window.aiOffice.getLanguage().catch(() => 'en' as const)
    : Promise.resolve('en' as const)

const onboardingPromise =
  typeof window.aiOffice?.onboardingSeen === 'function'
    ? window.aiOffice.onboardingSeen().catch(() => true)
    : Promise.resolve(true)

// Windows 11: the main process created the window with Mica — go transparent
// before first paint, otherwise the body's in-app wallpaper would cover it
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
    createRoot(document.getElementById('root')!).render(
      <React.StrictMode>
        <LocaleProvider initial={lang}>
          <ToastProvider>
            <AppFrame initialOnboardingSeen={onboardingSeen} />
          </ToastProvider>
        </LocaleProvider>
      </React.StrictMode>,
    )
  },
)
