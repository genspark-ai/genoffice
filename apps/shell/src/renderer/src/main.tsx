import React from 'react'
import { createRoot } from 'react-dom/client'
import { htmlLang } from '@genoffice/i18n'
import { AppFrame } from './AppFrame'
import { LocaleProvider } from './locale'
import '@genoffice/ui/tokens.css'
import '@genoffice/ui/screentip.css'
import '@genoffice/ui/dropdown.css'
import './home.css'
import './tabbar.css'
import { installScreenTips } from '@genoffice/ui'

installScreenTips()

// macOS shell window is created with vibrancy; a transparent body lets the
// editor views' translucent regions (e.g. slides thumbnail pane) show it
const IS_MAC = navigator.platform.toLowerCase().includes('mac')
if (IS_MAC) document.body.classList.add('vib')
// non-mac: the tab strip doubles as the title bar (caption buttons overlay it)
document.body.classList.add(IS_MAC ? 'mac' : 'overlay-title-bar')

// resolve the persisted language, first-run flag, and theme before first paint
// so the UI never flashes (home showing briefly before the onboarding overlay)
void Promise.all([
  window.aiOffice.getLanguage(),
  // if the flag is unreadable, skip onboarding rather than block the home screen
  window.aiOffice.onboardingSeen().catch(() => true),
  window.aiOffice.getTheme().catch(() => 'system' as const),
]).then(([lang, onboardingSeen, theme]) => {
  document.documentElement.lang = htmlLang(lang)
  // apply theme attribute before first paint to avoid flash
  if (theme !== 'system') {
    document.documentElement.setAttribute('data-theme', theme)
  }
  window.aiOffice.onThemeChanged((next) => {
    if (next === 'system') document.documentElement.removeAttribute('data-theme')
    else document.documentElement.setAttribute('data-theme', next)
  })
  // The manual tab is the same bundle branched on ?mode=help: no onboarding,
  // no home state, just the searchable topic browser (issue #1520).
  //
  // Split, though: this entry chunk is parsed on every launch whatever the tab
  // is, and the screen plus its stylesheet are only ever wanted on the other
  // one. A static import meant Home's first paint paid for a manual nobody had
  // opened.
  const isHelp = new URLSearchParams(location.search).get('mode') === 'help'
  const HelpScreen = React.lazy(() =>
    import('./i18n/help/HelpScreen').then((m) => ({ default: m.HelpScreen })),
  )
  createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <LocaleProvider initial={lang}>
        {isHelp ? (
          <React.Suspense fallback={<div className="boot" />}>
            <HelpScreen />
          </React.Suspense>
        ) : (
          <AppFrame initialOnboardingSeen={onboardingSeen} />
        )}
      </LocaleProvider>
    </React.StrictMode>,
  )
})
