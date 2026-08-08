import React from 'react'
import { htmlLang } from '@genoffice/i18n'
import { App as DocsApp } from '../../docs/src/renderer/App'
import { LocaleProvider, setModuleLang } from '../../docs/src/renderer/i18n/locale'
import { installAndroidDesktopApi } from './docs-platform'
import '../../docs/src/renderer/styles.css'
import '../../docs/src/renderer/fonts/fonts.css'
import './docs-mobile.css'

// Docs reads window.desktop during its initial render/effects. The Android bridge
// must therefore exist BEFORE <DocsApp /> mounts; installing it in useEffect caused
// a first-render race that could leave the Android app on a white screen.
installAndroidDesktopApi()

export function DocsEditorScreen(): React.JSX.Element {
  setModuleLang('en')
  document.documentElement.lang = htmlLang('en')
  document.body.classList.add('genoffice-android-editor')

  return (
    <LocaleProvider initial="en">
      <div className="android-docs-host">
        <DocsApp />
      </div>
    </LocaleProvider>
  )
}
