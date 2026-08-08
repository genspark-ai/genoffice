import { useEffect } from 'react'
import { htmlLang } from '@genoffice/i18n'
import { App as DocsApp } from '../../docs/src/renderer/App'
import { LocaleProvider, setModuleLang } from '../../docs/src/renderer/i18n/locale'
import { installAndroidDesktopApi } from './docs-platform'
import '../../docs/src/renderer/styles.css'
import '../../docs/src/renderer/fonts/fonts.css'

let installed = false

export function DocsEditorScreen(): React.JSX.Element {
  useEffect(() => {
    if (!installed) {
      installAndroidDesktopApi()
      installed = true
    }
    setModuleLang('en')
    document.documentElement.lang = htmlLang('en')
    document.body.classList.add('genoffice-android-editor')
    return () => document.body.classList.remove('genoffice-android-editor')
  }, [])

  return (
    <LocaleProvider initial="en">
      <div className="android-docs-host">
        <DocsApp />
      </div>
    </LocaleProvider>
  )
}
