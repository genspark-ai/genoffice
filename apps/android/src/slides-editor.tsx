import { useEffect } from 'react'
import { App as SlidesApp } from '../../slides/src/renderer/App'
import { installRendererStub, aiDefaults } from './editor-stubs'
import '../../slides/src/renderer/styles.css'

let installed = false

export function SlidesEditorScreen(): React.JSX.Element {
  useEffect(() => {
    if (!installed) {
      installRendererStub('slidesApi', {
        getAiSettings: async () => aiDefaults(),
        setAiSettings: async (settings: unknown) => localStorage.setItem('genoffice.android.ai.settings', JSON.stringify(settings)),
        aiGskStatus: async () => ({ loggedIn: false }),
        consumePendingOpen: async () => null,
        consumeNewPresentation: async () => true,
        setAutoSavePref: () => {},
        onMenuAction: () => () => {},
        onCloseCheck: () => () => {},
        reportCloseCheck: () => {},
        onCloseSaveRequest: () => () => {},
        reportCloseSaveResult: () => {},
      })
      installRendererStub('desktopApi', {
        getAiSettings: async () => aiDefaults(),
        aiGskStatus: async () => ({ loggedIn: false }),
      })
      installed = true
    }
    document.body.classList.add('genoffice-android-slides')
    return () => document.body.classList.remove('genoffice-android-slides')
  }, [])

  return <div className="android-slides-host"><SlidesApp /></div>
}
