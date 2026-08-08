import { useEffect } from 'react'
import { App as SheetsApp } from '../../sheets/src/renderer/App'
import { installRendererStub, aiDefaults } from './editor-stubs'
import '../../sheets/src/renderer/styles.css'

let installed = false

export function SheetsEditorScreen(): React.JSX.Element {
  useEffect(() => {
    if (!installed) {
      installRendererStub('desktopApi', {
        getAiSettings: async () => aiDefaults(),
        setAiSettings: async (settings: unknown) => localStorage.setItem('genoffice.android.ai.settings', JSON.stringify(settings)),
        aiGskStatus: async () => ({ loggedIn: false }),
        consumeNewBlankWorkbook: async () => true,
        hasQueuedWorkbook: async () => false,
        notifyPendingEdits: () => {},
        onWorkbookRenamed: () => () => {},
        onMenuAction: () => () => {},
        onCloseSaveRequest: () => () => {},
        onCloseCheck: () => () => {},
        reportCloseCheck: () => {},
      })
      installed = true
    }
    document.body.classList.add('genoffice-android-sheets')
    return () => document.body.classList.remove('genoffice-android-sheets')
  }, [])

  return <div className="android-sheets-host"><SheetsApp /></div>
}
