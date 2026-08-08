import { useEffect } from 'react'
import { App as SheetsApp } from '../../sheets/src/renderer/App'
import { installRendererStub, aiDefaults } from './editor-stubs'
import '../../sheets/src/renderer/styles.css'

// Install the Android compatibility bridge BEFORE SheetsApp mounts.
// The desktop renderer can read window.desktop during initial render/effects;
// installing it in useEffect creates a first-render race that can result in a
// completely blank WebView.
let installed = false

function installSheetsBridge(): void {
  if (installed) return
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

installSheetsBridge()

document.body.classList.add('genoffice-android-sheets')

export function SheetsEditorScreen(): React.JSX.Element {
  useEffect(() => () => document.body.classList.remove('genoffice-android-sheets'), [])
  return <div className="android-sheets-host"><SheetsApp /></div>
}
