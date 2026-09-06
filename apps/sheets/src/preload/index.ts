import { contextBridge, ipcRenderer, webUtils } from 'electron'
import { installDropOpenBridge } from '@genoffice/electron-utils/drop-open'
import { createElectronIpcTransport } from '@genoffice/ipc-bridge/client'
import { createSheetsApi, createSheetsProjectApi } from '../shared/sheets-api-factory'

const transport = createElectronIpcTransport(ipcRenderer)
const overrides = { getPathForFile: (file: File) => webUtils.getPathForFile(file) }
contextBridge.exposeInMainWorld('desktopApi', createSheetsApi(transport, overrides))
contextBridge.exposeInMainWorld('projectApi', createSheetsProjectApi(transport))

// Off by default. e2e drivers launch the BUILT app with GENOFFICE_DEBUG_HOOKS=1
// so the renderer exposes window.__genofficeDebug (see App.tsx) — the dev-only
// __univerAPI hook does not exist in production bundles.
if (process.env.GENOFFICE_DEBUG_HOOKS === '1') {
  contextBridge.exposeInMainWorld('__genofficeDebugHooks', true)
}

// open documents dragged from the OS onto this tab as a new shell tab
installDropOpenBridge()
