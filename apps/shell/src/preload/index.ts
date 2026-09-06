import { contextBridge, ipcRenderer } from 'electron'
import { installDropOpenBridge } from '@genoffice/electron-utils/drop-open'
import { createElectronIpcTransport } from '@genoffice/ipc-bridge/client'
import {
  createShellHomeApi,
  createShellProjectApi,
  createShellTabsApi,
} from '../shared/shell-api-factory'

const transport = createElectronIpcTransport(ipcRenderer)
contextBridge.exposeInMainWorld('aiOffice', createShellHomeApi(transport))
contextBridge.exposeInMainWorld('aiOfficeProject', createShellProjectApi(transport))
contextBridge.exposeInMainWorld('aiOfficeTabs', createShellTabsApi(transport))

// open documents dragged from the OS anywhere over Home or the tab strip
installDropOpenBridge()
