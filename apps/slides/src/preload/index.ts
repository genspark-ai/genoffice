import { contextBridge, ipcRenderer, webUtils } from 'electron'
import { installDropOpenBridge } from '@genoffice/electron-utils/drop-open'
import { createElectronIpcTransport } from '@genoffice/ipc-bridge/client'
import {
  createSlidesApi,
  createSlidesFilesApi,
  createSlidesProjectApi,
} from '../shared/slides-api-factory'

const transport = createElectronIpcTransport(ipcRenderer)
const overrides = { getPathForFile: (file: File) => webUtils.getPathForFile(file) }
contextBridge.exposeInMainWorld('slidesApi', createSlidesApi(transport))
contextBridge.exposeInMainWorld('desktop', createSlidesFilesApi(transport, overrides))
contextBridge.exposeInMainWorld('projectApi', createSlidesProjectApi(transport))

// open documents dragged from the OS onto this tab as a new shell tab
installDropOpenBridge()
