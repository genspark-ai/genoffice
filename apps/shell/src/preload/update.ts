import { contextBridge, ipcRenderer } from 'electron'
import type { IpcRendererEvent } from 'electron'
import type { UpdateUiState, UpdateWindowApi } from '../shared/update-api'
import { UPDATE_CHANNELS } from '../shared/update-api'

const api: UpdateWindowApi = {
  async getState() {
    const result: unknown = await ipcRenderer.invoke(UPDATE_CHANNELS.getState)
    return (result ?? null) as UpdateUiState | null
  },
  download() {
    void ipcRenderer.invoke(UPDATE_CHANNELS.download)
  },
  install() {
    void ipcRenderer.invoke(UPDATE_CHANNELS.install)
  },
  later() {
    void ipcRenderer.invoke(UPDATE_CHANNELS.later)
  },
  openDownload() {
    void ipcRenderer.invoke(UPDATE_CHANNELS.openDownload)
  },
  onState(handler) {
    const listener = (_event: IpcRendererEvent, state: UpdateUiState) => handler(state)
    ipcRenderer.on(UPDATE_CHANNELS.changed, listener)
    return () => ipcRenderer.removeListener(UPDATE_CHANNELS.changed, listener)
  },
}

contextBridge.exposeInMainWorld('aiOfficeUpdate', api)
