import { contextBridge, ipcRenderer } from 'electron'
import type { IpcRendererEvent } from 'electron'
import type { PdfPasswordUiState, PdfPasswordWindowApi } from '../shared/pdf-password-api'
import { PDF_PASSWORD_CHANNELS } from '../shared/pdf-password-api'

const api: PdfPasswordWindowApi = {
  async getState() {
    const result: unknown = await ipcRenderer.invoke(PDF_PASSWORD_CHANNELS.getState)
    return (result ?? null) as PdfPasswordUiState | null
  },
  submit(password) {
    void ipcRenderer.invoke(PDF_PASSWORD_CHANNELS.submit, password)
  },
  cancel() {
    void ipcRenderer.invoke(PDF_PASSWORD_CHANNELS.cancel)
  },
  onState(handler) {
    const listener = (_event: IpcRendererEvent, state: PdfPasswordUiState) => handler(state)
    ipcRenderer.on(PDF_PASSWORD_CHANNELS.changed, listener)
    return () => ipcRenderer.removeListener(PDF_PASSWORD_CHANNELS.changed, listener)
  },
}

contextBridge.exposeInMainWorld('aiOfficePdfPassword', api)
