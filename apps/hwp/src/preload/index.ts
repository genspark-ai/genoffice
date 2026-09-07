import { contextBridge, ipcRenderer } from 'electron'
import type { Lang } from '@genoffice/i18n'
import { installDropOpenBridge } from '@genoffice/electron-utils/drop-open'
import { HWP_CHANNELS } from '../shared/ipc'
import type { HwpApi, SaveMode, UiTheme } from '../shared/ipc'

const api: HwpApi = {
  consumePending: () => ipcRenderer.invoke(HWP_CHANNELS.consumePending),
  readFile: (path) => ipcRenderer.invoke(HWP_CHANNELS.readFile, path),
  save: (request) => ipcRenderer.invoke(HWP_CHANNELS.save, request),
  setDirty: (dirty) => ipcRenderer.send(HWP_CHANNELS.dirtyChanged, dirty),
  onSaveRequest: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, mode: SaveMode) => handler(mode)
    ipcRenderer.on(HWP_CHANNELS.saveRequest, listener)
    return () => ipcRenderer.removeListener(HWP_CHANNELS.saveRequest, listener)
  },
  onCloseSaveRequest: (handler) => {
    const listener = () => handler()
    ipcRenderer.on(HWP_CHANNELS.closeSaveRequest, listener)
    return () => ipcRenderer.removeListener(HWP_CHANNELS.closeSaveRequest, listener)
  },
  sendCloseSaveResult: (ok) => ipcRenderer.send(HWP_CHANNELS.closeSaveResult, ok),
  sendSaveRequestAck: (ok) => ipcRenderer.send(HWP_CHANNELS.saveRequestAck, ok),
  onFileRenamed: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, newPath: string) => handler(newPath)
    ipcRenderer.on(HWP_CHANNELS.fileRenamed, listener)
    return () => ipcRenderer.removeListener(HWP_CHANNELS.fileRenamed, listener)
  },
  getLanguage: () => ipcRenderer.invoke(HWP_CHANNELS.getLanguage),
  onLanguageChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, lang: Lang) => handler(lang)
    ipcRenderer.on(HWP_CHANNELS.languageChanged, listener)
    return () => ipcRenderer.removeListener(HWP_CHANNELS.languageChanged, listener)
  },
  getTheme: () => ipcRenderer.invoke(HWP_CHANNELS.getTheme),
  onThemeChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, theme: UiTheme) => handler(theme)
    ipcRenderer.on(HWP_CHANNELS.themeChanged, listener)
    return () => ipcRenderer.removeListener(HWP_CHANNELS.themeChanged, listener)
  },
}

contextBridge.exposeInMainWorld('hwpApi', api)

installDropOpenBridge()
