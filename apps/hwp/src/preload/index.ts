import { contextBridge, ipcRenderer } from 'electron'
import type { Lang } from '@genoffice/i18n'
import { installDropOpenBridge } from '@genoffice/electron-utils/drop-open'
import { HWP_CHANNELS } from '../shared/ipc'
import type { HwpApi, UiTheme } from '../shared/ipc'

const api: HwpApi = {
  consumePending: () => ipcRenderer.invoke(HWP_CHANNELS.consumePending),
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
  onChromePressed: (handler) => {
    const listener = () => handler()
    ipcRenderer.on('app:chrome-pressed', listener)
    return () => ipcRenderer.removeListener('app:chrome-pressed', listener)
  },
}

contextBridge.exposeInMainWorld('hwpApi', api)

installDropOpenBridge()
