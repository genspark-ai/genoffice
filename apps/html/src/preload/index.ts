import type { AiPanelPrefs } from '@genoffice/ui'
import { contextBridge, ipcRenderer, webUtils } from 'electron'
import type { Lang } from '@genoffice/i18n'
import type { AiStreamChunk } from '@genoffice/ai-provider'
import type { ProjectApi } from '@genoffice/project-store'
import { installDropOpenBridge } from '@genoffice/electron-utils/drop-open'
import { AI_CHANNELS, HTML_CHANNELS } from '../shared/ipc'
import type { AutoSaveDefault, ExportFormat, HtmlApi, SaveMode, UiTheme } from '../shared/ipc'

const api: HtmlApi = {
  consumePending: () => ipcRenderer.invoke(HTML_CHANNELS.consumePending),
  readFile: (path) => ipcRenderer.invoke(HTML_CHANNELS.readFile, path),
  updatePreview: (text) => ipcRenderer.send(HTML_CHANNELS.previewUpdate, text),
  getPreviewInfo: () => ipcRenderer.invoke(HTML_CHANNELS.previewInfo),
  setPresentFullScreen: (on) => ipcRenderer.invoke(HTML_CHANNELS.presentFullScreen, on),
  presentInNewTab: (title) => ipcRenderer.invoke(HTML_CHANNELS.presentNewTab, title),
  save: (request) => ipcRenderer.invoke(HTML_CHANNELS.save, request),
  setDirty: (dirty) => ipcRenderer.send(HTML_CHANNELS.dirtyChanged, dirty),
  onSaveRequest: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, mode: SaveMode) => handler(mode)
    ipcRenderer.on(HTML_CHANNELS.saveRequest, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.saveRequest, listener)
  },
  onCloseSaveRequest: (handler) => {
    const listener = () => handler()
    ipcRenderer.on(HTML_CHANNELS.closeSaveRequest, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.closeSaveRequest, listener)
  },
  sendCloseSaveResult: (ok) => ipcRenderer.send(HTML_CHANNELS.closeSaveResult, ok),
  sendSaveRequestAck: (ok) => ipcRenderer.send(HTML_CHANNELS.saveRequestAck, ok),
  onFileRenamed: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, newPath: string) => handler(newPath)
    ipcRenderer.on(HTML_CHANNELS.fileRenamed, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.fileRenamed, listener)
  },
  setProvisionalTitle: (title) => ipcRenderer.send(HTML_CHANNELS.provisionalTitle, title),
  pickImage: () => ipcRenderer.invoke(HTML_CHANNELS.pickImage),
  saveImage: (data) => ipcRenderer.invoke(HTML_CHANNELS.saveImage, data),
  readImage: (src) => ipcRenderer.invoke(HTML_CHANNELS.readImage, src),
  pickAttachments: () => ipcRenderer.invoke(HTML_CHANNELS.filesPick),
  addAttachmentPaths: (paths) => ipcRenderer.invoke(HTML_CHANNELS.filesAdd, paths),
  addPastedImage: (data, ext) => ipcRenderer.invoke(HTML_CHANNELS.filesAddPastedImage, data, ext),
  readAttachment: (path, offset, maxChars) =>
    ipcRenderer.invoke(HTML_CHANNELS.filesRead, path, offset, maxChars),
  readAttachmentImage: (path) => ipcRenderer.invoke(HTML_CHANNELS.filesReadImage, path),
  getPathForFile: (file) => webUtils.getPathForFile(file),
  onExportRequest: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, format: ExportFormat) => handler(format)
    ipcRenderer.on(HTML_CHANNELS.exportRequest, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.exportRequest, listener)
  },
  onPrintRequest: (handler) => {
    const listener = () => handler()
    ipcRenderer.on(HTML_CHANNELS.printRequest, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.printRequest, listener)
  },
  exportDocx: (request) => ipcRenderer.invoke(HTML_CHANNELS.exportDocx, request),
  exportPdf: (request) => ipcRenderer.invoke(HTML_CHANNELS.exportPdf, request),
  getLanguage: () => ipcRenderer.invoke(HTML_CHANNELS.getLanguage),
  onLanguageChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, lang: Lang) => handler(lang)
    ipcRenderer.on(HTML_CHANNELS.languageChanged, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.languageChanged, listener)
  },
  getTheme: () => ipcRenderer.invoke(HTML_CHANNELS.getTheme),
  onThemeChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, theme: UiTheme) => handler(theme)
    ipcRenderer.on(HTML_CHANNELS.themeChanged, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.themeChanged, listener)
  },
  getAutoSaveDefault: () => ipcRenderer.invoke(HTML_CHANNELS.getAutoSaveDefault),
  onAutoSaveDefaultChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, value: AutoSaveDefault) => handler(value)
    ipcRenderer.on(HTML_CHANNELS.autoSaveDefaultChanged, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.autoSaveDefaultChanged, listener)
  },
  getAiPanelPrefs: () => ipcRenderer.invoke(HTML_CHANNELS.getAiPanelPrefs),
  onAiPanelPrefsChanged: (handler) => {
    const listener = (_event: Electron.IpcRendererEvent, prefs: AiPanelPrefs) => handler(prefs)
    ipcRenderer.on(HTML_CHANNELS.aiPanelPrefsChanged, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.aiPanelPrefsChanged, listener)
  },
  onChromePressed: (handler) => {
    const listener = () => handler()
    ipcRenderer.on('app:chrome-pressed', listener)
    return () => ipcRenderer.removeListener('app:chrome-pressed', listener)
  },
  getAiSettings: () => ipcRenderer.invoke(AI_CHANNELS.getSettings),
  aiGskStatus: () => ipcRenderer.invoke(AI_CHANNELS.gskStatus),
  aiStream: (request) => ipcRenderer.invoke(AI_CHANNELS.stream, request),
  aiStreamCancel: (requestId) => ipcRenderer.invoke(AI_CHANNELS.streamCancel, requestId),
  onAiStream: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, chunk: AiStreamChunk) => handler(chunk)
    ipcRenderer.on(AI_CHANNELS.streamChunk, listener)
    return () => ipcRenderer.removeListener(AI_CHANNELS.streamChunk, listener)
  },
  webSearch: (query, maxResults) => ipcRenderer.invoke(AI_CHANNELS.webSearch, query, maxResults),
  imageSearch: (query, maxResults) =>
    ipcRenderer.invoke(AI_CHANNELS.imageSearch, query, maxResults),
  fetchImage: (url) => ipcRenderer.invoke(HTML_CHANNELS.fetchImage, url),
  aiGenerateImage: (op) => ipcRenderer.invoke(HTML_CHANNELS.aiGenerateImage, op),
}

/** Chat persistence: the shared project:* handlers are registered once by the shell (docs-main registerProjectIpc) */
const projectApi: Pick<ProjectApi, 'resolveChat' | 'appendChat' | 'loadChat' | 'rebindChat'> = {
  resolveChat: (args) => ipcRenderer.invoke('project:resolveChat', args),
  appendChat: (args) => ipcRenderer.invoke('project:appendChat', args),
  loadChat: (args) => ipcRenderer.invoke('project:loadChat', args),
  rebindChat: (args) => ipcRenderer.invoke('project:rebindChat', args),
}

contextBridge.exposeInMainWorld('htmlApi', api)
contextBridge.exposeInMainWorld('projectApi', projectApi)

// open documents dragged from the OS onto this tab as a new shell tab
installDropOpenBridge()
