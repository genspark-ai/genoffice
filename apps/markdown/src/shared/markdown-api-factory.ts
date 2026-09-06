/// Transport-agnostic construction of the markdown renderer bridge APIs.
///
/// One source of truth for the `window.markdownApi` / `window.projectApi`
/// surface: the sandboxed preload binds it to ipcRenderer (Electron), the
/// browser web-bridge binds the exact same object to the HTTP/SSE transport
/// (web version). Channel names, argument shapes and listener wrappers are
/// identical to what the preload has always exposed — only the transport
/// differs.
import type { Lang } from '@genoffice/i18n'
import type { AiStreamChunk } from '@genoffice/ai-provider'
import type { ProjectApi } from '@genoffice/project-store'
import type { IpcTransport } from '@genoffice/ipc-bridge/client'
import { AI_CHANNELS, MARKDOWN_CHANNELS } from './ipc'
import type { ExportFormat, MarkdownApi, SaveMode, UiTheme } from './ipc'

export interface MarkdownApiOverrides {
  /** Web-native image picker (browser file input + saveImage). */
  pickImage?: () => Promise<string | null>
  /** Web-native DOCX export (browser download of the serialized bytes). */
  exportDocx?: (request: { base64: string; suggestedName?: string; mode?: string }) => Promise<{
    ok: boolean
    path?: string
    canceled?: boolean
    error?: string
  }>
  /** Web-native PDF export (browser print of the print HTML). */
  exportPdf?: (request: { html: string; suggestedName?: string }) => Promise<{
    ok: boolean
    path?: string
    canceled?: boolean
    error?: string
  }>
}

export function createMarkdownApi(
  t: IpcTransport,
  overrides: MarkdownApiOverrides = {},
): MarkdownApi {
  const api: MarkdownApi = {
    consumePending: () => t.invoke(MARKDOWN_CHANNELS.consumePending),
    readFile: (path) => t.invoke(MARKDOWN_CHANNELS.readFile, path),
    save: (request) => t.invoke(MARKDOWN_CHANNELS.save, request),
    setDirty: (dirty) => t.send(MARKDOWN_CHANNELS.dirtyChanged, dirty),
    onSaveRequest: (handler) =>
      t.on(MARKDOWN_CHANNELS.saveRequest, (mode) => handler(mode as SaveMode)),
    onCloseSaveRequest: (handler) => t.on(MARKDOWN_CHANNELS.closeSaveRequest, () => handler()),
    sendCloseSaveResult: (ok) => t.send(MARKDOWN_CHANNELS.closeSaveResult, ok),
    sendSaveRequestAck: (ok) => t.send(MARKDOWN_CHANNELS.saveRequestAck, ok),
    onFileRenamed: (handler) =>
      t.on(MARKDOWN_CHANNELS.fileRenamed, (newPath) => handler(newPath as string)),
    pickImage: overrides.pickImage ?? (() => t.invoke(MARKDOWN_CHANNELS.pickImage)),
    saveImage: (data) => t.invoke(MARKDOWN_CHANNELS.saveImage, data),
    readImage: (src) => t.invoke(MARKDOWN_CHANNELS.readImage, src),
    onExportRequest: (handler) =>
      t.on(MARKDOWN_CHANNELS.exportRequest, (format) => handler(format as ExportFormat)),
    onPrintRequest: (handler) => t.on(MARKDOWN_CHANNELS.printRequest, () => handler()),
    exportDocx: overrides.exportDocx ?? ((request) => t.invoke(MARKDOWN_CHANNELS.exportDocx, request)),
    exportPdf: overrides.exportPdf ?? ((request) => t.invoke(MARKDOWN_CHANNELS.exportPdf, request)),
    getLanguage: () => t.invoke(MARKDOWN_CHANNELS.getLanguage),
    onLanguageChanged: (handler) =>
      t.on(MARKDOWN_CHANNELS.languageChanged, (lang) => handler(lang as Lang)),
    getTheme: () => t.invoke(MARKDOWN_CHANNELS.getTheme),
    onThemeChanged: (handler) =>
      t.on(MARKDOWN_CHANNELS.themeChanged, (theme) => handler(theme as UiTheme)),
    onChromePressed: (handler) => t.on('app:chrome-pressed', () => handler()),
    getAiSettings: () => t.invoke(AI_CHANNELS.getSettings),
    aiStream: (request) => t.invoke(AI_CHANNELS.stream, request),
    aiStreamCancel: (requestId) => t.invoke(AI_CHANNELS.streamCancel, requestId),
    onAiStream: (handler) =>
      t.on(AI_CHANNELS.streamChunk, (chunk) => handler(chunk as AiStreamChunk)),
    webSearch: (query, maxResults) => t.invoke(AI_CHANNELS.webSearch, query, maxResults),
    imageSearch: (query, maxResults) => t.invoke(AI_CHANNELS.imageSearch, query, maxResults),
    fetchImage: (url) => t.invoke(AI_CHANNELS.fetchImage, url),
    aiGenerateImage: (op) => t.invoke(MARKDOWN_CHANNELS.aiGenerateImage, op),
  }
  return api
}

/** Chat persistence: the shared project:* handlers are registered once by the shell (docs-main registerProjectIpc) */
export function createMarkdownProjectApi(
  t: IpcTransport,
): Pick<ProjectApi, 'resolveChat' | 'appendChat' | 'loadChat' | 'rebindChat'> {
  return {
    resolveChat: (args) => t.invoke('project:resolveChat', args),
    appendChat: (args) => t.invoke('project:appendChat', args),
    loadChat: (args) => t.invoke('project:loadChat', args),
    rebindChat: (args) => t.invoke('project:rebindChat', args),
  }
}
