/// Transport-agnostic construction of the pdf renderer bridge API.
///
/// One source of truth for the `window.pdfApi` / `window.projectApi` surface:
/// the sandboxed preload builds it on an ipcRenderer transport (Electron), the
/// browser web-bridge builds the exact same object on the HTTP/SSE transport
/// (web version). Channel names, argument shapes, listener wrappers and return
/// coercion are identical — only the transport differs.
import type { Lang } from '@genoffice/i18n'
import type { AiStreamChunk } from '@genoffice/ai-provider'
import type { IpcTransport } from '@genoffice/ipc-bridge/client'
import { AI_CHANNELS, PDF_CHANNELS } from './ipc'
import type { PdfApi, UiTheme } from './ipc'

export interface PdfApiOverrides {
  /** Web-native open: grant a path to the bridge sender and return it as pending. */
  consumePending?: () => Promise<string | null>
}

export function createPdfApi(t: IpcTransport, overrides: PdfApiOverrides = {}): PdfApi {
  return {
    consumePending: overrides.consumePending ?? (() => t.invoke(PDF_CHANNELS.consumePending)),
    readFile: (path) => t.invoke(PDF_CHANNELS.readFile, path),
    save: (request) => t.invoke(PDF_CHANNELS.save, request),
    autoRename: (path, baseName) => t.invoke(PDF_CHANNELS.autoRename, path, baseName),
    isUntitled: (path) => t.invoke(PDF_CHANNELS.isUntitled, path),
    validateTextEdits: (request) => t.invoke(PDF_CHANNELS.validateTextEdits, request),
    listEditFonts: () => t.invoke(PDF_CHANNELS.listEditFonts),
    canDrawText: (text, font, bold, italic) =>
      t.invoke(PDF_CHANNELS.canDrawText, text, font, bold, italic),
    listPageImages: (path) => t.invoke(PDF_CHANNELS.listPageImages, path),
    listStaticFormFills: (path) => t.invoke(PDF_CHANNELS.listStaticFormFills, path),
    ocrPage: (png) => t.invoke(PDF_CHANNELS.ocrPage, png),
    pageImagePng: (request) => t.invoke(PDF_CHANNELS.pageImagePng, request),
    pagePreviewPng: (request) => t.invoke(PDF_CHANNELS.pagePreviewPng, request),
    extractPages: (request) => t.invoke(PDF_CHANNELS.extractPages, request),
    insertPdf: (request) => t.invoke(PDF_CHANNELS.insertPdf, request),
    insertBlankPage: (request) => t.invoke(PDF_CHANNELS.insertBlankPage, request),
    splitPdf: (request) => t.invoke(PDF_CHANNELS.splitPdf, request),
    mergePdf: (request) => t.invoke(PDF_CHANNELS.mergePdf, request),
    mergePages: (request) => t.invoke(PDF_CHANNELS.mergePages, request),
    replacePages: (request) => t.invoke(PDF_CHANNELS.replacePages, request),
    setPageSize: (request) => t.invoke(PDF_CHANNELS.setPageSize, request),
    splitPages: (request) => t.invoke(PDF_CHANNELS.splitPages, request),
    cropPages: (request) => t.invoke(PDF_CHANNELS.cropPages, request),
    exportImages: (request) => t.invoke(PDF_CHANNELS.exportImages, request),
    convertOffice: (format) => t.invoke(PDF_CHANNELS.convertOffice, format),
    createDocument: (request) => t.invoke(PDF_CHANNELS.createDocument, request),
    imageSearch: (query, maxResults) =>
      t.invoke(AI_CHANNELS.imageSearch, query, maxResults),
    fetchImage: (url) => t.invoke(AI_CHANNELS.fetchImage, url),
    generateImage: (op) => t.invoke(PDF_CHANNELS.generateImage, op),
    listSavedSignatures: () => t.invoke(PDF_CHANNELS.listSignatures),
    addSavedSignature: (data) => t.invoke(PDF_CHANNELS.addSignature, data),
    removeSavedSignature: (id) => t.invoke(PDF_CHANNELS.removeSignature, id),
    getUsername: () => t.invoke(PDF_CHANNELS.getUsername),
    setDirty: (dirty) => t.send(PDF_CHANNELS.dirtyChanged, dirty),
    onCloseSaveRequest: (handler) => t.on(PDF_CHANNELS.closeSaveRequest, () => handler()),
    sendCloseSaveResult: (ok) => t.send(PDF_CHANNELS.closeSaveResult, ok),
    onSaveAsRequest: (handler) =>
      t.on(PDF_CHANNELS.saveAsRequest, (targetPath) => handler(targetPath as string)),
    sendSaveAsResult: (ok) => t.send(PDF_CHANNELS.saveAsResult, ok),
    onSaveAsFlow: (handler) =>
      t.on(PDF_CHANNELS.saveAsFlow, (inFlight) => handler(inFlight as boolean)),
    onPrintRequest: (handler) => t.on(PDF_CHANNELS.printRequest, () => handler()),
    getLanguage: () => t.invoke(PDF_CHANNELS.getLanguage),
    onLanguageChanged: (handler) =>
      t.on(PDF_CHANNELS.languageChanged, (lang) => handler(lang as Lang)),
    getTheme: () => t.invoke(PDF_CHANNELS.getTheme),
    onThemeChanged: (handler) =>
      t.on(PDF_CHANNELS.themeChanged, (theme) => handler(theme as UiTheme)),
    onChromePressed: (handler) => t.on('app:chrome-pressed', () => handler()),
    getAiSettings: () => t.invoke(AI_CHANNELS.getSettings),
    gskStatus: () => t.invoke(AI_CHANNELS.gskStatus),
    aiStream: (request) => t.invoke(AI_CHANNELS.stream, request),
    aiStreamCancel: (requestId) => t.invoke(AI_CHANNELS.streamCancel, requestId),
    onAiStream: (handler) =>
      t.on(AI_CHANNELS.streamChunk, (chunk) => handler(chunk as AiStreamChunk)),
  }
}

export interface ProjectApi {
  resolveChat: (args: { filePath: string | null; tempChatId?: string }) => Promise<unknown>
  appendChat: (args: unknown) => Promise<unknown>
  loadChat: (args: { projectId: string; chatId: string; limit?: number }) => Promise<unknown>
  rebindChat: (args: { projectId: string; tempChatId: string; newFilePath: string }) => Promise<unknown>
}

export function createPdfProjectApi(t: IpcTransport): ProjectApi {
  return {
    resolveChat: (args) => t.invoke('project:resolveChat', args),
    appendChat: (args) => t.invoke('project:appendChat', args),
    loadChat: (args) => t.invoke('project:loadChat', args),
    rebindChat: (args) => t.invoke('project:rebindChat', args),
  }
}
