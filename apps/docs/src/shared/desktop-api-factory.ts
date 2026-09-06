/// Transport-agnostic construction of the docs renderer bridge APIs.
///
/// One source of truth for the `window.desktop` / `window.projectApi` surface:
/// the sandboxed preload builds it on an ipcRenderer transport (Electron), the
/// browser web-bridge builds the exact same object on the HTTP/SSE transport
/// (web version). Channel names, argument shapes, listener wrappers and return
/// coercion are identical to what the preload has always exposed — only the
/// transport differs, which is what makes both versions behave the same.
import type {
  AiChatRequest,
  AiSettings,
  AiStreamChunk,
  AiStreamRequest,
  DesktopApi,
  MenuCommand,
  UiTheme,
} from './ipc'
import type { ProjectApi } from '@genoffice/project-store'
import type { IpcTransport } from '@genoffice/ipc-bridge/client'

export interface DesktopApiOverrides {
  /**
   * The preload supplies `webUtils.getPathForFile` (Electron-only). The web
   * bridge has no OS file access, so it omits the override and dropped-file
   * resolution surfaces a desktop-only error instead.
   */
  getPathForFile?: (file: File) => string
  /** Web-native open (browser file picker → temp file → open-path). */
  openDocx?: () => Promise<unknown>
  /** Web-native image picker (browser file input → base64 result). */
  pickImage?: () => Promise<unknown>
  /** Web-native Save As (browser download of the serialized bytes). */
  saveDocxAs?: (
    defaultName: string,
    data: ArrayBuffer,
    sourcePath?: string | null,
  ) => Promise<{ ok: boolean; path?: string; error?: string; passwordIntentPending?: boolean }>
  /** Web-native first save (browser download of the serialized bytes). */
  saveDocxNew?: (
    defaultName: string,
    data: ArrayBuffer,
  ) => Promise<{ ok: boolean; path?: string; error?: string; passwordIntentPending?: boolean }>
  /** Web-native print (browser print dialog). */
  print?: () => Promise<{ ok: boolean; error?: string }>
  /** Web-native PDF export (browser print dialog → save as PDF). */
  exportPdf?: (
    defaultName: string,
    pageWidthTwips: number,
    pageHeightTwips: number,
    outPath?: string,
  ) => Promise<{ ok: boolean; path?: string; error?: string }>
  /** Web-native clipboard write. */
  copyImageToClipboard?: (dataUrl: string, metaJson?: string) => Promise<boolean>
  /** Web-native font metrics (canvas measureText). */
  fontMetrics?: (family: string) => Promise<unknown>
  /** Web-native attachment picker (browser file input → temp files → files:add). */
  pickAttachments?: () => Promise<unknown>
  /** Web-native tab management (browser tabs). */
  openNewTab?: (openPath?: string | null) => Promise<void>
  listDocsTabs?: () => Promise<unknown>
  focusDocsTab?: (id: string) => Promise<void>
  /** Web-native mixed-size export: browser print dialog (no PDF bytes in the browser). */
  printPdfBuffer?: (
    pageWidthTwips: number,
    pageHeightTwips: number,
  ) => Promise<{ ok: boolean; base64?: string; error?: string }>
  /** Web-native merged-PDF save: browser print dialog (save as PDF). */
  saveMergedPdf?: (
    defaultName: string,
    base64Parts: string[],
    outPath?: string,
  ) => Promise<{ ok: boolean; path?: string; error?: string }>
}

export function createDesktopApi(t: IpcTransport, overrides: DesktopApiOverrides = {}): DesktopApi {
  const api: DesktopApi = {
    getLanguage: () => t.invoke('app:get-language'),
    onLanguageChanged: (handler) =>
      t.on('app:language-changed', (lang) => handler(lang as Parameters<typeof handler>[0])),
    getTheme: () => t.invoke('app:get-theme'),
    onThemeChanged: (handler) => t.on('app:theme-changed', (theme) => handler(theme as UiTheme)),
    onChromePressed: (handler) => t.on('app:chrome-pressed', () => handler()),
    openDocx: overrides.openDocx ?? (() => t.invoke('docs:open')),
    openDocxPath: (path: string) => t.invoke('docs:open-path', path),
    openDocxDecrypt: (path: string, password: string) =>
      t.invoke('docs:open-decrypt', path, password),
    setDocPassword: (filePath: string | null, password: string | null) =>
      t.invoke('docs:set-password', filePath, password),
    docPasswordIntentRevision: async () => {
      const revision: unknown = await t.invoke('docs:password-intent-revision')
      return typeof revision === 'number' && Number.isSafeInteger(revision) && revision >= 0
        ? revision
        : 0
    },
    discardDocPasswordIntents: (throughRevision: number) =>
      t.invoke('docs:discard-password-intents', throughRevision),
    consumePendingOpenDocx: () => t.invoke('docs:consume-pending-open'),
    consumeNewBlankDoc: () => t.invoke('docs:consume-new-blank'),
    consumeAiDocContent: () => t.invoke('docs:consume-ai-doc-content'),
    createDocument: (request) => t.invoke('docs:create-document', request),
    onOpenDocx: (handler) =>
      t.on('docs:opened', (result) => handler(result as Parameters<typeof handler>[0])),
    onRenamedDocx: (handler) =>
      t.on('docs:renamed', (paths) => handler(paths as Parameters<typeof handler>[0])),
    saveDocx: (path: string, data: ArrayBuffer, auto?: boolean) =>
      t.invoke('docs:save', path, data, auto === true),
    writeRecoveryCopy: (path: string, data: ArrayBuffer) =>
      t.invoke('docs:write-recovery', path, data),
    onTeardown: (handler) => t.on('docs:teardown', () => handler()),
    saveDocxAs: overrides.saveDocxAs ??
      ((defaultName: string, data: ArrayBuffer, sourcePath?: string | null) =>
        t.invoke('docs:save-as', defaultName, data, sourcePath ?? null)),
    saveDocxNew: overrides.saveDocxNew ??
      ((defaultName: string, data: ArrayBuffer) => t.invoke('docs:save-new', defaultName, data)),
    getRecentFiles: () => t.invoke('docs:recent'),
    pickImage: overrides.pickImage ?? (() => t.invoke('docs:pick-image')),
    fontMetrics: overrides.fontMetrics ?? ((family: string) => t.invoke('docs:font-metrics', family)),
    print: overrides.print ?? (() => t.invoke('docs:print')),
    exportPdf: overrides.exportPdf ??
      ((
        defaultName: string,
        pageWidthTwips: number,
        pageHeightTwips: number,
        outPath?: string,
      ) => t.invoke('docs:export-pdf', defaultName, pageWidthTwips, pageHeightTwips, outPath)),
    printPdfBuffer: overrides.printPdfBuffer ??
      ((pageWidthTwips: number, pageHeightTwips: number) =>
        t.invoke('docs:print-pdf-buffer', pageWidthTwips, pageHeightTwips)),
    saveMergedPdf: overrides.saveMergedPdf ??
      ((defaultName: string, base64Parts: string[], outPath?: string) =>
        t.invoke('docs:save-merged-pdf', defaultName, base64Parts, outPath)),
    getAiSettings: () => t.invoke('ai:get-settings'),
    setAiSettings: (settings: AiSettings) => t.invoke('ai:set-settings', settings),
    aiChat: (request: AiChatRequest) => t.invoke('ai:chat', request),
    aiStream: (request: AiStreamRequest) => t.invoke('ai:stream', request),
    aiStreamCancel: (requestId: string) => t.invoke('ai:stream-cancel', requestId),
    aiGskStatus: (withEmail?: boolean) => t.invoke('ai:gsk-status', withEmail),
    aiGskLogin: () => t.invoke('ai:gsk-login'),
    webSearch: (query: string, maxResults?: number) => t.invoke('ai:web-search', query, maxResults),
    imageSearch: (query: string, maxResults?: number) =>
      t.invoke('ai:image-search', query, maxResults),
    fetchImage: (url: string) => t.invoke('ai:fetch-image', url),
    aiGenerateImage: (op: { prompt: string; aspectRatio?: string }) =>
      t.invoke('docs:ai-generate-image', op),
    pickAttachments: overrides.pickAttachments ?? (() => t.invoke('files:pick')),
    addAttachmentPaths: (paths: string[]) => t.invoke('files:add', paths),
    addPastedImage: (data: ArrayBuffer, ext: string) =>
      t.invoke('files:add-pasted-image', data, ext),
    copyImageToClipboard: overrides.copyImageToClipboard ??
      ((dataUrl: string, metaJson?: string) =>
        t.invoke('docs:copy-image-to-clipboard', dataUrl, metaJson)),
    readAttachment: (path: string, offset: number, maxChars: number) =>
      t.invoke('files:read', path, offset, maxChars),
    readAttachmentImage: (path: string) => t.invoke('files:read-image', path),
    getPathForFile: (file: File) => {
      const resolve = overrides.getPathForFile
      if (!resolve) {
        throw new Error(
          `WEB_UNSUPPORTED: resolving dropped files needs the desktop file picker and is unavailable in the web version`,
        )
      }
      return resolve(file)
    },
    openNewTab: overrides.openNewTab ?? ((openPath?: string | null) => t.invoke('win:new', openPath ?? null)),
    listDocsTabs: overrides.listDocsTabs ?? (() => t.invoke('win:list')),
    focusDocsTab: overrides.focusDocsTab ?? ((id: string) => t.invoke('win:focus', id)),
    onAiStream: (handler: (chunk: AiStreamChunk) => void) =>
      t.on('ai:stream-chunk', (chunk) => handler(chunk as AiStreamChunk)),
    onMenuCommand: (handler: (command: MenuCommand, payload?: string) => void) =>
      t.on('menu:command', (command, payload) =>
        handler(command as MenuCommand, payload as string | undefined),
      ),
    onCloseCheck: (handler: () => void) => t.on('docs:close-check', () => handler()),
    reportViewMenuState: (state: { aiSidebar: boolean; darkCanvas: boolean }) =>
      t.send('docs:view-menu-state', {
        aiSidebar: state?.aiSidebar === true,
        darkCanvas: state?.darkCanvas === true,
      }),
    reportCloseCheck: (state: { dirty: boolean; autoSave: boolean; filePath?: string | null }) =>
      t.send('docs:close-check-result', {
        dirty: state?.dirty === true,
        autoSave: state?.autoSave === true,
        filePath: typeof state?.filePath === 'string' ? state.filePath : null,
      }),
    onCloseSaveRequest: (handler: () => void) => t.on('docs:close-save-request', () => handler()),
    reportCloseSaveResult: (ok: boolean) => t.send('docs:close-save-result', ok === true),
  }
  return api
}

/** Chat/project persistence over the shared project:* handlers. */
export function createProjectApi(t: IpcTransport): ProjectApi {
  const projectApi: ProjectApi = {
    resolveChat: (args) => t.invoke('project:resolveChat', args),
    appendChat: (args) => t.invoke('project:appendChat', args),
    loadChat: (args) => t.invoke('project:loadChat', args),
    rebindChat: (args) => t.invoke('project:rebindChat', args),
    // P1 extensions
    listProjects: () => t.invoke('project:list'),
    createProject: (args) => t.invoke('project:create', args),
    renameProject: (args) => t.invoke('project:rename', args),
    deleteProject: (args) => t.invoke('project:delete', args),
    moveFile: (args) => t.invoke('project:moveFile', args),
    getTimeline: (args) => t.invoke('project:timeline', args),
  }
  return projectApi
}
