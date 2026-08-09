import type { Lang } from '@genoffice/i18n'
import type { AiSettings, AiStreamChunk, AiStreamRequest } from '@genoffice/ai-provider'

export const MARKDOWN_CHANNELS = {
  consumePending: 'markdown:consume-pending',
  readFile: 'markdown:read-file',
  save: 'markdown:save',
  saveRequest: 'markdown:save-request',
  saveRequestAck: 'markdown:save-request-ack',
  dirtyChanged: 'markdown:dirty-changed',
  closeSaveRequest: 'markdown:close-save-request',
  closeSaveResult: 'markdown:close-save-result',
  fileRenamed: 'markdown:file-renamed',
  pickImage: 'markdown:pick-image',
  saveImage: 'markdown:save-image',
  readImage: 'markdown:read-image',
  exportRequest: 'markdown:export-request',
  exportDocx: 'markdown:export-docx',
  exportPdf: 'markdown:export-pdf',
  getLanguage: 'app:get-language',
  languageChanged: 'app:language-changed',
  getTheme: 'app:get-theme',
  themeChanged: 'app:theme-changed',
} as const

export type UiTheme = 'light' | 'dark' | 'system'

export type SaveMode = 'save' | 'saveAs'

export interface SaveMarkdownRequest {
  /** full document text (frontmatter included) */
  text: string
  mode: SaveMode
  /**
   * Silent first save for an untitled document (AI auto-naming): saves to a
   * unique path under Documents derived from this name, without a dialog.
   * Ignored when the document already has a path.
   */
  suggestedName?: string
}

export type SaveMarkdownResult =
  { ok: true; path: string } | { ok: true; canceled: true } | { ok: false; error: string }

/** AI channels are app-wide shared ipcMain handlers (shell registers via docs-main registerAiIpc); pass-through only */
export const AI_CHANNELS = {
  getSettings: 'ai:get-settings',
  stream: 'ai:stream',
  streamChunk: 'ai:stream-chunk',
  streamCancel: 'ai:stream-cancel',
  webSearch: 'ai:web-search',
} as const

export interface WebSearchResult {
  answer?: string
  results: Array<{ title: string; url: string; snippet: string }>
}

export type ExportFormat = 'pdf' | 'docx' | 'docs'

export interface ExportDocxRequest {
  /** .docx bytes, base64 */
  base64: string
  /** file name (no extension) suggested in the dialog / used for the silent convert */
  suggestedName: string
  /** 'dialog' = save dialog; 'openInDocs' = silent save next to the .md, then open in AI Docs */
  mode: 'dialog' | 'openInDocs'
}

export interface ExportPdfRequest {
  /** self-contained print HTML */
  html: string
  suggestedName: string
}

export type ExportResult =
  { ok: true; path: string } | { ok: true; canceled: true } | { ok: false; error: string }

export interface ImageData {
  base64: string
  mime: 'image/png' | 'image/jpeg' | 'image/gif'
}

/** API exposed by preload to the renderer (window.markdownApi) */
export interface MarkdownApi {
  /** Take the md path pending for this view (queued at tab creation); null = new untitled document */
  consumePending(): Promise<string | null>
  /** Read the file as UTF-8 text. Only paths granted to this view are allowed */
  readFile(path: string): Promise<string>
  /**
   * Write the document text. With a granted file path the write is atomic
   * (tmp + rename); untitled documents and mode 'saveAs' go through a main-process
   * save dialog first. The resolved path is granted to the view and returned.
   */
  save(request: SaveMarkdownRequest): Promise<SaveMarkdownResult>
  /** Mirror unsaved-changes state to the main process; drives the save prompt before closing a tab/window */
  setDirty(dirty: boolean): void
  /** Shell menu Save / Save As → renderer serializes and calls save() with the given mode */
  onSaveRequest(handler: (mode: SaveMode) => void): () => void
  /** Resolves a menu-save waiter when doSave exits without ever invoking save() (busy/loading) */
  sendSaveRequestAck(ok: boolean): void
  /** Main process picked "Save" in the close prompt → renderer saves and replies via sendCloseSaveResult */
  onCloseSaveRequest(handler: () => void): () => void
  sendCloseSaveResult(ok: boolean): void
  /** The file was renamed on disk (Home list rename) — renderer syncs its display path */
  onFileRenamed(handler: (newPath: string) => void): () => void
  /**
   * Pick an image file and copy it into `assets/` next to the open document;
   * returns the relative path to author into the markdown, or null when the
   * document is untitled or the picker was canceled.
   */
  pickImage(): Promise<string | null>
  /**
   * Persist pasted/dropped image bytes into `assets/` next to the open
   * document; returns the relative path to author, or null when untitled.
   */
  saveImage(data: { base64: string; ext: string }): Promise<string | null>
  /**
   * Read an image referenced by the document for DOCX embedding. Only paths
   * inside the document's directory are allowed; anything else returns null.
   */
  readImage(src: string): Promise<ImageData | null>
  /** Shell menu export → renderer serializes and calls exportDocx/exportPdf */
  onExportRequest(handler: (format: ExportFormat) => void): () => void
  exportDocx(request: ExportDocxRequest): Promise<ExportResult>
  exportPdf(request: ExportPdfRequest): Promise<ExportResult>
  getLanguage(): Promise<Lang>
  onLanguageChanged(handler: (lang: Lang) => void): () => void
  getTheme(): Promise<UiTheme>
  onThemeChanged(handler: (theme: UiTheme) => void): () => void
  getAiSettings(): Promise<AiSettings>
  aiStream(request: AiStreamRequest): Promise<void>
  aiStreamCancel(requestId: string): Promise<void>
  onAiStream(handler: (chunk: AiStreamChunk) => void): () => void
  /** Main-process web search (Serper/DuckDuckGo via the shared ai:web-search handler) */
  webSearch(query: string, maxResults?: number): Promise<WebSearchResult>
}
