export interface OpenFileResult {
  path: string
  name: string
  /** raw docx bytes */
  data: ArrayBuffer
  /** sha256 of the original file; original archived under this hash */
  hash: string
}

export interface PickImageResult {
  /** raw image bytes, base64 encoded */
  base64: string
  mime: 'image/png' | 'image/jpeg' | 'image/gif'
  name: string
}

// ---- AI provider settings/config/streaming: canonical types live in @genoffice/ai-provider ----

import type {
  AiChatRequest,
  AiChatResponse,
  AiSettings,
  AiStreamChunk,
  AiStreamRequest,
  GenSparkAccountStatus,
} from '@genoffice/ai-provider'
import type { FaceVerticalMetrics } from '@genoffice/font-metrics'

export type { FaceVerticalMetrics }

export type {
  AiChatRequest,
  AiChatResponse,
  AiProviderConfig,
  AiProviderId,
  AiProviderMeta,
  AiSettings,
  AiStreamChunk,
  AiStreamRequest,
  GenSparkAccountStatus,
} from '@genoffice/ai-provider'
export { AI_PROVIDERS } from '@genoffice/ai-provider'

// ---- agent protocol: canonical types live in @genoffice/agent-core ----

export type {
  AgentMessage,
  AgentToolCall,
  AgentToolDef,
  AgentToolResult,
} from '@genoffice/agent-core'

// ---- chat attachments (local files fed to the agent via tools) ----

/** Image attachment extensions: no text extraction; read as base64 on send and passed to the model as a multimodal image with the user message */
export const ATTACHMENT_IMAGE_EXTS = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp'])

export interface AttachmentMeta {
  /** absolute local path; the file never leaves the machine */
  path: string
  name: string
  /** lowercased extension without the dot */
  ext: string
  sizeBytes: number
}

export interface AttachmentAddResult {
  accepted: AttachmentMeta[]
  /** per-file rejection messages (too large / unsupported type / unreadable) */
  rejected: string[]
}

export interface AttachmentReadResult {
  ok: boolean
  error?: string
  name?: string
  /** total characters of the extracted text */
  totalChars?: number
  /** requested slice */
  text?: string
  offset?: number
}

/** an image attachment read as raw bytes for multimodal input (files:read-image) */
export interface AttachmentImageResult {
  ok: boolean
  /** raw base64 (no data: URL prefix) */
  base64?: string
  mime?: string
  error?: string
}

// ---- Google Docs integration ----

export interface GoogleAuthStatus {
  /** false when no client id/secret could be resolved (env/Keychain/config file) */
  configured: boolean
  signedIn: boolean
}

export interface GoogleDocSummary {
  id: string
  name: string
  modifiedTime: string
  webViewLink: string
  iconLink?: string
}

export type GoogleRole = 'reader' | 'commenter' | 'writer'

export interface GooglePermissionSummary {
  id: string
  type: 'user' | 'anyone' | 'group' | 'domain'
  role: GoogleRole | 'owner'
  emailAddress?: string
  displayName?: string
}

export type GoogleApiResult<T> = { ok: true; data: T } | { ok: false; error: string }

export interface GoogleFolderSummary {
  id: string
  name: string
}

export interface GoogleSettings {
  /** null = send new docs to My Drive root */
  defaultFolderId: string | null
  defaultFolderName: string | null
}

/** an open docs tab, for View → Switch Tab */
export interface DocsTabInfo {
  id: string
  title: string
  focused: boolean
}

/** commands dispatched from the native application menu to the renderer */
export type MenuCommand =
  | 'new'
  | 'open'
  | 'open-path'
  | 'save'
  | 'save-as'
  | 'undo'
  | 'redo'
  | 'zoom-in'
  | 'zoom-out'
  | 'zoom-100'
  | 'zoom-page-width'
  | 'zoom-whole-page'
  | 'toggle-ai'
  | 'toggle-dark'
  | 'insert-table'
  | 'insert-image'
  | 'insert-page-break'
  | 'insert-link'
  | 'insert-equation'
  | 'insert-comment'
  | 'font-dialog'
  | 'paragraph-dialog'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'align-left'
  | 'align-center'
  | 'align-right'
  | 'align-justify'
  | 'page-setup'
  | 'find'
  | 'print'
  | 'export-pdf'
  | 'word-count'

export interface DesktopApi {
  /** current UI language (persisted by the shell in app-settings.json) */
  getLanguage(): Promise<'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar'>
  /** language switched from the shell home page */
  onLanguageChanged(
    handler: (
      lang: 'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar',
    ) => void,
  ): () => void
  /** current theme setting from the shell */
  getTheme(): Promise<'light' | 'dark' | 'system'>
  /** theme changed from the shell home page */
  onThemeChanged(handler: (theme: string) => void): () => void
  openDocx(): Promise<OpenFileResult | null>
  openDocxPath(path: string): Promise<OpenFileResult | null>
  /** mark the renderer ready and consume a file passed by Finder/Explorer at launch */
  consumePendingOpenDocx(): Promise<OpenFileResult | null>
  /** returns true when this tab was created via "New Document" and should start blank */
  consumeNewBlankDoc(): Promise<boolean>
  /** receive documents opened from Finder/Explorer while the app is running */
  onOpenDocx(handler: (result: OpenFileResult) => void): () => void
  /** File was renamed externally (renamed in the shell Home list) — pushes old and new paths; renderer syncs its save path and title bar */
  onRenamedDocx(handler: (paths: { oldPath: string; newPath: string }) => void): () => void
  /** auto=true marks an autosave: an externally modified file then fails with
   *  reason 'external-modified' instead of prompting (manual saves get an
   *  Overwrite/Cancel dialog in the main process) */
  saveDocx(
    path: string,
    data: ArrayBuffer,
    auto?: boolean,
  ): Promise<{ ok: boolean; error?: string; reason?: 'external-modified' }>
  /** crash-recovery copy of a dirty document, stored under userData */
  writeRecoveryCopy(path: string, data: ArrayBuffer): Promise<{ ok: boolean }>
  /** tab closed but webContents kept alive (shell freeze workaround) — stop background timers */
  onTeardown(handler: () => void): () => void
  saveDocxAs(
    defaultName: string,
    data: ArrayBuffer,
  ): Promise<{ ok: boolean; path?: string; error?: string }>
  /** first save of a new document: silently writes into the default folder, no dialog */
  saveDocxNew(
    defaultName: string,
    data: ArrayBuffer,
  ): Promise<{ ok: boolean; path?: string; error?: string }>
  getRecentFiles(): Promise<string[]>
  pickImage(): Promise<PickImageResult | null>
  /** vertical metrics of an installed family (exact name match), null when missing */
  fontMetrics(family: string): Promise<FaceVerticalMetrics | null>
  getAiSettings(): Promise<AiSettings>
  setAiSettings(settings: AiSettings): Promise<void>
  /** system print dialog for the current window */
  print(): Promise<void>
  /** render the document to PDF and ask where to save; size in twips.
   *  outPath is only honored when a previous export dialog chose that exact path */
  exportPdf(
    defaultName: string,
    pageWidthTwips: number,
    pageHeightTwips: number,
    outPath?: string,
  ): Promise<{ ok: boolean; path?: string; error?: string }>
  /** Mixed paper-size export: produce a set of PDF bytes (base64) at given sizes per the current print layout */
  printPdfBuffer(
    pageWidthTwips: number,
    pageHeightTwips: number,
  ): Promise<{ ok: boolean; base64?: string; error?: string }>
  /** Merge grouped PDF fragments in order and write to disk (missing outPath opens
   *  the save dialog; a given outPath must come from a previous export dialog) */
  saveMergedPdf(
    defaultName: string,
    base64Parts: string[],
    outPath?: string,
  ): Promise<{ ok: boolean; path?: string; error?: string }>
  aiChat(request: AiChatRequest): Promise<AiChatResponse>
  /** start a streaming AI call; deltas arrive via onAiStream with the same requestId */
  aiStream(request: AiStreamRequest): Promise<void>
  aiStreamCancel(requestId: string): Promise<void>
  /** Genspark account status (gsk login state); withEmail also returns the email (needs a network request, slower) */
  aiGskStatus(withEmail?: boolean): Promise<GenSparkAccountStatus>
  /** Open the browser to log in to Genspark (fire-and-forget; aiGskStatus flips to logged-in when done) */
  aiGskLogin(): Promise<void>
  webSearch(
    query: string,
    maxResults?: number,
  ): Promise<{
    results: Array<{ title: string; url: string; snippet: string }>
    answer?: string
    method: string
    /** failure reason when method === 'error' */
    error?: string
  }>
  imageSearch(
    query: string,
    maxResults?: number,
  ): Promise<{
    images: Array<{
      title: string
      imageUrl: string
      sourceUrl: string
      source: string
      width?: number
      height?: number
    }>
    method: string
    /** failure reason when method === 'error' */
    error?: string
  }>
  fetchImage(url: string): Promise<{ base64: string; mime: string } | null>
  /** file picker for chat attachments (multi-select) */
  pickAttachments(): Promise<AttachmentAddResult | null>
  /** validate dropped paths and return attachment metadata */
  addAttachmentPaths(paths: string[]): Promise<AttachmentAddResult>
  /** persist a pasted clipboard image (no local path) to a temp file and add it as an attachment */
  addPastedImage(data: ArrayBuffer, ext: string): Promise<AttachmentAddResult>
  /** read a slice of the extracted text of an attachment */
  readAttachment(path: string, offset: number, maxChars: number): Promise<AttachmentReadResult>
  /** read an image attachment as base64 for multimodal input (≤5MB) */
  readAttachmentImage(path: string): Promise<AttachmentImageResult>
  /** absolute path of a File dropped onto the window (Electron webUtils) */
  getPathForFile(file: File): string
  /** View → New Tab: open another docs tab, optionally loading the same document */
  openNewTab(openPath?: string | null): Promise<void>
  /** all open docs tabs, for View → Switch Tab */
  listDocsTabs(): Promise<DocsTabInfo[]>
  focusDocsTab(id: string): Promise<void>
  /** subscribe to AI stream chunks; returns unsubscribe */
  onAiStream(handler: (chunk: AiStreamChunk) => void): () => void
  /** subscribe to native menu commands; returns unsubscribe */
  onMenuCommand(handler: (command: MenuCommand, payload?: string) => void): () => void
  /** Close guard: main process queries pre-close state (dirty flag + autosave switch; if autosave is on, save silently without a dialog) */
  onCloseCheck(handler: () => void): () => void
  reportCloseCheck(state: { dirty: boolean; autoSave: boolean; filePath?: string | null }): void
  /** Close guard chose "Save": main process asks the renderer to run the full save flow */
  onCloseSaveRequest(handler: () => void): () => void
  reportCloseSaveResult(ok: boolean): void
  /** Report view menu state to the shell (for menu checkbox state updates) */
  reportViewMenuState(state: { aiSidebar: boolean; darkCanvas: boolean }): void

  // ---- Google Docs integration ----

  /** configured = a client id/secret was resolved; signedIn = a refresh token is stored */
  googleAuthStatus(): Promise<GoogleAuthStatus>
  /** opens the system browser for the loopback+PKCE consent flow */
  googleSignIn(): Promise<{ ok: boolean; error?: string; unconfigured?: boolean }>
  googleSignOut(): Promise<void>
  /** the user's Google Docs, newest-modified first, for the import picker */
  googleListDocs(): Promise<GoogleApiResult<GoogleDocSummary[]>>
  /** exports the Google Doc to .docx and opens it through the normal open pipeline */
  googleImportDoc(
    fileId: string,
  ): Promise<
    GoogleApiResult<OpenFileResult & { googleFileId: string; googleWebViewLink: string | null }>
  >
  /** webViewLink + name lookup for a fileId already known this session */
  googleGetFileMeta(
    fileId: string,
  ): Promise<GoogleApiResult<{ id: string; name: string; webViewLink: string }>>
  /** uploads .docx bytes as a new native Google Doc; lands in the default
   *  destination folder when one is set (see googleGetSettings) */
  googleCreateDoc(
    name: string,
    data: ArrayBuffer,
  ): Promise<
    GoogleApiResult<{
      id: string
      webViewLink: string
      /** name of the default folder it was created in, if any */
      folderName?: string
      /** true when the stored default folder was stale (404) and got cleared */
      folderCleared?: boolean
    }>
  >
  /** replaces the content of an existing Google Doc (previously sent/imported).
   *  `fallbackName` is used if the app never got write access to `fileId` (e.g.
   *  the doc was imported via the picker, not created by this app under
   *  drive.file scope) — in that case a new Google Doc is created instead, and
   *  `createdNew: true` is returned so the caller can re-point googleFileId. */
  googleUpdateDoc(
    fileId: string,
    data: ArrayBuffer,
    fallbackName: string,
  ): Promise<GoogleApiResult<{ id: string; webViewLink: string; createdNew?: boolean }>>
  googleListPermissions(fileId: string): Promise<GoogleApiResult<GooglePermissionSummary[]>>
  googleAddPermission(
    fileId: string,
    emailAddress: string,
    role: GoogleRole,
  ): Promise<GoogleApiResult<GooglePermissionSummary>>
  googleUpdatePermission(
    fileId: string,
    permissionId: string,
    role: GoogleRole,
  ): Promise<GoogleApiResult<GooglePermissionSummary>>
  googleRemovePermission(fileId: string, permissionId: string): Promise<GoogleApiResult<null>>
  /** general access ("anyone with the link"); role=null revokes it */
  googleSetAnyoneAccess(
    fileId: string,
    role: 'reader' | 'writer' | null,
  ): Promise<GoogleApiResult<null>>
  googleOpenExternal(url: string): void

  /** default destination folder for "Send to Google Docs" */
  googleGetSettings(): Promise<GoogleApiResult<GoogleSettings>>
  googleSetSettings(settings: GoogleSettings): Promise<GoogleApiResult<null>>
  /** folders under parentId (omit/undefined = My Drive root), for the folder picker */
  googleListFolders(parentId?: string): Promise<GoogleApiResult<GoogleFolderSummary[]>>
  /** move an existing Drive file to a different folder */
  googleMoveFile(
    fileId: string,
    folderId: string,
  ): Promise<GoogleApiResult<{ id: string; parents: string[] }>>
  /** duplicate an existing Drive file (Make a copy); folderId places the copy
   *  directly in a Drive folder instead of wherever the source file lives */
  googleCopyFile(
    fileId: string,
    name: string,
    folderId?: string,
  ): Promise<GoogleApiResult<{ id: string; name: string; webViewLink: string }>>
  /** rename an existing Drive file (GDocsHeader title commit, once the doc is
   *  linked to Google and app-writable) */
  googleRenameFile(
    fileId: string,
    name: string,
  ): Promise<GoogleApiResult<{ id: string; name: string }>>
}
