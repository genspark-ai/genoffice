import type { Lang } from '@genoffice/i18n'
import type {
  AiSettings,
  AiStreamChunk,
  AiStreamRequest,
  GenSparkAccountStatus,
} from '@genoffice/ai-provider'

export const HWP_CHANNELS = {
  consumePending: 'hwp:consume-pending',
  readFile: 'hwp:read-file',
  save: 'hwp:save',
  saveRequest: 'hwp:save-request',
  saveRequestAck: 'hwp:save-request-ack',
  dirtyChanged: 'hwp:dirty-changed',
  closeSaveRequest: 'hwp:close-save-request',
  closeSaveResult: 'hwp:close-save-result',
  fileRenamed: 'hwp:file-renamed',
  getLanguage: 'app:get-language',
  languageChanged: 'app:language-changed',
  getTheme: 'app:get-theme',
  themeChanged: 'app:theme-changed',
} as const

/** AI channels are app-wide shared ipcMain handlers (shell registers them); pass-through only */
export const AI_CHANNELS = {
  getSettings: 'ai:get-settings',
  stream: 'ai:stream',
  streamChunk: 'ai:stream-chunk',
  streamCancel: 'ai:stream-cancel',
  gskStatus: 'ai:gsk-status',
  gskLogin: 'ai:gsk-login',
} as const

export type SaveMode = 'save' | 'saveAs'

export interface SaveHwpRequest {
  mode: SaveMode
  hwp: Uint8Array
  hwpx: Uint8Array
  hml?: Uint8Array
}

export type SaveHwpResult =
  { ok: true; path: string } | { ok: true; canceled: true } | { ok: false; error: string }

export type UiTheme = 'light' | 'dark' | 'system'

export interface HwpApi {
  consumePending(): Promise<string | null>
  /** granted path only; raw Hangul file bytes */
  readFile(path: string): Promise<Uint8Array>
  save(request: SaveHwpRequest): Promise<SaveHwpResult>
  setDirty(dirty: boolean): void
  onSaveRequest(handler: (mode: SaveMode) => void): () => void
  onCloseSaveRequest(handler: () => void): () => void
  sendCloseSaveResult(ok: boolean): void
  sendSaveRequestAck(ok: boolean): void
  onFileRenamed(handler: (newPath: string) => void): () => void
  getLanguage(): Promise<Lang>
  onLanguageChanged(handler: (lang: Lang) => void): () => void
  getTheme(): Promise<UiTheme>
  onThemeChanged(handler: (theme: UiTheme) => void): () => void
  getAiSettings(): Promise<AiSettings>
  aiStream(request: AiStreamRequest): Promise<void>
  aiStreamCancel(requestId: string): Promise<void>
  onAiStream(handler: (chunk: AiStreamChunk) => void): () => void
  aiGskStatus(withEmail?: boolean): Promise<GenSparkAccountStatus>
  aiGskLogin(): Promise<void>
}
