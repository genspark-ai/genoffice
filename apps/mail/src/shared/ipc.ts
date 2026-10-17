import type { Lang } from '@genoffice/i18n'
import type { AiPanelPrefs } from '@genoffice/ui'
import type {
  AiSettings,
  AiStreamChunk,
  AiStreamRequest,
  GenSparkAccountStatus,
} from '@genoffice/ai-provider'

/**
 * CANONICAL mail channel table (Lane S owns this file — Lane R/C consume,
 * never edit). Names are locked by the P3 app-shell spec amendment; zod
 * schemas for each payload live in `./ipc-schemas` (sheets precedent: this
 * file stays zod-free so the preload can import it).
 *
 * Lane C's compose-session extras (beyond this table) live in a separate
 * `compose-channels.ts` (Lane C owns); the table rows for compose-send /
 * save-draft / outbox-status stay here as the canonical names. The outbox
 * intent-recording stub shape itself is Lane C's call.
 */
export const MAIL_CHANNELS = {
  openMessage: 'mail:open-message',
  listAccounts: 'mail:list-accounts',
  addAccount: 'mail:add-account',
  removeAccount: 'mail:remove-account',
  listFolders: 'mail:list-folders',
  listMessages: 'mail:list-messages',
  fetchBody: 'mail:fetch-body',
  searchLocal: 'mail:search-local',
  searchServer: 'mail:search-server',
  composeSend: 'mail:compose-send',
  saveDraft: 'mail:save-draft',
  outboxStatus: 'mail:outbox-status',
  move: 'mail:move',
  flags: 'mail:flags',
  delete: 'mail:delete',
  syncEvents: 'mail:sync-events',
  attachmentFetch: 'mail:attachment-fetch',
  attachmentProgress: 'mail:attachment-progress',
  dirtyChanged: 'mail:dirty-changed',
  closeSaveRequest: 'mail:close-save-request',
  closeSaveResult: 'mail:close-save-result',
  // P4 sync/outbox additions (Y owns)
  syncNow: 'mail:sync-now',
  outboxList: 'mail:outbox-list',
  outboxRetry: 'mail:outbox-retry',
  outboxCancel: 'mail:outbox-cancel',
  offlineStatus: 'mail:offline-status',
  getLanguage: 'app:get-language',
  languageChanged: 'app:language-changed',
  getTheme: 'app:get-theme',
  themeChanged: 'app:theme-changed',
  // P5 AI retrieval (W owns this entry only): the sole new mail:ai-*
  // channel. Streaming reuses the generic AI_CHANNELS entries above.
  aiRetrieve: 'mail:ai-retrieve',
} as const

export type MailChannel = (typeof MAIL_CHANNELS)[keyof typeof MAIL_CHANNELS]

export type UiTheme = 'light' | 'dark' | 'system'

export interface MailAccountSummary {
  accountId: string
  label: string
  /** sessions only on Linux without a keyring; never a secret here */
  sessionOnly: boolean
}

export interface MessagePage {
  envelopes: import('@genoffice/mail-protocols').MailEnvelope[]
  /** opaque cursor for the next page; absent when fully loaded */
  cursor?: string
}

export interface OutboxStatus {
  queued: number
  failed: number
}

/** Outbox entry as the renderer sees it (P4 spec section 5.1). */
export interface OutboxRecord {
  id: string
  accountId: string
  req: import('@genoffice/mail-protocols').SendRequest
  state: 'queued' | 'sending' | 'failed' | 'sent' | 'cancelled'
  attempts: number
  nextRetryAt?: string
  lastError?: string
  queuedAt: string
  sentAt?: string
}

/** Offline status event payload (P4 spec section 6). */
export interface OfflineStatus {
  online: boolean
  accountId?: string
  detail?: string
}

/** API exposed by the preload to the renderer (window.mailApi). */
export interface MailApi {
  /** Open (or focus) the mail tab for a mailRef; resolves the tab id. */
  openMessage(mailRef: string): Promise<string | null>
  listAccounts(): Promise<MailAccountSummary[]>
  addAccount(label: string): Promise<MailAccountSummary>
  removeAccount(accountId: string): Promise<void>
  listFolders(accountId: string): Promise<import('@genoffice/mail-protocols').MailFolder[]>
  listMessages(
    accountId: string,
    folderId: string,
    limit: number,
    cursor?: string,
  ): Promise<MessagePage>
  fetchBody(
    ref: import('@genoffice/mail-protocols').MailMessageRef,
  ): Promise<import('@genoffice/mail-protocols').MailBody>
  searchLocal(
    accountIds: string[],
    query: import('@genoffice/mail-protocols').MailQuery,
  ): Promise<import('@genoffice/mail-protocols').MailEnvelope[]>
  searchServer(
    accountId: string,
    folderId: string,
    query: import('@genoffice/mail-protocols').MailQuery,
  ): Promise<import('@genoffice/mail-protocols').MailEnvelope[]>
  composeSend(req: import('@genoffice/mail-protocols').SendRequest): Promise<{ outboxId: string }>
  saveDraft(accountId: string, threadKey: string, text: string): Promise<void>
  outboxStatus(accountId: string): Promise<OutboxStatus>
  move(
    ref: import('@genoffice/mail-protocols').MailMessageRef,
    toFolderId: string,
    expectedModseq?: string,
  ): Promise<void>
  setFlags(
    ref: import('@genoffice/mail-protocols').MailMessageRef,
    flags: Partial<import('@genoffice/mail-protocols').MailFlags>,
    expectedModseq?: string,
  ): Promise<void>
  deleteMessage(
    ref: import('@genoffice/mail-protocols').MailMessageRef,
    expectedModseq?: string,
  ): Promise<void>
  /** Push events from main (refs/flags only, never bodies or secrets). */
  onSyncEvent(handler: (event: import('@genoffice/mail-protocols').MailEvent) => void): () => void
  /** Trigger an incremental sync step; false when one already runs or offline. */
  syncNow(accountId: string, folderId?: string): Promise<{ started: boolean }>
  /** Outbox entries newest-first (capped at 200). */
  outboxList(accountId: string): Promise<OutboxRecord[]>
  /** Manual retry: reset the budget and re-queue at head. */
  outboxRetry(accountId: string, outboxId: string): Promise<{ outboxId: string; state: string }>
  /** Cancel a queued or failed entry; false when already sent/sending. */
  outboxCancel(accountId: string, outboxId: string): Promise<{ cancelled: boolean }>
  /** Offline status; replays the current value on subscribe. */
  onOfflineStatus(handler: (status: OfflineStatus) => void): () => void
  attachmentFetch(attachmentPath: string): Promise<{ ok: boolean; path?: string }>
  onAttachmentProgress(handler: (completed: number, total: number) => void): () => void
  /** Mirror unsaved-changes state to the main process; drives the draft prompt before closing a tab */
  setDirty(dirty: boolean): void
  /** Main process picked "Save draft" in the close prompt → renderer saves and replies via sendCloseSaveResult */
  onCloseSaveRequest(handler: () => void): () => void
  sendCloseSaveResult(ok: boolean): void
  getLanguage(): Promise<Lang>
  onLanguageChanged(handler: (lang: Lang) => void): () => void
  getTheme(): Promise<UiTheme>
  onThemeChanged(handler: (theme: UiTheme) => void): () => void
  getAiPanelPrefs(): Promise<AiPanelPrefs>
  onAiPanelPrefsChanged(handler: (prefs: AiPanelPrefs) => void): () => void
  /**
   * Fire-and-forget: a pointerdown landed on the shell chrome (tab strip is a
   * sibling WebContentsView whose clicks produce no DOM event here).
   */
  onChromePressed(handler: () => void): () => void
  getAiSettings(): Promise<AiSettings>
  aiGskStatus(): Promise<GenSparkAccountStatus>
  aiStream(request: AiStreamRequest): Promise<void>
  aiStreamCancel(requestId: string): Promise<void>
  onAiStream(handler: (chunk: AiStreamChunk) => void): () => void
  /**
   * P5 scoped retrieval (W owns). Optional until the UI lane lands the
   * FakeMailApi stub: the preload bridge implements it today, renderer
   * fakes follow with the AiPanel lane.
   */
  aiRetrieve?(request: AiRetrieveRequest): Promise<AiRetrieveResponse>
}

/**
 * P5 AI retrieval wire shapes (W owns this block only). The picker-selected
 * account set travels with every call; main rejects unknown accounts and
 * opens only per-account stores, while the renderer discards results whose
 * echoed scope no longer matches the live picker.
 */
export interface AiScope {
  /** Picker-selected account ids; non-empty, max 8. */
  accountIds: string[]
}

export type AiRetrieveOp = 'triage-headers' | 'fetch-bodies' | 'cite-resolve'

export interface AiRetrieveRequest {
  scope: AiScope
  op: AiRetrieveOp
  /** triage-headers: FTS query over headers + snippet. */
  query?: string
  folderId?: string
  /** Default 50, hard cap 200. */
  limit?: number
  /**
   * fetch-bodies / cite-resolve: explicit user-cited set, max 50. Each ref
   * may carry the optional exact-match accountId hint (H1 D-H1-C3): the
   * citation's scope account, resolved exactly there with loud failure on
   * stale hints; absent means deterministic first-scope-match fallback.
   */
  refs?: import('@genoffice/mail-protocols').MailMessageRef[]
}

export interface AiBodyEntry {
  ref: import('@genoffice/mail-protocols').MailMessageRef
  /** Scope account that held this ref (refs carry no account id). */
  accountId: string
  /** Cited body text, redacted for prompts. */
  text: string
}

export interface AiCitation {
  uid: string
  folderId: string
  msgIdHeader?: string
  /** Scope account that held this ref (refs carry no account id). */
  accountId: string
}

export interface AiRetrieveResponse {
  /** Echo of the request scope; the renderer discards on mismatch. */
  scope: AiScope
  /** Headers + snippet only; bodies never appear here. */
  envelopes: import('@genoffice/mail-protocols').MailEnvelope[]
  /** fetch-bodies only, cited set only, redacted for prompts. */
  bodies?: AiBodyEntry[]
  citations: AiCitation[]
}

/** AI channels are app-wide shared ipcMain handlers (shell registers via docs-main registerAiIpc); pass-through only */
export const AI_CHANNELS = {
  getSettings: 'ai:get-settings',
  gskStatus: 'ai:gsk-status',
  stream: 'ai:stream',
  streamChunk: 'ai:stream-chunk',
  streamCancel: 'ai:stream-cancel',
  webSearch: 'ai:web-search',
  imageSearch: 'ai:image-search',
  fetchImage: 'ai:fetch-image',
} as const
